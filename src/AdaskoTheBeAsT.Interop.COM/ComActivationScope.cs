using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1121 // Use built-in type alias
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
internal sealed class ComActivationScope : IComActivationScope
{
    private readonly Thread _creatingThread = Thread.CurrentThread;
    private readonly uint _creatingNativeThreadId = NativeMethods.GetCurrentThreadId();
    private readonly List<string> _referencedDllPaths;

    internal ComActivationScope(
        IReadOnlyList<IntPtr> activationContextHandles,
        IReadOnlyList<IntPtr> activationCookies,
        IReadOnlyList<ActivationContextApi.ModuleLease> moduleLeases,
        IReadOnlyList<string> referencedDllPaths,
        ComActivationOptions options,
        ActivationContextApi? activationContextApi = null)
    {
        ThrowHelper.ThrowIfNull(activationContextHandles, nameof(activationContextHandles));
        ThrowHelper.ThrowIfNull(activationCookies, nameof(activationCookies));
        ThrowHelper.ThrowIfNull(moduleLeases, nameof(moduleLeases));
        ThrowHelper.ThrowIfNull(referencedDllPaths, nameof(referencedDllPaths));
        ThrowHelper.ThrowIfNull(options, nameof(options));

        ActivationContextHandles = [.. activationContextHandles];
        ActivationCookies = [.. activationCookies];
        ModuleLeases = [.. moduleLeases];
        _referencedDllPaths = [.. referencedDllPaths];
        Options = options;
        ContextApi = activationContextApi ?? ActivationContextApi.Instance;

        for (var i = 0; i < ActivationContextHandles.Count; i++)
        {
            if (ActivationContextApi.IsFallbackHandle(ActivationContextHandles[i]))
            {
                IsUsingNanoServerFallback = true;
                break;
            }
        }
    }

#pragma warning disable MA0055 // Do not use finalizer (intentional diagnostic safety net)
#pragma warning disable IDISP023 // Don't use reference types in finalizer context (diagnostic emission only)
    ~ComActivationScope()
    {
        if (!IsReleased)
        {
            ComInteropEventSource.Log.ScopeLeaked();
            if (Debugger.IsAttached)
            {
                Debug.Fail(
                    "ComActivationScope was not disposed. "
                    + "Call Dispose() or Executor.Deactivate on the creating thread to release native activation contexts and module leases.");
            }
        }
    }
#pragma warning restore IDISP023
#pragma warning restore MA0055

    public bool IsReleased { get; private set; }

    public bool IsUsingNanoServerFallback { get; }

    public IReadOnlyList<IntPtr> LoadedModuleHandles
    {
        get
        {
            if (IsReleased)
            {
                return [];
            }

            var handles = new List<IntPtr>();
            var seen = new HashSet<IntPtr>();
            for (var i = 0; i < ModuleLeases.Count; i++)
            {
                var h = ModuleLeases[i].Handle;
                if (h != IntPtr.Zero && seen.Add(h))
                {
                    handles.Add(h);
                }
            }

            for (var i = 0; i < ActivationCookies.Count; i++)
            {
                var fallbackHandles = ContextApi.GetFallbackLoadedHandles(ActivationCookies[i]);
                for (var j = 0; j < fallbackHandles.Count; j++)
                {
                    if (seen.Add(fallbackHandles[j]))
                    {
                        handles.Add(fallbackHandles[j]);
                    }
                }
            }

            var pathHandles = ActivationContextApi.GetLoadedHandlesForPaths(_referencedDllPaths);
            for (var i = 0; i < pathHandles.Count; i++)
            {
                if (seen.Add(pathHandles[i]))
                {
                    handles.Add(pathHandles[i]);
                }
            }

            return handles;
        }
    }

    internal List<IntPtr> ActivationContextHandles { get; }

    internal List<IntPtr> ActivationCookies { get; }

    internal List<ActivationContextApi.ModuleLease> ModuleLeases { get; }

    internal List<object> TrackedObjects { get; } = [];

    internal ComActivationOptions Options { get; }

    internal ActivationContextApi ContextApi { get; }

    internal bool IsReleasing { get; set; }

    public void TrackObject(object comObject)
    {
        ThrowHelper.ThrowIfNull(comObject, nameof(comObject));
        ValidateOwnerThread("The COM activation scope can only track objects on the thread that activated it.");
        if (IsReleased || IsReleasing)
        {
            throw new ObjectDisposedException(nameof(ComActivationScope));
        }

        TrackedObjects.Add(comObject);
    }

    public ComCreateResult<object> CreateInstance(Guid clsid, int moduleIndex = 0, bool trackObject = true)
        => ComActivationScopeExtensions.CreateInstanceWithOptionalInterface(this, clsid, interfaceId: null, moduleIndex, trackObject);

    public ComCreateResult<TInterface> CreateInstance<TInterface>(Guid clsid, int moduleIndex = 0, bool trackObject = true)
        where TInterface : class
    {
        if (!TryResolveScopeModuleHandle(moduleIndex, out var moduleHandle, out var error))
        {
            return new ComCreateResult<TInterface>
            {
                Success = false,
                Exception = error,
            };
        }

        var creation = ComClassFactory.CreateInstance<TInterface>(moduleHandle, clsid);
        if (creation.Success && creation.Value is not null && trackObject)
        {
            TrackedObjects.Add(creation.Value);
        }

        return creation;
    }

    public Result PumpPendingMessages(int maxBatch = 64)
    {
        if (maxBatch <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(maxBatch), "Maximum batch size must be greater than zero.");
        }

        var result = new Result { Success = false };
        try
        {
            ValidateOwnerThread("Pending messages must be pumped on the thread that activated the COM scope.");
            if (IsReleased)
            {
                throw new ObjectDisposedException(nameof(ComActivationScope));
            }

            NativeMethods.PumpPendingMessages(maxBatch);
            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }

        return result;
    }

    public void Dispose()
    {
        if (!IsReleased)
        {
            var result = Executor.Deactivate(this);
            if (!result.Success)
            {
                var errorMessage = result.Exception?.Message ?? "unknown error";
                ComInteropEventSource.Log.ScopeReleaseFailed(errorMessage);
                if (Debugger.IsAttached)
                {
                    Debug.Fail("ComActivationScope.Dispose() failed to deactivate cleanly: " + errorMessage);
                }
            }
        }

        if (IsReleased)
        {
            GC.SuppressFinalize(this);
        }
    }

    internal bool TryResolveScopeModuleHandle(int moduleIndex, out IntPtr moduleHandle, out Exception? error)
    {
        moduleHandle = IntPtr.Zero;
        error = null;
        try
        {
            ValidateOwnerThread("COM instances can only be created on the thread that activated the COM scope.");
            if (IsReleased || IsReleasing)
            {
                throw new ObjectDisposedException(nameof(ComActivationScope));
            }

            if (moduleIndex < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(moduleIndex), "Module index cannot be negative.");
            }

            var handles = LoadedModuleHandles;
            if (handles.Count == 0)
            {
                throw new InvalidOperationException(
                    "No loaded module handles are available in this activation scope. "
                    + "Set ComActivationOptions.PreloadModules = true (or ForceInProcessFallback = true) when activating the scope.");
            }

            if (moduleIndex >= handles.Count)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(moduleIndex),
                    "Module index must be less than LoadedModuleHandles.Count (" + handles.Count.ToString(CultureInfo.InvariantCulture) + ").");
            }

            moduleHandle = handles[moduleIndex];
            return true;
        }
        catch (Exception ex)
        {
            error = ex;
            return false;
        }
    }

    internal void ValidateRelease()
    {
        ValidateOwnerThread("The COM activation scope must be released on the thread that created it.");
        if (IsReleasing)
        {
            throw new InvalidOperationException("The COM activation scope is already being released.");
        }
    }

    internal void MarkReleased()
    {
        IsReleased = true;
        TrackedObjects.Clear();
        ActivationContextHandles.Clear();
        ActivationCookies.Clear();
        ModuleLeases.Clear();
#pragma warning disable S3971, CA1816
        GC.SuppressFinalize(this);
#pragma warning restore S3971, CA1816
    }

    private void ValidateOwnerThread(string errorMessage)
    {
        if (!ReferenceEquals(Thread.CurrentThread, _creatingThread)
            || NativeMethods.GetCurrentThreadId() != _creatingNativeThreadId)
        {
            throw new InvalidOperationException(errorMessage);
        }
    }
}
#pragma warning restore SA1121
