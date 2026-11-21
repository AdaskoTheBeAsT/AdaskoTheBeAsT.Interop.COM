using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1121 // Use built-in type alias
/// <summary>
/// Provides helpers for activating registration-free COM components, executing work inside their activation
/// contexts, and explicitly managing longer-lived COM object and multi-coclass scope lifetimes on the current thread.
/// </summary>
/// <remarks>
/// <para>
/// <b>New code should prefer <see cref="IComExecutor"/> / <see cref="IComActivationExecutor"/>
/// (implemented by <see cref="ComExecutor"/>)</b>, which exposes the same surface through an injectable,
/// testable interface. This static class is retained for direct and backward-compatible usage.
/// </para>
/// <para>
/// Work runs synchronously on the calling thread. Callers should provide an STA thread when required by their component.
/// Nested handles and scopes created inside a callback or factory must be released in reverse creation order
/// before that outer scope or delegate completes.
/// </para>
/// </remarks>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public static class Executor
{
    private const int NativeActCtxSizeX86 = 0x20;
    private const int NativeActCtxSizeX64 = 0x38;

    // Track cookies, not COM handles, so this stack does not prevent leak diagnostics from running.
    [ThreadStatic]
    private static List<IntPtr>? _activeCookies;

    /// <summary>
    /// Activates a single registration-free COM context, executes the supplied callback, pumps pending COM
    /// messages, and then releases the activation context.
    /// </summary>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="action">Action to execute within the activation context.</param>
    /// <returns>The outcome of the operation.</returns>
    public static Result Execute(
        string comAssemblyPath,
        string manifestPath,
        Action action)
        => Execute(comAssemblyPath, manifestPath, action, pumpPendingMessages: true);

    /// <inheritdoc cref="Execute(string,string,Action)" />
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file.</param>
    /// <param name="action">Action to execute within the activation context.</param>
    /// <param name="pumpPendingMessages">Whether to pump messages after the callback. Disable when the host owns message processing.</param>
    public static Result Execute(
        string comAssemblyPath,
        string manifestPath,
        Action action,
        bool pumpPendingMessages)
    {
        ValidateSinglePaths(comAssemblyPath, manifestPath);
        ThrowHelper.ThrowIfNull(action, nameof(action));

        var descriptor = new ComPathDescriptor(comAssemblyPath, manifestPath);
        return Execute([descriptor], action, pumpPendingMessages);
    }

    /// <summary>
    /// Activates a single registration-free COM context with explicit <paramref name="options"/> and executes <paramref name="action"/>.
    /// </summary>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="action">Action to execute within the activation context.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the operation.</returns>
    public static Result Execute(
        string comAssemblyPath,
        string manifestPath,
        Action action,
        ComActivationOptions? options)
    {
        ValidateSinglePaths(comAssemblyPath, manifestPath);
        ThrowHelper.ThrowIfNull(action, nameof(action));

        var descriptor = new ComPathDescriptor(comAssemblyPath, manifestPath);
        return Execute([descriptor], action, options);
    }

    /// <summary>
    /// Activates multiple registration-free COM contexts, executes the supplied callback, pumps pending COM
    /// messages, and then releases the activation contexts in reverse order.
    /// </summary>
    /// <param name="comPathDescriptors">Collection of COM path descriptors containing DLL and manifest paths.</param>
    /// <param name="action">Action to execute within all activation contexts.</param>
    /// <returns>The outcome of the operation.</returns>
    public static Result Execute(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Action action)
        => Execute(comPathDescriptors, action, pumpPendingMessages: true);

    /// <inheritdoc cref="Execute(ICollection{ComPathDescriptor},Action)" />
    /// <param name="comPathDescriptors">Collection of DLL and manifest paths.</param>
    /// <param name="action">Action to execute within the activation contexts.</param>
    /// <param name="pumpPendingMessages">Whether to pump messages after the callback. Disable when the host owns message processing.</param>
    public static Result Execute(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Action action,
        bool pumpPendingMessages)
        => Execute(comPathDescriptors, action, pumpPendingMessages, ActivationContextApi.Instance);

    /// <summary>
    /// Activates multiple registration-free COM contexts with explicit <paramref name="options"/> and executes <paramref name="action"/>.
    /// </summary>
    /// <param name="comPathDescriptors">Collection of COM path descriptors.</param>
    /// <param name="action">Action to execute within all activation contexts.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the operation.</returns>
    public static Result Execute(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Action action,
        ComActivationOptions? options)
        => ExecuteCore(comPathDescriptors, action, CloneOrDefault(options), ActivationContextApi.Instance);

    /// <summary>
    /// Creates a COM object inside a single registration-free COM activation context and returns a handle that
    /// keeps that activation state alive until the object is explicitly released.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="factory">Factory that must create and return a non-<see langword="null"/> COM object while the activation context is active.</param>
    /// <returns>The outcome of the creation.</returns>
    public static ComObjectCreationResult<T> Create<T>(
        string comAssemblyPath,
        string manifestPath,
        Func<T> factory)
        where T : class
        => Create(comAssemblyPath, manifestPath, factory, pumpPendingMessages: true);

    /// <inheritdoc cref="Create{T}(string,string,Func{T})" />
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file.</param>
    /// <param name="factory">Factory transferring exclusive ownership of its non-null COM object.</param>
    /// <param name="pumpPendingMessages">Whether to pump messages after creation and during release. The handle retains this policy.</param>
    public static ComObjectCreationResult<T> Create<T>(
        string comAssemblyPath,
        string manifestPath,
        Func<T> factory,
        bool pumpPendingMessages)
        where T : class
    {
        ValidateSinglePaths(comAssemblyPath, manifestPath);
        ThrowHelper.ThrowIfNull(factory, nameof(factory));

        var descriptor = new ComPathDescriptor(comAssemblyPath, manifestPath);
        return Create([descriptor], factory, pumpPendingMessages);
    }

    /// <summary>
    /// Creates a COM object inside a single registration-free COM context with explicit <paramref name="options"/>.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="factory">Factory that creates the COM object while the activation context is active.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the creation.</returns>
    public static ComObjectCreationResult<T> Create<T>(
        string comAssemblyPath,
        string manifestPath,
        Func<T> factory,
        ComActivationOptions? options)
        where T : class
    {
        ValidateSinglePaths(comAssemblyPath, manifestPath);
        ThrowHelper.ThrowIfNull(factory, nameof(factory));

        var descriptor = new ComPathDescriptor(comAssemblyPath, manifestPath);
        return Create([descriptor], factory, options);
    }

    /// <summary>
    /// Creates a COM object inside multiple registration-free COM activation contexts and returns a handle that
    /// keeps those activation states alive until the object is explicitly released.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comPathDescriptors">Collection of COM path descriptors containing DLL and manifest paths.</param>
    /// <param name="factory">Factory that must create and return a non-<see langword="null"/> COM object while the activation contexts are active.</param>
    /// <returns>The outcome of the creation.</returns>
    public static ComObjectCreationResult<T> Create<T>(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Func<T> factory)
        where T : class
        => Create(comPathDescriptors, factory, pumpPendingMessages: true);

    /// <inheritdoc cref="Create{T}(ICollection{ComPathDescriptor},Func{T})" />
    /// <param name="comPathDescriptors">Collection of DLL and manifest paths.</param>
    /// <param name="factory">Factory transferring exclusive ownership of its non-null COM object.</param>
    /// <param name="pumpPendingMessages">Whether to pump messages after creation and during release. The handle retains this policy.</param>
    public static ComObjectCreationResult<T> Create<T>(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Func<T> factory,
        bool pumpPendingMessages)
        where T : class
        => Create(comPathDescriptors, factory, pumpPendingMessages, ActivationContextApi.Instance);

    /// <summary>
    /// Creates a COM object inside multiple registration-free COM contexts with explicit <paramref name="options"/>.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comPathDescriptors">Collection of COM path descriptors.</param>
    /// <param name="factory">Factory that creates the COM object while the activation contexts are active.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the creation.</returns>
    public static ComObjectCreationResult<T> Create<T>(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Func<T> factory,
        ComActivationOptions? options)
        where T : class
        => CreateCore(comPathDescriptors, factory, CloneOrDefault(options), ActivationContextApi.Instance);

    /// <summary>
    /// Activates a registration-free COM scope on the calling thread so multiple coclasses can be instantiated
    /// across requests until the scope is disposed or passed to <see cref="Deactivate(IComActivationScope)"/>.
    /// </summary>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the scope activation.</returns>
    public static ComCreateResult<IComActivationScope> Activate(
        string comAssemblyPath,
        string manifestPath,
        ComActivationOptions? options = null)
    {
        ValidateSinglePaths(comAssemblyPath, manifestPath);
        var descriptor = new ComPathDescriptor(comAssemblyPath, manifestPath);
        return Activate([descriptor], options);
    }

    /// <summary>
    /// Activates multiple registration-free COM contexts as a single LIFO scope on the calling thread.
    /// </summary>
    /// <param name="comPathDescriptors">Collection of COM path descriptors.</param>
    /// <param name="options">Optional activation and module-lifetime options.</param>
    /// <returns>The outcome of the scope activation.</returns>
    public static ComCreateResult<IComActivationScope> Activate(
        ICollection<ComPathDescriptor> comPathDescriptors,
        ComActivationOptions? options = null)
        => Activate(comPathDescriptors, options, ActivationContextApi.Instance);

    /// <summary>
    /// Deactivates an <see cref="IComActivationScope"/> previously returned by <c>Activate</c>, releasing any
    /// scope-tracked RCWs, reversing activation cookies and factory registrations, and unloading leased modules
    /// when <see cref="ComActivationOptions.UnloadModulesOnDeactivate"/> is enabled and <c>DllCanUnloadNow</c> allows it.
    /// </summary>
    /// <param name="activationScope">The scope to deactivate.</param>
    /// <returns>The outcome of the deactivation.</returns>
    public static Result Deactivate(IComActivationScope activationScope)
    {
        ThrowHelper.ThrowIfNull(activationScope, nameof(activationScope));
        if (activationScope is not ComActivationScope scope)
        {
            return new Result
            {
                Success = false,
                Exception = new ArgumentException("Unsupported activation scope implementation.", nameof(activationScope)),
            };
        }

        if (scope.IsReleased)
        {
            return new Result { Success = true };
        }

        var result = new Result { Success = false };
        if (!TryValidateScopeRelease(scope, result))
        {
            return result;
        }

        scope.IsReleasing = true;
        try
        {
            ReleaseTrackedObjects(scope.TrackedObjects);
            if (scope.Options.PumpPendingMessages)
            {
                NativeMethods.PumpPendingMessages();
            }

            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }
        finally
        {
            if (scope.TrackedObjects.Count == 0
                && CleanupContexts(
                    scope.ActivationCookies,
                    scope.ActivationContextHandles,
                    scope.ModuleLeases,
                    scope.Options.UnloadModulesOnDeactivate,
                    result,
                    scope.ContextApi))
            {
                scope.MarkReleased();
            }

            scope.IsReleasing = false;
        }

        return result;
    }

    /// <summary>
    /// Releases a COM object handle that was previously created by <c>Create</c>, then tears down the activation
    /// contexts and optional module leases that keep the object alive.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comObjectHandle">The COM object handle to release.</param>
    /// <returns>The outcome of the release.</returns>
    public static Result Free<T>(ComObjectHandle<T> comObjectHandle)
        where T : class
    {
        ThrowHelper.ThrowIfNull(comObjectHandle, nameof(comObjectHandle));

        if (comObjectHandle.IsReleased)
        {
            return new Result { Success = true };
        }

        var result = new Result { Success = false };
        try
        {
            comObjectHandle.ValidateRelease();
            ValidateReleaseOrder(comObjectHandle.ActivationCookies, comObjectHandle.ActivationContextHandles, comObjectHandle.ContextApi);
        }
        catch (Exception ex)
        {
            result.Exception = ex;
            return result;
        }

        comObjectHandle.IsReleasing = true;
        try
        {
            var comObject = comObjectHandle.ComObject;
            if (comObject is not null && Marshal.IsComObject(comObject))
            {
                Marshal.FinalReleaseComObject(comObject);
            }

            comObjectHandle.ComObject = null;
            if (comObjectHandle.PumpPendingMessages)
            {
                NativeMethods.PumpPendingMessages();
            }

            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }
        finally
        {
            if (comObjectHandle.ComObject is null
                && CleanupContexts(
                    comObjectHandle.ActivationCookies,
                    comObjectHandle.ActivationContextHandles,
                    comObjectHandle.ModuleLeases,
                    comObjectHandle.UnloadModulesOnDeactivate,
                    result,
                    comObjectHandle.ContextApi))
            {
                comObjectHandle.MarkReleased();
            }

            comObjectHandle.IsReleasing = false;
        }

        return result;
    }

    /// <summary>
    /// Sweeps any previously deactivated opt-in unload modules whose <c>DllCanUnloadNow</c> earlier reported
    /// live native objects, unloading those that have since become idle.
    /// </summary>
    /// <returns>The number of modules unloaded during this sweep.</returns>
    public static int TryUnloadUnusedModules()
        => ActivationContextApi.CollectDeferredUnloads();

    internal static Result Execute(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Action action,
        bool pumpPendingMessages,
        ActivationContextApi contextApi)
        => ExecuteCore(
            comPathDescriptors,
            action,
            new ComActivationOptions { PumpPendingMessages = pumpPendingMessages },
            contextApi);

    internal static ComObjectCreationResult<T> Create<T>(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Func<T> factory,
        bool pumpPendingMessages,
        ActivationContextApi contextApi)
        where T : class
        => CreateCore(
            comPathDescriptors,
            factory,
            new ComActivationOptions { PumpPendingMessages = pumpPendingMessages },
            contextApi);

    internal static ComCreateResult<IComActivationScope> Activate(
        ICollection<ComPathDescriptor> comPathDescriptors,
        ComActivationOptions? options,
        ActivationContextApi contextApi)
    {
        ValidateComPathDescriptors(comPathDescriptors);
        NanoServerOptions.EnsureCurrentStaThreadInitialized();
        var effectiveOptions = CloneOrDefault(options);

        var result = new ComCreateResult<IComActivationScope> { Success = false };
        var hActCtxs = new List<IntPtr>(comPathDescriptors.Count);
        var cookies = new List<IntPtr>(comPathDescriptors.Count);
        var moduleLeases = new List<ActivationContextApi.ModuleLease>();
        var referencedDllPaths = new List<string>();

        try
        {
            SetupActivationState(
                comPathDescriptors,
                effectiveOptions,
                hActCtxs,
                cookies,
                moduleLeases,
                referencedDllPaths,
                contextApi);

            result.Value = new ComActivationScope(
                hActCtxs,
                cookies,
                moduleLeases,
                referencedDllPaths,
                effectiveOptions,
                contextApi);
            if (effectiveOptions.PumpPendingMessages)
            {
                NativeMethods.PumpPendingMessages();
            }

            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }
        finally
        {
            RollbackFailedScopeActivation(result, cookies, hActCtxs, moduleLeases, effectiveOptions.UnloadModulesOnDeactivate, contextApi);
        }

        return result;
    }

    private static void RollbackFailedScopeActivation(
        ComCreateResult<IComActivationScope> result,
        List<IntPtr> cookies,
        List<IntPtr> hActCtxs,
        List<ActivationContextApi.ModuleLease> moduleLeases,
        bool unloadModulesOnDeactivate,
        ActivationContextApi contextApi)
    {
        if (result.Success)
        {
            return;
        }

        if (result.Value is not null)
        {
            var deactivate = Deactivate(result.Value);
            if (!deactivate.Success && deactivate.Exception is not null)
            {
                RecordFailure(result, deactivate.Exception);
            }

            result.Value = null;
            return;
        }

        CleanupContexts(
            cookies,
            hActCtxs,
            moduleLeases,
            unloadModulesOnDeactivate,
            result,
            contextApi);
    }

    private static bool TryValidateScopeRelease(ComActivationScope scope, Result result)
    {
        try
        {
            scope.ValidateRelease();
            ValidateReleaseOrder(scope.ActivationCookies, scope.ActivationContextHandles, scope.ContextApi);
            return true;
        }
        catch (Exception ex)
        {
            result.Exception = ex;
            return false;
        }
    }

    private static Result ExecuteCore(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Action action,
        ComActivationOptions effectiveOptions,
        ActivationContextApi contextApi)
    {
        ValidateComPathDescriptors(comPathDescriptors);
        ThrowHelper.ThrowIfNull(action, nameof(action));
        NanoServerOptions.EnsureCurrentStaThreadInitialized();

        var result = new Result { Success = false };
        var hActCtxs = new List<IntPtr>(comPathDescriptors.Count);
        var cookies = new List<IntPtr>(comPathDescriptors.Count);
        var moduleLeases = new List<ActivationContextApi.ModuleLease>();
        var referencedDllPaths = new List<string>();
        try
        {
            SetupActivationState(
                comPathDescriptors,
                effectiveOptions,
                hActCtxs,
                cookies,
                moduleLeases,
                referencedDllPaths,
                contextApi);
#pragma warning disable CC0031 // Validated by ThrowHelper before allocating native resources.
            action();
#pragma warning restore CC0031

            if (effectiveOptions.PumpPendingMessages)
            {
                NativeMethods.PumpPendingMessages();
            }

            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }
        finally
        {
            CleanupContexts(
                cookies,
                hActCtxs,
                moduleLeases,
                effectiveOptions.UnloadModulesOnDeactivate,
                result,
                contextApi);
        }

        return result;
    }

    private static ComObjectCreationResult<T> CreateCore<T>(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Func<T> factory,
        ComActivationOptions effectiveOptions,
        ActivationContextApi contextApi)
        where T : class
    {
        ValidateComPathDescriptors(comPathDescriptors);
        ThrowHelper.ThrowIfNull(factory, nameof(factory));
        NanoServerOptions.EnsureCurrentStaThreadInitialized();

        var result = new ComObjectCreationResult<T> { Success = false };
        var hActCtxs = new List<IntPtr>(comPathDescriptors.Count);
        var cookies = new List<IntPtr>(comPathDescriptors.Count);
        var moduleLeases = new List<ActivationContextApi.ModuleLease>();
        var referencedDllPaths = new List<string>();
        try
        {
            SetupActivationState(
                comPathDescriptors,
                effectiveOptions,
                hActCtxs,
                cookies,
                moduleLeases,
                referencedDllPaths,
                contextApi);

#pragma warning disable CC0031
            var comObject = factory() ?? throw new InvalidOperationException("The COM factory returned null.");
#pragma warning restore CC0031

            result.Value = new ComObjectHandle<T>(
                comObject,
                hActCtxs,
                cookies,
                effectiveOptions.PumpPendingMessages,
                contextApi,
                moduleLeases,
                effectiveOptions.UnloadModulesOnDeactivate);
            if (effectiveOptions.PumpPendingMessages)
            {
                NativeMethods.PumpPendingMessages();
            }

            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }
        finally
        {
            RollbackFailedObjectCreation(result, cookies, hActCtxs, moduleLeases, effectiveOptions.UnloadModulesOnDeactivate, contextApi);
        }

        return result;
    }

    private static void RollbackFailedObjectCreation<T>(
        ComObjectCreationResult<T> result,
        List<IntPtr> cookies,
        List<IntPtr> hActCtxs,
        List<ActivationContextApi.ModuleLease> moduleLeases,
        bool unloadModulesOnDeactivate,
        ActivationContextApi contextApi)
        where T : class
    {
        if (result.Success)
        {
            return;
        }

        if (result.Value is not null)
        {
            var release = Free(result.Value);
            if (!release.Success && release.Exception is not null)
            {
                RecordFailure(result, release.Exception);
            }

            result.Value = null;
            return;
        }

        CleanupContexts(
            cookies,
            hActCtxs,
            moduleLeases,
            unloadModulesOnDeactivate,
            result,
            contextApi);
    }

    private static ComActivationOptions CloneOrDefault(ComActivationOptions? options)
        => options is null ? ComActivationOptions.Default.Clone() : options.Clone();

    private static void ValidateSinglePaths(string comAssemblyPath, string manifestPath)
    {
        ThrowHelper.ThrowIfNull(comAssemblyPath, nameof(comAssemblyPath));
        ThrowHelper.ThrowIfNull(manifestPath, nameof(manifestPath));

        if (string.IsNullOrWhiteSpace(comAssemblyPath))
        {
            throw new ArgumentException("COM assembly path cannot be empty or whitespace.", nameof(comAssemblyPath));
        }

        if (string.IsNullOrWhiteSpace(manifestPath))
        {
            throw new ArgumentException("Manifest path cannot be empty or whitespace.", nameof(manifestPath));
        }
    }

    private static void ValidateComPathDescriptors(ICollection<ComPathDescriptor> comPathDescriptors)
    {
        ThrowHelper.ThrowIfNull(comPathDescriptors, nameof(comPathDescriptors));

        if (comPathDescriptors.Count == 0)
        {
            throw new ArgumentException("COM path descriptors collection cannot be empty.", nameof(comPathDescriptors));
        }

        if (comPathDescriptors.Any(descriptor => descriptor is null))
        {
            throw new ArgumentException("COM path descriptors collection cannot contain null.", nameof(comPathDescriptors));
        }
    }

    private static void SetupActivationState(
        ICollection<ComPathDescriptor> comPathDescriptors,
        ComActivationOptions options,
        List<IntPtr> hActCtxs,
        List<IntPtr> cookies,
        List<ActivationContextApi.ModuleLease> moduleLeases,
        List<string> referencedDllPaths,
        ActivationContextApi contextApi)
    {
        var dllPathSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var hasNativeSxsContext = false;

        foreach (var comPathDescriptor in comPathDescriptors)
        {
            var ac = PrepareContext(comPathDescriptor, out var fullAssemblyPath);
            var hActCtx = CreateContext(ac, fullAssemblyPath, options, contextApi, out var parsedManifest);
            hActCtxs.Add(hActCtx);

            if (!ActivationContextApi.IsFallbackHandle(hActCtx))
            {
                hasNativeSxsContext = true;
            }

            if (parsedManifest is not null)
            {
                AppendReferencedDllPaths(parsedManifest.ReferencedDllPaths, dllPathSet, referencedDllPaths);
            }

            if (dllPathSet.Add(fullAssemblyPath))
            {
                referencedDllPaths.Add(fullAssemblyPath);
            }
        }

        ActivateContexts(hActCtxs, cookies, contextApi);

        if (hasNativeSxsContext && (options.PreloadModules || options.UnloadModulesOnDeactivate))
        {
            AcquireNativeSxsModuleLeases(options, referencedDllPaths, moduleLeases, contextApi);
        }
    }

    private static void AppendReferencedDllPaths(
        IReadOnlyList<string> manifestDllPaths,
        HashSet<string> dllPathSet,
        List<string> referencedDllPaths)
    {
        for (var i = 0; i < manifestDllPaths.Count; i++)
        {
            var path = manifestDllPaths[i];
            if (dllPathSet.Add(path))
            {
                referencedDllPaths.Add(path);
            }
        }
    }

    private static void AcquireNativeSxsModuleLeases(
        ComActivationOptions options,
        List<string> referencedDllPaths,
        List<ActivationContextApi.ModuleLease> moduleLeases,
        ActivationContextApi contextApi)
    {
        for (var i = 0; i < referencedDllPaths.Count; i++)
        {
            var dllPath = referencedDllPaths[i];
            if (!options.PreloadModules && !File.Exists(dllPath))
            {
                continue;
            }

            var lease = contextApi.AcquireModuleLease(dllPath, options.UnloadModulesOnDeactivate);
            if (lease is null || lease.Handle == IntPtr.Zero)
            {
                if (options.PreloadModules)
                {
                    throw CreateWin32Exception(contextApi.GetLastWin32Error());
                }

                continue;
            }

            moduleLeases.Add(lease);
        }
    }

    private static void ActivateContexts(List<IntPtr> activationContextHandles, List<IntPtr> cookies, ActivationContextApi contextApi)
    {
        foreach (var hActCtx in activationContextHandles)
        {
            var cookie = ActivateContext(hActCtx, contextApi);
            cookies.Add(cookie);
        }
    }

    private static void ReleaseTrackedObjects(List<object> trackedObjects)
    {
        while (trackedObjects.Count > 0)
        {
            var index = trackedObjects.Count - 1;
            var obj = trackedObjects[index];
            if (Marshal.IsComObject(obj))
            {
                Marshal.FinalReleaseComObject(obj);
            }

            trackedObjects.RemoveAt(index);
        }
    }

    private static void ValidateReleaseOrder(List<IntPtr> cookies, List<IntPtr> handles, ActivationContextApi contextApi)
    {
        if (cookies.Count == 0)
        {
            return;
        }

        var stack = _activeCookies;
        if (stack is null || stack.Count < cookies.Count)
        {
            throw new InvalidOperationException("Activation contexts must be released on the creating thread in reverse order.");
        }

        for (var i = 0; i < cookies.Count; i++)
        {
            if (stack[stack.Count - cookies.Count + i] != cookies[i])
            {
                throw new InvalidOperationException("Release nested handles and callbacks before releasing this handle.");
            }
        }

        if (!contextApi.GetCurrentActCtx(out var current))
        {
            throw new Win32Exception();
        }

        try
        {
            if (current != handles[cookies.Count - 1])
            {
                throw new InvalidOperationException("A different activation context is active. Release it before releasing this handle.");
            }
        }
        finally
        {
            if (current != IntPtr.Zero)
            {
                contextApi.ReleaseActCtx(current);
            }
        }
    }

    private static bool CleanupContexts(
        List<IntPtr> cookies,
        List<IntPtr> handles,
        List<ActivationContextApi.ModuleLease> moduleLeases,
        bool unloadModulesOnDeactivate,
        Result result,
        ActivationContextApi contextApi)
    {
        try
        {
            ValidateReleaseOrder(cookies, handles, contextApi);
            while (cookies.Count > 0)
            {
                var index = cookies.Count - 1;
                if (!contextApi.DeactivateActCtx(cookies[index]))
                {
                    throw CreateWin32Exception(contextApi.ConsumeLastError());
                }

                cookies.RemoveAt(index);
                _activeCookies!.RemoveAt(_activeCookies.Count - 1);
            }

            ReleaseActivationContexts(handles, contextApi);
            handles.Clear();

            if (moduleLeases.Count > 0 || unloadModulesOnDeactivate)
            {
                contextApi.ReleaseModuleLeases(moduleLeases, forceFreeUnusedLibraries: unloadModulesOnDeactivate);
            }

            return true;
        }
        catch (Exception ex)
        {
            RecordFailure(result, ex);
            return false;
        }
    }

    private static void RecordFailure(Result result, Exception exception)
    {
        result.Success = false;
        result.Exception = result.Exception is null ? exception : new AggregateException(result.Exception, exception);
    }

    private static void ReleaseActivationContexts(List<IntPtr> activationContextHandles, ActivationContextApi contextApi)
    {
        for (var i = activationContextHandles.Count - 1; i >= 0; i--)
        {
            contextApi.ReleaseActCtx(activationContextHandles[i]);
        }
    }

    private static ActCtx PrepareContext(ComPathDescriptor comPathDescriptor, out string fullAssemblyPath)
    {
        var ac = default(ActCtx);
        ac.cbSize = Marshal.SizeOf<ActCtx>();
        var expected = IntPtr.Size == 4 ? NativeActCtxSizeX86 : NativeActCtxSizeX64;
        if (ac.cbSize != expected)
        {
            throw new ActCtxWrongSizeException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "ActCtx.cbSize is wrong (expected {0} bytes, got {1}).",
                    expected,
                    ac.cbSize));
        }

        fullAssemblyPath = Path.GetFullPath(comPathDescriptor.ComAssemblyPath);
        ac.lpAssemblyDirectory = Path.GetDirectoryName(fullAssemblyPath)
            ?? throw new ArgumentException("COM assembly path must have a containing directory.", nameof(comPathDescriptor));
        ac.lpSource = Path.GetFullPath(comPathDescriptor.ComManifestPath);
        ac.dwFlags = NativeMethods.ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID;
        return ac;
    }

    private static IntPtr CreateContext(
        ActCtx actCtx,
        string fullAssemblyPath,
        ComActivationOptions options,
        ActivationContextApi contextApi,
        out ComManifestParser.ParsedManifest? parsedManifest)
    {
        var hActCtx = contextApi.CreateActCtxWithOptions(ref actCtx, fullAssemblyPath, options, out parsedManifest, out var errorCode);
        if (hActCtx == (IntPtr)(-1))
        {
            throw CreateWin32Exception(errorCode);
        }

        return hActCtx;
    }

    private static IntPtr ActivateContext(IntPtr hActCtx, ActivationContextApi contextApi)
    {
        var stack = _activeCookies ??= [];
        if (!contextApi.ActivateActCtx(hActCtx, out var cookie))
        {
            throw CreateWin32Exception(contextApi.ConsumeLastError());
        }

        stack.Add(cookie);
        return cookie;
    }

    private static Win32Exception CreateWin32Exception(int errorCode)
        => errorCode == 0 ? new Win32Exception() : new Win32Exception(errorCode);
}
#pragma warning restore SA1121
