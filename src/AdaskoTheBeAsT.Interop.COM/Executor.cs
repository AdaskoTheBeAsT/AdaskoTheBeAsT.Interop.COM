using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Provides helpers for activating registration-free COM components, executing work inside their activation
/// contexts, and explicitly managing longer-lived COM object lifetimes on the current thread.
/// </summary>
/// <remarks>
/// <para>
/// <b>New code should prefer <see cref="IComExecutor"/> (implemented by <see cref="ComExecutor"/>)</b>, which
/// exposes the same surface through an injectable, testable interface. This static class is retained for
/// source-level compatibility with v2.x callers.
/// </para>
/// <para>
/// Work runs synchronously on the calling thread. This library does not initialize COM or change the
/// apartment state. Callers must provide an STA thread when required by their component.
/// Nested handles created inside a callback or factory must be released before that delegate returns.
/// </para>
/// </remarks>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public static class Executor
{
    /// <summary>
    /// Expected <c>sizeof(ACTCTXW)</c> on 32-bit Windows, as defined by the Windows SDK.
    /// Compared against <see cref="Marshal.SizeOf(Type)"/> of the managed <c>ActCtx</c> struct to catch ABI drift.
    /// </summary>
    private const int NativeActCtxSizeX86 = 0x20;

    /// <summary>
    /// Expected <c>sizeof(ACTCTXW)</c> on 64-bit Windows, as defined by the Windows SDK.
    /// Compared against <see cref="Marshal.SizeOf(Type)"/> of the managed <c>ActCtx</c> struct to catch ABI drift.
    /// </summary>
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
    /// <returns>
    /// A <see cref="Result"/> whose <see cref="Result.Success"/> value is <see langword="true"/> when the callback
    /// completes successfully. When the operation fails, <see cref="Result.Success"/> is
    /// <see langword="false"/> and <see cref="Result.Exception"/> contains the captured exception.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when comAssemblyPath, manifestPath, or action is null.</exception>
    /// <exception cref="ArgumentException">Thrown when comAssemblyPath or manifestPath is empty or whitespace.</exception>
    /// <remarks>Exceptions raised while activating the context or running <paramref name="action"/> are captured in the returned result instead of being rethrown.</remarks>
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
        ThrowHelper.ThrowIfNull(comAssemblyPath, nameof(comAssemblyPath));
        ThrowHelper.ThrowIfNull(manifestPath, nameof(manifestPath));
        ThrowHelper.ThrowIfNull(action, nameof(action));

        if (string.IsNullOrWhiteSpace(comAssemblyPath))
        {
            throw new ArgumentException("COM assembly path cannot be empty or whitespace.", nameof(comAssemblyPath));
        }

        if (string.IsNullOrWhiteSpace(manifestPath))
        {
            throw new ArgumentException("Manifest path cannot be empty or whitespace.", nameof(manifestPath));
        }

        var descriptor = new ComPathDescriptor(comAssemblyPath, manifestPath);
        return Execute([descriptor], action, pumpPendingMessages);
    }

    /// <summary>
    /// Activates multiple registration-free COM contexts, executes the supplied callback, pumps pending COM
    /// messages, and then releases the activation contexts in reverse order.
    /// </summary>
    /// <param name="comPathDescriptors">Collection of COM path descriptors containing DLL and manifest paths.</param>
    /// <param name="action">Action to execute within all activation contexts.</param>
    /// <returns>
    /// A <see cref="Result"/> whose <see cref="Result.Success"/> value is <see langword="true"/> when the callback
    /// completes successfully. When the operation fails, <see cref="Result.Success"/> is
    /// <see langword="false"/> and <see cref="Result.Exception"/> contains the captured exception.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when comPathDescriptors or action is null.</exception>
    /// <exception cref="ArgumentException">Thrown when comPathDescriptors is empty or contains null.</exception>
    /// <remarks>Exceptions raised while activating the contexts or running <paramref name="action"/> are captured in the returned result instead of being rethrown.</remarks>
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
    /// Creates a COM object inside a single registration-free COM activation context and returns a handle that
    /// keeps that activation state alive until the object is explicitly released.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="manifestPath">Full path to the manifest file describing the COM component.</param>
    /// <param name="factory">Factory that must create and return a non-<see langword="null"/> COM object while the activation context is active.</param>
    /// <returns>
    /// A <see cref="ComObjectCreationResult{T}"/>. On success, <see cref="Result.Success"/> is
    /// <see langword="true"/> and <see cref="ComObjectCreationResult{T}.Value"/> contains the created
    /// <see cref="ComObjectHandle{T}"/>. On failure, <see cref="Result.Success"/> is <see langword="false"/>,
    /// <see cref="ComObjectCreationResult{T}.Value"/> is <see langword="null"/>, and
    /// <see cref="Result.Exception"/> contains the captured exception.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when comAssemblyPath, manifestPath, or factory is null.</exception>
    /// <exception cref="ArgumentException">Thrown when comAssemblyPath or manifestPath is empty or whitespace.</exception>
    /// <remarks>
    /// The returned handle is thread-affine. Create, use, and release it on the same thread by calling
    /// <see cref="Free{T}(ComObjectHandle{T})"/>. Release handles in reverse creation order.
    /// The factory transfers exclusive ownership of its RCW. Do not return a borrowed or shared RCW:
    /// release calls <c>Marshal.FinalReleaseComObject</c> and invalidates every alias to that wrapper.
    /// </remarks>
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
        ThrowHelper.ThrowIfNull(comAssemblyPath, nameof(comAssemblyPath));
        ThrowHelper.ThrowIfNull(manifestPath, nameof(manifestPath));
        ThrowHelper.ThrowIfNull(factory, nameof(factory));

        if (string.IsNullOrWhiteSpace(comAssemblyPath))
        {
            throw new ArgumentException("COM assembly path cannot be empty or whitespace.", nameof(comAssemblyPath));
        }

        if (string.IsNullOrWhiteSpace(manifestPath))
        {
            throw new ArgumentException("Manifest path cannot be empty or whitespace.", nameof(manifestPath));
        }

        var descriptor = new ComPathDescriptor(comAssemblyPath, manifestPath);
        return Create([descriptor], factory, pumpPendingMessages);
    }

    /// <summary>
    /// Creates a COM object inside multiple registration-free COM activation contexts and returns a handle that
    /// keeps those activation states alive until the object is explicitly released.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comPathDescriptors">Collection of COM path descriptors containing DLL and manifest paths.</param>
    /// <param name="factory">Factory that must create and return a non-<see langword="null"/> COM object while the activation contexts are active.</param>
    /// <returns>
    /// A <see cref="ComObjectCreationResult{T}"/>. On success, <see cref="Result.Success"/> is
    /// <see langword="true"/> and <see cref="ComObjectCreationResult{T}.Value"/> contains the created
    /// <see cref="ComObjectHandle{T}"/>. On failure, <see cref="Result.Success"/> is <see langword="false"/>,
    /// <see cref="ComObjectCreationResult{T}.Value"/> is <see langword="null"/>, and
    /// <see cref="Result.Exception"/> contains the captured exception.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when comPathDescriptors or factory is null.</exception>
    /// <exception cref="ArgumentException">Thrown when comPathDescriptors is empty or contains null.</exception>
    /// <remarks>
    /// The returned handle is thread-affine. Create, use, and release it on the same thread by calling
    /// <see cref="Free{T}(ComObjectHandle{T})"/>. Release handles in reverse creation order.
    /// The factory transfers exclusive ownership of its RCW. Do not return a borrowed or shared RCW:
    /// release calls <c>Marshal.FinalReleaseComObject</c> and invalidates every alias to that wrapper.
    /// </remarks>
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
    /// Releases a COM object handle that was previously created by <see cref="Create{T}(string,string,Func{T})"/>
    /// or <see cref="Create{T}(ICollection{ComPathDescriptor},Func{T})"/>, then tears down the activation
    /// contexts that keep the object alive.
    /// </summary>
    /// <typeparam name="T">The COM object type.</typeparam>
    /// <param name="comObjectHandle">The COM object handle to release.</param>
    /// <returns>
    /// A <see cref="Result"/> whose <see cref="Result.Success"/> value is <see langword="true"/> when the COM
    /// object and activation contexts are released successfully. Releasing an already released handle returns a
    /// successful result without performing additional work.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when comObjectHandle is null.</exception>
    /// <remarks>
    /// This method is idempotent and must be called on the creating thread, in reverse creation order.
    /// An invalid thread or release order returns a failed result without releasing the object.
    /// After successful release,
    /// <see cref="ComObjectHandle{T}.ComObject"/> becomes <see langword="null"/> and the handle should no longer be used.
    /// </remarks>
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
                && CleanupContexts(comObjectHandle.ActivationCookies, comObjectHandle.ActivationContextHandles, result, comObjectHandle.ContextApi))
            {
                comObjectHandle.MarkReleased();
            }

            comObjectHandle.IsReleasing = false;
        }

        return result;
    }

    internal static Result Execute(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Action action,
        bool pumpPendingMessages,
        ActivationContextApi contextApi)
    {
        ValidateComPathDescriptors(comPathDescriptors);
        ThrowHelper.ThrowIfNull(action, nameof(action));

        var result = new Result { Success = false };
        var hActCtxs = new List<IntPtr>(comPathDescriptors.Count);
        var cookies = new List<IntPtr>(comPathDescriptors.Count);
        try
        {
            CreateActivationContexts(comPathDescriptors, hActCtxs, contextApi);
            ActivateContexts(hActCtxs, cookies, contextApi);
#pragma warning disable CC0031 // Validated by ThrowHelper before allocating native resources.
            action();
#pragma warning restore CC0031

            if (pumpPendingMessages)
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
            CleanupContexts(cookies, hActCtxs, result, contextApi);
        }

        return result;
    }

    internal static ComObjectCreationResult<T> Create<T>(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Func<T> factory,
        bool pumpPendingMessages,
        ActivationContextApi contextApi)
        where T : class
    {
        ValidateComPathDescriptors(comPathDescriptors);
        ThrowHelper.ThrowIfNull(factory, nameof(factory));

        var result = new ComObjectCreationResult<T> { Success = false };
        var hActCtxs = new List<IntPtr>(comPathDescriptors.Count);
        var cookies = new List<IntPtr>(comPathDescriptors.Count);
        try
        {
            CreateActivationContexts(comPathDescriptors, hActCtxs, contextApi);
            ActivateContexts(hActCtxs, cookies, contextApi);

#pragma warning disable CC0031
            var comObject = factory() ?? throw new InvalidOperationException("The COM factory returned null.");
#pragma warning restore CC0031

            result.Value = new ComObjectHandle<T>(comObject, hActCtxs, cookies, pumpPendingMessages, contextApi);
            if (pumpPendingMessages)
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
            if (!result.Success)
            {
                if (result.Value is not null)
                {
                    var release = Free(result.Value);
                    if (!release.Success && release.Exception is not null)
                    {
                        RecordFailure(result, release.Exception);
                    }

                    result.Value = null;
                }
                else
                {
                    CleanupContexts(cookies, hActCtxs, result, contextApi);
                }
            }
        }

        return result;
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

    private static void CreateActivationContexts(
        ICollection<ComPathDescriptor> comPathDescriptors,
        List<IntPtr> hActCtxs,
        ActivationContextApi contextApi)
    {
        foreach (var comPathDescriptor in comPathDescriptors)
        {
            var ac = PrepareContext(comPathDescriptor);
            var hActCtx = CreateContext(ac, contextApi);
            hActCtxs.Add(hActCtx);
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

        for (int i = 0; i < cookies.Count; i++)
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

    private static bool CleanupContexts(List<IntPtr> cookies, List<IntPtr> handles, Result result, ActivationContextApi contextApi)
    {
        try
        {
            ValidateReleaseOrder(cookies, handles, contextApi);
            while (cookies.Count > 0)
            {
                var index = cookies.Count - 1;
                if (!contextApi.DeactivateActCtx(cookies[index]))
                {
                    throw new Win32Exception();
                }

                cookies.RemoveAt(index);
                _activeCookies!.RemoveAt(_activeCookies.Count - 1);
            }

            ReleaseActivationContexts(handles, contextApi);
            handles.Clear();
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
        for (int i = activationContextHandles.Count - 1; i >= 0; i--)
        {
            contextApi.ReleaseActCtx(activationContextHandles[i]);
        }
    }

    private static ActCtx PrepareContext(ComPathDescriptor comPathDescriptor)
    {
        var ac = default(ActCtx);
        ac.cbSize = Marshal.SizeOf<ActCtx>();
        var expected = IntPtr.Size == 4 ? NativeActCtxSizeX86 : NativeActCtxSizeX64;
        if (ac.cbSize != expected)
        {
            throw new ActCtxWrongSizeException(
                string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    "ActCtx.cbSize is wrong (expected {0} bytes, got {1}).",
                    expected,
                    ac.cbSize));
        }

        ac.lpAssemblyDirectory = Path.GetDirectoryName(Path.GetFullPath(comPathDescriptor.ComAssemblyPath))
            ?? throw new ArgumentException("COM assembly path must have a containing directory.", nameof(comPathDescriptor));
        ac.lpSource = Path.GetFullPath(comPathDescriptor.ComManifestPath);
        ac.dwFlags = NativeMethods.ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID;
        return ac;
    }

    private static IntPtr CreateContext(ActCtx actCtx, ActivationContextApi contextApi)
    {
        var hActCtx = contextApi.CreateActCtx(ref actCtx);
        if (hActCtx == (IntPtr)(-1))
        {
            throw new Win32Exception();
        }

        return hActCtx;
    }

    private static IntPtr ActivateContext(IntPtr hActCtx, ActivationContextApi contextApi)
    {
        var stack = _activeCookies ??= [];
        if (!contextApi.ActivateActCtx(hActCtx, out var cookie))
        {
            throw new Win32Exception();
        }

        stack.Add(cookie);
        return cookie;
    }
}
