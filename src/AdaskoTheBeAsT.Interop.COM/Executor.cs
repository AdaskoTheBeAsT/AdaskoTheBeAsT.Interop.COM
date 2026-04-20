using System;
using System.Collections.Generic;
using System.ComponentModel;
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
    /// <exception cref="ArgumentNullException">Thrown when comAssemblyPath or manifestPath is null.</exception>
    /// <exception cref="ArgumentException">Thrown when comAssemblyPath or manifestPath is empty or whitespace.</exception>
    /// <remarks>Exceptions raised while activating the context or running <paramref name="action"/> are captured in the returned result instead of being rethrown.</remarks>
    public static Result Execute(
        string comAssemblyPath,
        string manifestPath,
        Action action)
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

        var descriptor = new ComPathDescriptor(comAssemblyPath, manifestPath);
        var ctx = PrepareContext(descriptor);
        var hActCtx = CreateContext(ctx);

        var result = new Result { Success = false };
        try
        {
            var cookie = ActivateContext(hActCtx);

            try
            {
                action?.Invoke();

                // Pump COM messages in STA apartment
                NativeMethods.PumpPendingMessages();

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Exception = ex;
            }
            finally
            {
                NativeMethods.DeactivateActCtx(0, cookie);
            }
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }
        finally
        {
            NativeMethods.ReleaseActCtx(hActCtx);
        }

        return result;
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
    /// <exception cref="ArgumentNullException">Thrown when comPathDescriptors is null.</exception>
    /// <exception cref="ArgumentException">Thrown when comPathDescriptors is empty.</exception>
    /// <remarks>Exceptions raised while activating the contexts or running <paramref name="action"/> are captured in the returned result instead of being rethrown.</remarks>
    public static Result Execute(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Action action)
    {
        ValidateComPathDescriptors(comPathDescriptors);
        var hActCtxs = CreateActivationContexts(comPathDescriptors);

        var result = new Result { Success = false };
        var cookies = new List<IntPtr>(hActCtxs.Count);
        try
        {
            cookies = ActivateContexts(hActCtxs);

            try
            {
                action?.Invoke();

                // Pump COM messages after callback
                NativeMethods.PumpPendingMessages();

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Exception = ex;
            }
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }
        finally
        {
            DeactivateContexts(cookies);
            ReleaseActivationContexts(hActCtxs);
        }

        return result;
    }

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
    /// <see cref="Free{T}(ComObjectHandle{T})"/>.
    /// </remarks>
    public static ComObjectCreationResult<T> Create<T>(
        string comAssemblyPath,
        string manifestPath,
        Func<T> factory)
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
        return Create([descriptor], factory);
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
    /// <exception cref="ArgumentException">Thrown when comPathDescriptors is empty.</exception>
    /// <remarks>
    /// The returned handle is thread-affine. Create, use, and release it on the same thread by calling
    /// <see cref="Free{T}(ComObjectHandle{T})"/>.
    /// </remarks>
    public static ComObjectCreationResult<T> Create<T>(
        ICollection<ComPathDescriptor> comPathDescriptors,
        Func<T> factory)
        where T : class
    {
        ValidateComPathDescriptors(comPathDescriptors);
        ThrowHelper.ThrowIfNull(factory, nameof(factory));

        var hActCtxs = CreateActivationContexts(comPathDescriptors);

        var result = new ComObjectCreationResult<T> { Success = false };
        var cookies = new List<IntPtr>(hActCtxs.Count);
        try
        {
            cookies = ActivateContexts(hActCtxs);

            try
            {
#pragma warning disable CC0031
                var comObject = factory() ?? throw new InvalidOperationException("The COM factory returned null.");
#pragma warning restore CC0031

                // Pump COM messages in STA apartment
                NativeMethods.PumpPendingMessages();

                result.Value = new ComObjectHandle<T>(comObject, [.. hActCtxs], [.. cookies]);
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Exception = ex;
            }
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }
        finally
        {
            if (!result.Success)
            {
                DeactivateContexts(cookies);
                ReleaseActivationContexts(hActCtxs);
            }
        }

        return result;
    }

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
    /// This method is idempotent and should be called on the same thread that created the handle. After release,
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
            var comObject = comObjectHandle.ComObject;
            if (comObject is not null && Marshal.IsComObject(comObject))
            {
                Marshal.FinalReleaseComObject(comObject);
            }

            NativeMethods.PumpPendingMessages();
            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Exception = ex;
        }
        finally
        {
            DeactivateContexts(comObjectHandle.ActivationCookies);
            ReleaseActivationContexts(comObjectHandle.ActivationContextHandles);
            comObjectHandle.MarkReleased();
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
    }

    private static List<IntPtr> CreateActivationContexts(ICollection<ComPathDescriptor> comPathDescriptors)
    {
        var hActCtxs = new List<IntPtr>(comPathDescriptors.Count);

        foreach (var comPathDescriptor in comPathDescriptors)
        {
            var ac = PrepareContext(comPathDescriptor);
            var hActCtx = CreateContext(ac);
            hActCtxs.Add(hActCtx);
        }

        return hActCtxs;
    }

    private static List<IntPtr> ActivateContexts(List<IntPtr> activationContextHandles)
    {
        var cookies = new List<IntPtr>(activationContextHandles.Count);

        try
        {
            foreach (var hActCtx in activationContextHandles)
            {
                var cookie = ActivateContext(hActCtx);
                cookies.Add(cookie);
            }

            return cookies;
        }
        catch
        {
            DeactivateContexts(cookies);
            throw;
        }
    }

    private static void DeactivateContexts(IReadOnlyList<IntPtr> cookies)
    {
        for (int i = cookies.Count - 1; i >= 0; i--)
        {
            NativeMethods.DeactivateActCtx(0, cookies[i]);
        }
    }

    private static void ReleaseActivationContexts(IReadOnlyList<IntPtr> activationContextHandles)
    {
        for (int i = activationContextHandles.Count - 1; i >= 0; i--)
        {
            NativeMethods.ReleaseActCtx(activationContextHandles[i]);
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

        ac.lpAssemblyDirectory = comPathDescriptor.ComAssemblyPath;
        ac.lpSource = comPathDescriptor.ComManifestPath;
        ac.dwFlags = NativeMethods.ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID;
        return ac;
    }

    private static IntPtr CreateContext(ActCtx actCtx)
    {
        var hActCtx = NativeMethods.CreateActCtx(ref actCtx);
        if (hActCtx == (IntPtr)(-1))
        {
            throw new Win32Exception();
        }

        return hActCtx;
    }

    private static IntPtr ActivateContext(IntPtr hActCtx)
    {
        if (!NativeMethods.ActivateActCtx(hActCtx, out var cookie))
        {
            throw new Win32Exception();
        }

        return cookie;
    }
}
