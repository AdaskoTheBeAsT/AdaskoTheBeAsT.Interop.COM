using System;
using System.Collections.Generic;
using System.Diagnostics;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Represents a COM object together with the activation-context state that keeps it alive after
/// <see cref="Executor.Create{T}(string,string,Func{T})"/> or
/// <see cref="Executor.Create{T}(ICollection{ComPathDescriptor},Func{T})"/> returns.
/// Release it through <see cref="Dispose"/> (preferred, e.g. with a <c>using</c> statement) or
/// <see cref="Executor.Free{T}(ComObjectHandle{T})"/>.
/// </summary>
/// <typeparam name="T">The COM object type.</typeparam>
/// <remarks>
/// Instances are thread-affine and should be used and released on the same thread that created them.
/// A finalizer is provided as a diagnostic safety net only; failing to release the handle leaks native
/// activation-context state and COM references, so always prefer deterministic disposal.
/// </remarks>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public sealed class ComObjectHandle<T>
    : IDisposable
    where T : class
{
    internal ComObjectHandle(
        T comObject,
        IReadOnlyList<IntPtr> activationContextHandles,
        IReadOnlyList<IntPtr> activationCookies)
    {
        if (comObject is null)
        {
            throw new ArgumentNullException(nameof(comObject));
        }

        if (activationContextHandles is null)
        {
            throw new ArgumentNullException(nameof(activationContextHandles));
        }

        if (activationCookies is null)
        {
            throw new ArgumentNullException(nameof(activationCookies));
        }

        ComObject = comObject;
        ActivationContextHandles = activationContextHandles;
        ActivationCookies = activationCookies;
    }

    /// <summary>
    /// Finalizes an instance of the <see cref="ComObjectHandle{T}"/> class.
    /// Acts as a diagnostic safety net when <see cref="Dispose"/> has not been called; it does not attempt
    /// any native cleanup because the finalizer thread may not match the COM apartment that owns the object.
    /// A leak is surfaced via <see cref="ComInteropEventSource"/> (for <c>dotnet-trace</c> / PerfView /
    /// <see cref="System.Diagnostics.Tracing.EventListener"/>) and, in <c>DEBUG</c> builds, via
    /// <see cref="Debug.Fail(string)"/>.
    /// </summary>
#pragma warning disable MA0055 // Do not use finalizer (intentional diagnostic safety net)
#pragma warning disable IDISP023 // Don't use reference types in finalizer context (diagnostic emission only, no cleanup)
    ~ComObjectHandle()
    {
        if (!IsReleased)
        {
            ComInteropEventSource.Log.HandleLeaked(typeof(T).Name);

            // Debug.Fail only when a debugger is attached — otherwise the trace-listener
            // pipeline can block on assertion dialogs / message boxes during process shutdown.
            // Production visibility is covered by the EventSource above.
            if (Debugger.IsAttached)
            {
                Debug.Fail(
                    "ComObjectHandle<" + typeof(T).Name + "> was not disposed. "
                    + "Call Dispose() or Executor.Free on the creating thread to release native activation contexts and COM references.");
            }
        }
    }
#pragma warning restore IDISP023
#pragma warning restore MA0055

    /// <summary>
    /// Gets the COM object instance. It becomes <see langword="null"/> after <see cref="Dispose"/> or
    /// <see cref="Executor.Free{T}(ComObjectHandle{T})"/> is called, and the handle should not be used afterwards.
    /// </summary>
    public T? ComObject { get; internal set; }

    /// <summary>
    /// Gets a value indicating whether the handle has already been released.
    /// </summary>
    public bool IsReleased { get; private set; }

    internal IReadOnlyList<IntPtr> ActivationContextHandles { get; private set; }

    internal IReadOnlyList<IntPtr> ActivationCookies { get; private set; }

    /// <summary>
    /// Releases the COM object and its associated activation contexts.
    /// Equivalent to <see cref="Executor.Free{T}(ComObjectHandle{T})"/>; the call is idempotent and must be
    /// performed on the thread that created the handle.
    /// </summary>
    /// <remarks>
    /// Native deactivation or <c>Marshal.FinalReleaseComObject</c> can fail. Because
    /// <see cref="IDisposable.Dispose"/> must not throw, a non-success <see cref="Result"/> returned by
    /// <see cref="Executor.Free{T}(ComObjectHandle{T})"/> is surfaced through
    /// <c>ComInteropEventSource.HandleReleaseFailed</c> (Event ID
    /// <c>ComInteropEventSource.HandleReleaseFailedEventId</c>), and via
    /// <see cref="Debug.Fail(string)"/> when a debugger is attached.
    /// </remarks>
    public void Dispose()
    {
        if (!IsReleased)
        {
            var result = Executor.Free(this);
            if (!result.Success)
            {
                var errorMessage = result.Exception?.Message ?? "unknown error";
                ComInteropEventSource.Log.HandleReleaseFailed(typeof(T).Name, errorMessage);
                if (Debugger.IsAttached)
                {
                    Debug.Fail(
                        "ComObjectHandle<" + typeof(T).Name + ">.Dispose() failed to release cleanly: " + errorMessage);
                }
            }
        }

        GC.SuppressFinalize(this);
    }

    internal void MarkReleased()
    {
        IsReleased = true;
        ComObject = default;
        ActivationContextHandles = Array.Empty<IntPtr>();
        ActivationCookies = Array.Empty<IntPtr>();
    }
}
