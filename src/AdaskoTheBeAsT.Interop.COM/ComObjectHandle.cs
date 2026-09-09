using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Represents a COM object together with the activation-context state that keeps it alive after
/// <see cref="Executor.Create{T}(string,string,Func{T})"/> or
/// <see cref="Executor.Create{T}(ICollection{ComPathDescriptor},Func{T})"/> returns.
/// Release it through <see cref="Dispose"/> (preferred, e.g. with a <see langword="using"/> statement) or
/// <see cref="Executor.Free{T}(ComObjectHandle{T})"/>.
/// </summary>
/// <typeparam name="T">The COM object type.</typeparam>
/// <remarks>
/// Instances must be used and released on the creating thread, in reverse creation order.
/// The handle exclusively owns the returned runtime callable wrapper (RCW). Release calls
/// <c>Marshal.FinalReleaseComObject</c>, invalidating every managed alias to that RCW.
/// Do not wrap a shared or borrowed COM object, or release the RCW independently.
/// Message pumping during release follows the policy selected when the handle was created.
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
    private readonly Thread _creatingThread = Thread.CurrentThread;
    private readonly uint _creatingNativeThreadId = NativeMethods.GetCurrentThreadId();

    internal ComObjectHandle(
        T comObject,
        IReadOnlyList<IntPtr> activationContextHandles,
        IReadOnlyList<IntPtr> activationCookies,
        bool pumpPendingMessages = true,
        ActivationContextApi? activationContextApi = null)
    {
        ThrowHelper.ThrowIfNull(comObject, nameof(comObject));
        ThrowHelper.ThrowIfNull(activationContextHandles, nameof(activationContextHandles));
        ThrowHelper.ThrowIfNull(activationCookies, nameof(activationCookies));

        ComObject = comObject;
        ActivationContextHandles = [.. activationContextHandles];
        ActivationCookies = [.. activationCookies];
        PumpPendingMessages = pumpPendingMessages;
        ContextApi = activationContextApi ?? ActivationContextApi.Instance;
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
    /// <see cref="Executor.Free{T}(ComObjectHandle{T})"/> releases it, and the object should not be used afterwards.
    /// A rejected wrong-thread or out-of-order release leaves the object unchanged.
    /// </summary>
    public T? ComObject { get; internal set; }

    /// <summary>
    /// Gets a value indicating whether the handle has already been released.
    /// </summary>
    public bool IsReleased { get; private set; }

    internal List<IntPtr> ActivationContextHandles { get; }

    internal List<IntPtr> ActivationCookies { get; }

    internal bool IsReleasing { get; set; }

    internal bool PumpPendingMessages { get; }

    internal ActivationContextApi ContextApi { get; }

    /// <summary>
    /// Releases the COM object and its associated activation contexts.
    /// Equivalent to <see cref="Executor.Free{T}(ComObjectHandle{T})"/>; the call is idempotent and must be
    /// performed on the creating thread, in reverse creation order.
    /// </summary>
    /// <remarks>
    /// Native deactivation or <c>Marshal.FinalReleaseComObject</c> can fail. Because
    /// this implementation of <see cref="IDisposable.Dispose"/> reports failures without throwing, a non-success <see cref="Result"/> returned by
    /// <see cref="Executor.Free{T}(ComObjectHandle{T})"/> is surfaced through
    /// <c>ComInteropEventSource.HandleReleaseFailed</c> (Event ID
    /// <c>ComInteropEventSource.HandleReleaseFailedEventId</c>), and via
    /// <see cref="Debug.Fail(string)"/> when a debugger is attached. A rejected release leaves the handle
    /// available for retry and does not suppress leak diagnostics.
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

        if (IsReleased)
        {
            GC.SuppressFinalize(this);
        }
    }

    internal void ValidateRelease()
    {
        if (!ReferenceEquals(Thread.CurrentThread, _creatingThread)
            || NativeMethods.GetCurrentThreadId() != _creatingNativeThreadId)
        {
            throw new InvalidOperationException("The COM handle must be released on the thread that created it.");
        }

        if (IsReleasing)
        {
            throw new InvalidOperationException("The COM handle is already being released.");
        }
    }

    internal void MarkReleased()
    {
        IsReleased = true;
        ComObject = default;
        ActivationContextHandles.Clear();
        ActivationCookies.Clear();
#pragma warning disable S3971 // Executor.Free is also an explicit release path, without calling Dispose.
        GC.SuppressFinalize(this);
#pragma warning restore S3971
    }
}
