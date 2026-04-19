using System.Diagnostics.Tracing;

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Diagnostic <see cref="EventSource"/> that emits production-visible events from
/// <c>AdaskoTheBeAsT.Interop.COM</c>. Consumable by <c>dotnet-trace</c>, <c>dotnet-counters</c>,
/// PerfView, or any in-process <see cref="EventListener"/>.
/// </summary>
/// <remarks>
/// Subscribe by provider name <c>"AdaskoTheBeAsT.Interop.COM"</c>, for example:
/// <code>dotnet-trace collect --providers AdaskoTheBeAsT.Interop.COM --process-id &lt;pid&gt;</code>.
/// </remarks>
[EventSource(Name = ProviderName)]
internal sealed class ComInteropEventSource
    : EventSource
{
    /// <summary>
    /// Provider name of this <see cref="EventSource"/>. Subscribe to it with <c>dotnet-trace</c> or
    /// an in-process <see cref="EventListener"/>.
    /// </summary>
    public const string ProviderName = "AdaskoTheBeAsT.Interop.COM";

    /// <summary>
    /// Singleton instance. EventSources are intended to be used as process-wide singletons.
    /// </summary>
#pragma warning disable MA0069 // Non-constant static fields should not be visible
    public static readonly ComInteropEventSource Log = new ComInteropEventSource();
#pragma warning restore MA0069

    internal const int HandleLeakedEventId = 1;

    internal const int HandleReleaseFailedEventId = 2;

    private ComInteropEventSource()
    {
    }

    /// <summary>
    /// Emitted when a <c>ComObjectHandle&lt;T&gt;</c> is finalized without having been disposed or freed.
    /// Indicates a leak of native activation-context state and COM references.
    /// </summary>
    /// <param name="typeName">Short name of the generic argument <c>T</c>.</param>
    [Event(
        HandleLeakedEventId,
        Level = EventLevel.Warning,
        Message = "ComObjectHandle<{0}> was finalized without being disposed. Native activation-context state and COM references have leaked.")]
    public void HandleLeaked(string typeName)
    {
        if (IsEnabled())
        {
            WriteEvent(HandleLeakedEventId, typeName);
        }
    }

    /// <summary>
    /// Emitted when <c>ComObjectHandle&lt;T&gt;.Dispose()</c> invokes <c>Executor.Free</c> and the
    /// release returns a non-success <c>Result</c> (native deactivation, release, or COM final
    /// release failed). Callers that consume the handle through <c>using</c> would otherwise miss
    /// the failure entirely, because <c>Dispose</c> cannot throw without violating the pattern.
    /// </summary>
    /// <param name="typeName">Short name of the generic argument <c>T</c>.</param>
    /// <param name="errorMessage">Exception message produced by the failed release, or a sentinel value.</param>
    [Event(
        HandleReleaseFailedEventId,
        Level = EventLevel.Error,
        Message = "ComObjectHandle<{0}>.Dispose() failed to release cleanly: {1}")]
    public void HandleReleaseFailed(string typeName, string errorMessage)
    {
        if (IsEnabled())
        {
            WriteEvent(HandleReleaseFailedEventId, typeName, errorMessage);
        }
    }
}
