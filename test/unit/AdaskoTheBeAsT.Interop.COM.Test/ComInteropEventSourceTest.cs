using System.Collections.Generic;
using System.Diagnostics.Tracing;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

public class ComInteropEventSourceTest
{
    [Fact]
    public void ShouldHaveExpectedProviderName()
    {
        ComInteropEventSource.Log.Name.Should().Be(ComInteropEventSource.ProviderName);
        ComInteropEventSource.ProviderName.Should().Be("AdaskoTheBeAsT.Interop.COM");
    }

    [Fact]
    public void HandleLeakedShouldEmitWarningEventWhenListenerIsAttached()
    {
        // Use a unique sentinel type-name per test: ComInteropEventSource.Log is a process-wide
        // singleton, so finalizers from other tests (e.g. ComObjectHandleDisposeTest) can also
        // emit HandleLeaked events on the same listener at unpredictable times. Filtering by
        // payload keeps the assertion stable regardless of cross-test GC timing.
        const string sentinel = "Sentinel_HandleLeakedShouldEmit";
        using var listener = new CapturingEventListener(ComInteropEventSource.Log);

        ComInteropEventSource.Log.HandleLeaked(sentinel);

        var matches = listener.Events
            .FindAll(e => e.EventId == ComInteropEventSource.HandleLeakedEventId
                && e.Payload is { Count: > 0 }
                && string.Equals(e.Payload[0] as string, sentinel, System.StringComparison.Ordinal));
        matches.Should().ContainSingle();
        var evt = matches[0];
        evt.Level.Should().Be(EventLevel.Warning);
    }

    [Fact]
    public void HandleLeakedShouldBeSilentWhenNoListenerIsAttached()
    {
        var act = () => ComInteropEventSource.Log.HandleLeaked("IBar");

        act.Should().NotThrow();
    }

    [Fact]
    public void HandleReleaseFailedShouldEmitErrorEventWhenListenerIsAttached()
    {
        const string sentinel = "Sentinel_HandleReleaseFailedShouldEmit";
        using var listener = new CapturingEventListener(ComInteropEventSource.Log);

        ComInteropEventSource.Log.HandleReleaseFailed(sentinel, "native release failed");

        var matches = listener.Events
            .FindAll(e => e.EventId == ComInteropEventSource.HandleReleaseFailedEventId
                && e.Payload is { Count: > 1 }
                && string.Equals(e.Payload[0] as string, sentinel, System.StringComparison.Ordinal));
        matches.Should().ContainSingle();
        var evt = matches[0];
        evt.Level.Should().Be(EventLevel.Error);
        (evt.Payload![1] as string).Should().Be("native release failed");
    }

    [Fact]
    public void HandleReleaseFailedShouldBeSilentWhenNoListenerIsAttached()
    {
        var act = () => ComInteropEventSource.Log.HandleReleaseFailed("IQux", "err");

        act.Should().NotThrow();
    }

    private sealed class CapturingEventListener
        : EventListener
    {
        private readonly EventSource _target;

        public CapturingEventListener(EventSource target)
        {
            _target = target;
            EnableEvents(target, EventLevel.Verbose);
        }

        public List<EventWrittenEventArgs> Events { get; } = new List<EventWrittenEventArgs>();

        protected override void OnEventWritten(EventWrittenEventArgs eventData)
        {
            if (ReferenceEquals(eventData.EventSource, _target))
            {
                Events.Add(eventData);
            }
        }
    }
}
