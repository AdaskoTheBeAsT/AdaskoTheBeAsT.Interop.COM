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
        using var listener = new CapturingEventListener(ComInteropEventSource.Log);

        ComInteropEventSource.Log.HandleLeaked("IFoo");

        listener.Events.Should().ContainSingle();
        var evt = listener.Events[0];
        evt.EventId.Should().Be(ComInteropEventSource.HandleLeakedEventId);
        evt.Level.Should().Be(EventLevel.Warning);
        evt.Payload.Should().NotBeNull();
        evt.Payload![0].Should().Be("IFoo");
    }

    [Fact]
    public void HandleLeakedShouldBeSilentWhenNoListenerIsAttached()
    {
        var act = () => ComInteropEventSource.Log.HandleLeaked("IBar");

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
