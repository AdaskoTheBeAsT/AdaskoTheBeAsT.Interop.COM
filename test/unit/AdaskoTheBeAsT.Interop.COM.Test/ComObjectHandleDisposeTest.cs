using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Tracing;
using System.Linq;
using System.Runtime.CompilerServices;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

#pragma warning disable IDISP007 // Don't dispose injected (tests intentionally exercise Dispose)
#pragma warning disable IDISP016 // Don't use disposed instance (tests assert state after Dispose)
public class ComObjectHandleDisposeTest
{
#pragma warning disable IDISP005 // Return type should indicate disposability (constructor throws, no instance produced)
    [Fact]
    public void CtorShouldThrowWhenComObjectIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => new ComObjectHandle<object>(null!, new List<IntPtr>(), new List<IntPtr>()));

        ex.ParamName.Should().Be("comObject");
    }

    [Fact]
    public void CtorShouldThrowWhenActivationContextHandlesIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => new ComObjectHandle<object>(new object(), null!, new List<IntPtr>()));

        ex.ParamName.Should().Be("activationContextHandles");
    }

    [Fact]
    public void CtorShouldThrowWhenActivationCookiesIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => new ComObjectHandle<object>(new object(), new List<IntPtr>(), null!));

        ex.ParamName.Should().Be("activationCookies");
    }
#pragma warning restore IDISP005

#pragma warning disable S1215 // GC.Collect is intentional: exercising the diagnostic finalizer
    [Fact]
    public void FinalizerShouldEmitLeakEventOnEventSourceWhenNotDisposed()
    {
        // Replace Trace.Listeners with a silent listener so Debug.Fail inside the finalizer
        // cannot block the test runner with an assertion dialog.
        var original = new TraceListener[Trace.Listeners.Count];
        Trace.Listeners.CopyTo(original, 0);
        Trace.Listeners.Clear();
        try
        {
            using var capture = new CapturingEventListener(ComInteropEventSource.Log);

            AllocateAndAbandonHandle();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            capture.Events.Should().Contain(
                e => e.EventId == ComInteropEventSource.HandleLeakedEventId);
        }
        finally
        {
            foreach (var listener in original)
            {
                Trace.Listeners.Add(listener);
            }
        }
    }
#pragma warning restore S1215

    [Fact]
    public void DisposeShouldReleaseComObjectAndActivationContexts()
    {
        // Arrange
        var (comAssemblyPath, manifestPath) = GetPaths();
        var creation = Executor.Create(
            comAssemblyPath,
            manifestPath,
            () => new NativeCOM.StringConcatenatorClass());
        creation.Success.Should().BeTrue();
        var handle = creation.Value!;

        // Act
        handle.Dispose();

        // Assert
        handle.IsReleased.Should().BeTrue();
        handle.ComObject.Should().BeNull();
        handle.ActivationContextHandles.Should().BeEmpty();
        handle.ActivationCookies.Should().BeEmpty();
    }

    [Fact]
    public void DisposeShouldBeIdempotent()
    {
        // Arrange
        var (comAssemblyPath, manifestPath) = GetPaths();
        var creation = Executor.Create(
            comAssemblyPath,
            manifestPath,
            () => new NativeCOM.StringConcatenatorClass());
        var handle = creation.Value!;

        // Act
        handle.Dispose();
        var act = handle.Dispose;

        // Assert
        act.Should().NotThrow();
        handle.IsReleased.Should().BeTrue();
    }

    [Fact]
    public void UsingStatementShouldReleaseHandle()
    {
        // Arrange
        var (comAssemblyPath, manifestPath) = GetPaths();
        var creation = Executor.Create(
            comAssemblyPath,
            manifestPath,
            () => new NativeCOM.StringConcatenatorClass());
        var handle = creation.Value!;

        // Act
        using (handle)
        {
            handle.ComObject!.ConcatStrings("A", "B").Should().Be("AB");
        }

        // Assert
        handle.IsReleased.Should().BeTrue();
        handle.ComObject.Should().BeNull();
    }

#pragma warning disable IDISP004 // Don't ignore created IDisposable (intentional leak for finalizer test)
#pragma warning disable CA2000 // Dispose objects before losing scope (same reason)
    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void AllocateAndAbandonHandle()
    {
        // Bypass Executor.Create — no real native resources are allocated, so the finalizer is safe to run.
        _ = new ComObjectHandle<string>(
            "fake-com-object",
            new List<IntPtr>(),
            new List<IntPtr>());
    }
#pragma warning restore CA2000
#pragma warning restore IDISP004

    private static (string ComAssemblyPath, string ManifestPath) GetPaths()
    {
        var currentPath = AppContext.BaseDirectory;
        var comAssemblyPath = Path.GetFullPath(Path.Combine(currentPath, "NativeCOM.dll"));
        var manifestPath = Path.GetFullPath(Path.Combine(currentPath, "NativeCOM.manifest"));

        File.Exists(comAssemblyPath).Should().BeTrue();
        File.Exists(manifestPath).Should().BeTrue();

        return (comAssemblyPath, manifestPath);
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
#pragma warning restore IDISP016
#pragma warning restore IDISP007
