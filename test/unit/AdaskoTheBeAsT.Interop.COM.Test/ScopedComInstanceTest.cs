using System;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using AwesomeAssertions;
using Moq;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

public sealed class ScopedComInstanceTest
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AcquireAndDispose_TracksScopeLifetime(bool reuse)
    {
        var activations = 0;
        var releases = 0;
        using (var owner = new ScopedComInstance<object>(
            reuse,
            () =>
            {
                activations++;
                var mock = new Mock<IComActivationScope>(MockBehavior.Strict);
                mock.Setup(s => s.Dispose()).Callback(() => releases++);
                return mock.Object;
            },
            _ => new object()))
        {
            activations.Should().Be(0);
            object first;
            using (var lease = owner.Acquire())
            {
                first = lease.Instance;
                releases.Should().Be(0);
            }

            releases.Should().Be(reuse ? 0 : 1);
            using var second = owner.Acquire();
            ReferenceEquals(first, second.Instance).Should().Be(reuse);
            activations.Should().Be(reuse ? 1 : 2);
        }

        releases.Should().Be(reuse ? 1 : 2);
    }

    [Fact]
    public void FailedCreation_ReleasesScopeAndAllowsRetry()
    {
        var releases = 0;
        var mock = new Mock<IComActivationScope>(MockBehavior.Strict);
        mock.Setup(s => s.Dispose()).Callback(() => releases++);
        var expected = new object();
        using var owner = new ScopedComInstance<object>(
            reuse: true,
            () => mock.Object,
            _ => releases == 0 ? throw new InvalidOperationException("Test activation failure.") : expected);

        Action create = () =>
        {
            using var unexpected = owner.Acquire();
        };
        create.Should().Throw<InvalidOperationException>().WithMessage("Test activation failure.");
        releases.Should().Be(1);
        using var lease = owner.Acquire();
        lease.Instance.Should().BeSameAs(expected);
    }

    [Fact]
    [SuppressMessage("IDisposableAnalyzers", "IDISP017", Justification = "Explicit invalid disposal is under test.")]
    public void NestedLeases_RejectOutOfOrderDisposalAndAllowRecovery()
    {
        var releases = 0;
        using var owner = new ScopedComInstance<object>(
            reuse: false,
            () =>
            {
                var mock = new Mock<IComActivationScope>(MockBehavior.Strict);
                mock.Setup(s => s.Dispose()).Callback(() => releases++);
                return mock.Object;
            },
            _ => new object());
        using var first = owner.Acquire();
        using (var second = owner.Acquire())
        {
            Action disposeFirst = first.Dispose;
            Action disposeOwner = owner.Dispose;
            disposeFirst.Should().Throw<InvalidOperationException>();
            disposeOwner.Should().Throw<InvalidOperationException>();
            releases.Should().Be(0);
            first.Instance.Should().NotBeSameAs(second.Instance);
        }

        releases.Should().Be(1);
    }

    [Fact]
    public void LeaseCannotBeUsedFromASecondThread()
    {
        var mock = new Mock<IComActivationScope>(MockBehavior.Strict);
        mock.Setup(s => s.Dispose());
        using var owner = new ScopedComInstance<object>(reuse: false, () => mock.Object, _ => new object());
        using var lease = owner.Acquire();
        Exception? observed = null;
        var other = new Thread(() => observed = Record.Exception(() => _ = lease.Instance));
        other.Start();
        other.Join();

        observed.Should().BeOfType<InvalidOperationException>();
        lease.Instance.Should().NotBeNull();
    }

    [Fact]
    public void X64Process_DoesNotEnableX86Compatibility()
    {
        if (Environment.Is64BitProcess)
        {
            NanoServerDataAccessCompatibility.IsAvailable().Should().BeFalse();
            NanoServerDataAccessCompatibility.EnsureDataAccessSeeded();
            NanoServerDataAccessCompatibility.EnsureJet40AndDataAccessSeeded("unused");
        }
    }
}
