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
    [InlineData(true, "activate")]
    [InlineData(false, "create")]
    public void Constructor_RejectsMissingFactory(bool omitActivator, string parameterName)
    {
        Action construct = () =>
        {
            using var owner = new ScopedComInstance<object>(
                reuse: false,
                omitActivator ? null! : () => throw new InvalidOperationException("Unexpected activation."),
                omitActivator ? _ => new object() : null!);
        };

        construct.Should().Throw<ArgumentNullException>().WithParameterName(parameterName);
    }

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
            _ => releases == 0 ? throw new InvalidOperationException("Test creation failure.") : expected);

        Action create = () =>
        {
            using var unexpected = owner.Acquire();
        };
        create.Should().Throw<InvalidOperationException>().WithMessage("Test creation failure.");
        releases.Should().Be(1);
        using var lease = owner.Acquire();
        lease.Instance.Should().BeSameAs(expected);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    [SuppressMessage("IDisposableAnalyzers", "IDISP017", Justification = "Explicit invalid disposal is under test.")]
    public void NestedLeases_RejectOutOfOrderDisposalAndAllowRecovery(bool reuse)
    {
        var releases = 0;
        using var owner = new ScopedComInstance<object>(
            reuse,
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
            ReferenceEquals(first.Instance, second.Instance).Should().Be(reuse);
        }

        releases.Should().Be(reuse ? 0 : 1);
    }

    [Fact]
    public void LifetimeOperations_RejectCallsFromAnotherThread()
    {
        var mock = new Mock<IComActivationScope>(MockBehavior.Strict);
        mock.Setup(s => s.Dispose());
        using var owner = new ScopedComInstance<object>(reuse: false, () => mock.Object, _ => new object());
        using var lease = owner.Acquire();
        Exception? acquireError = null;
        Exception? ownerDisposeError = null;
        Exception? instanceError = null;
        Exception? leaseDisposeError = null;
        var other = new Thread(() =>
        {
            acquireError = Record.Exception(() =>
            {
                using var unexpected = owner.Acquire();
            });
            ownerDisposeError = Record.Exception(owner.Dispose);
            instanceError = Record.Exception(() => _ = lease.Instance);
            leaseDisposeError = Record.Exception(lease.Dispose);
        });
        other.Start();
        other.Join();

        acquireError.Should().BeOfType<InvalidOperationException>().Which.Message.Should().Contain("creating thread");
        ownerDisposeError.Should().BeOfType<InvalidOperationException>().Which.Message.Should().Contain("creating thread");
        instanceError.Should().BeOfType<InvalidOperationException>().Which.Message.Should().Contain("creating thread");
        leaseDisposeError.Should().BeOfType<InvalidOperationException>().Which.Message.Should().Contain("creating thread");
        mock.Verify(s => s.Dispose(), Times.Never);
        lease.Instance.Should().NotBeNull();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    [SuppressMessage("IDisposableAnalyzers", "IDISP016", Justification = "Use after disposal is under test.")]
    [SuppressMessage("IDisposableAnalyzers", "IDISP017", Justification = "Repeated disposal is under test.")]
    public void Dispose_IsIdempotentAndPreventsFurtherUse(bool reuse)
    {
        var mock = new Mock<IComActivationScope>(MockBehavior.Strict);
        mock.Setup(s => s.Dispose());
        using var owner = new ScopedComInstance<object>(reuse, () => mock.Object, _ => new object());
        using var lease = owner.Acquire();

        lease.Dispose();
        lease.Dispose();
        Action readInstance = () => _ = lease.Instance;
        readInstance.Should().Throw<ObjectDisposedException>();

        owner.Dispose();
        owner.Dispose();
        Action acquire = () =>
        {
            using var unexpected = owner.Acquire();
        };
        acquire.Should().Throw<ObjectDisposedException>();
        mock.Verify(s => s.Dispose(), Times.Once);
    }

    [Fact]
    public void FailedActivationAndNullInstance_DoNotPreventRetry()
    {
        var activations = 0;
        var creations = 0;
        var expected = new object();
        var mock = new Mock<IComActivationScope>(MockBehavior.Strict);
        mock.Setup(s => s.Dispose());
        using var owner = new ScopedComInstance<object>(
            reuse: true,
            () => ++activations == 1 ? throw new InvalidOperationException("Test activation failure.") : mock.Object,
            _ => ++creations == 1 ? null! : expected);
        Action acquire = () =>
        {
            using var unexpected = owner.Acquire();
        };

        acquire.Should().Throw<InvalidOperationException>().WithMessage("Test activation failure.");
        mock.Verify(s => s.Dispose(), Times.Never);
        acquire.Should().Throw<InvalidOperationException>().WithMessage("The COM factory returned no instance.");
        mock.Verify(s => s.Dispose(), Times.Once);

        using var lease = owner.Acquire();
        lease.Instance.Should().BeSameAs(expected);
        activations.Should().Be(3);
        creations.Should().Be(2);
    }

    [Fact]
    [SuppressMessage("IDisposableAnalyzers", "IDISP017", Justification = "Explicit disposal triggers the callback under test.")]
    public void ActivationAndDeactivation_RejectReentrantLifetimeChanges()
    {
        var callbacks = 0;
        Action? duringTransition = null;
        var mock = new Mock<IComActivationScope>(MockBehavior.Strict);
        mock.Setup(s => s.Dispose()).Callback(() => duringTransition?.Invoke());
        using var owner = new ScopedComInstance<object>(
            reuse: true,
            () =>
            {
                duringTransition?.Invoke();
                return mock.Object;
            },
            _ => new object());
        Action verifyReentrancy = () =>
        {
            duringTransition = null;
            callbacks++;
            Action acquire = () =>
            {
                using var unexpected = owner.Acquire();
            };
            Action dispose = owner.Dispose;
            acquire.Should().Throw<InvalidOperationException>().WithMessage("*during activation or deactivation*");
            dispose.Should().Throw<InvalidOperationException>().WithMessage("*during activation or deactivation*");
        };

        duringTransition = verifyReentrancy;
        using (var lease = owner.Acquire())
        {
            lease.Instance.Should().NotBeNull();
        }

        duringTransition = verifyReentrancy;
        owner.Dispose();
        callbacks.Should().Be(2);
        mock.Verify(s => s.Dispose(), Times.Once);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    [SuppressMessage("IDisposableAnalyzers", "IDISP017", Justification = "Retrying disposal after a failure is under test.")]
    public void FailedScopeDisposal_CanBeRetried(bool reuse)
    {
        var expected = new object();
        var mock = new Mock<IComActivationScope>(MockBehavior.Strict);
        mock.SetupSequence(s => s.Dispose())
            .Throws(new InvalidOperationException("Test scope disposal failure."))
            .Pass();
        using var owner = new ScopedComInstance<object>(reuse, () => mock.Object, _ => expected);
        using var lease = owner.Acquire();
        if (reuse)
        {
            lease.Dispose();
            Action disposeOwner = owner.Dispose;
            disposeOwner.Should().Throw<InvalidOperationException>().WithMessage("Test scope disposal failure.");
            using var another = owner.Acquire();
            another.Instance.Should().BeSameAs(expected);
        }
        else
        {
            Action disposeLease = lease.Dispose;
            disposeLease.Should().Throw<InvalidOperationException>().WithMessage("Test scope disposal failure.");
            lease.Instance.Should().BeSameAs(expected);
            lease.Dispose();
        }

        owner.Dispose();
        mock.Verify(s => s.Dispose(), Times.Exactly(2));
    }
}
