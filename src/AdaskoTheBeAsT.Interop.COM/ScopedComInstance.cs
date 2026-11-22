using System;
using System.Threading;
#if NET10_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Owns a lazily activated COM instance, either per acquisition or for the lifetime of the owner.
/// </summary>
/// <typeparam name="T">The COM interface or coclass wrapper.</typeparam>
/// <remarks>
/// Use and dispose all leases on the creating thread, in reverse acquisition order.
/// When nesting scopes across owners, the caller must also respect the activation scopes' disposal order.
/// </remarks>
#if NET10_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public sealed class ScopedComInstance<T> : IDisposable
    where T : class
{
    private readonly Thread _owner = Thread.CurrentThread;
    private readonly bool _reuse;
    private readonly Func<IComActivationScope> _activate;
    private readonly Func<IComActivationScope, T> _create;
    private IComActivationScope? _retainedScope;
    private T? _retainedInstance;
    private int _activeLeases;
    private bool _inScopeTransition;
    private bool _disposed;

    /// <summary>Initializes an instance owner without activating a scope.</summary>
    /// <param name="reuse">Whether to retain the first successful activation and instance.</param>
    /// <param name="activate">Creates a new activation scope owned by this object.</param>
    /// <param name="create">Creates and tracks an instance in the supplied scope.</param>
    public ScopedComInstance(bool reuse, Func<IComActivationScope> activate, Func<IComActivationScope, T> create)
    {
        ThrowHelper.ThrowIfNull(activate, nameof(activate));
        ThrowHelper.ThrowIfNull(create, nameof(create));
        _reuse = reuse;
        _activate = activate;
        _create = create;
    }

    /// <summary>Acquires an instance and a lease that must be disposed before the owner.</summary>
    /// <returns>A lease holding the instance for the current call.</returns>
    public Lease Acquire()
    {
        VerifyLifetimeOperation();
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(ScopedComInstance<T>));
        }

        if (_retainedInstance is { } retained)
        {
            return CreateLease(retained, scope: null);
        }

        _inScopeTransition = true;
        try
        {
            var scope = _activate();
            try
            {
                var instance = _create(scope) ?? throw new InvalidOperationException("The COM factory returned no instance.");
                var lease = CreateLease(instance, _reuse ? null : scope);
                if (_reuse)
                {
                    _retainedScope = scope;
                    _retainedInstance = instance;
                }

                return lease;
            }
            catch
            {
                scope.Dispose();
                throw;
            }
        }
        finally
        {
            _inScopeTransition = false;
        }
    }

    /// <summary>Releases any retained instance and deactivates its scope on the creating thread.</summary>
    /// <exception cref="InvalidOperationException">Called from the wrong thread or with active leases.</exception>
#pragma warning disable CA1065, S3877 // Reject invalid lifetime operations without tearing down an active COM scope.
    public void Dispose()
    {
        VerifyLifetimeOperation();
        if (_disposed)
        {
            return;
        }

        if (_activeLeases != 0)
        {
            throw new InvalidOperationException("Dispose active COM instance leases before disposing the owner.");
        }

        DisposeScope(_retainedScope);
        _retainedScope = null;
        _retainedInstance = null;
        _disposed = true;
    }
#pragma warning restore CA1065, S3877

    private Lease CreateLease(T instance, IComActivationScope? scope)
    {
        var lease = new Lease(this, instance, scope, _activeLeases + 1);
        _activeLeases++;
        return lease;
    }

    private void VerifyLifetimeOperation()
    {
        VerifyThread();
        if (_inScopeTransition)
        {
            throw new InvalidOperationException("Cannot change the COM instance lifetime during activation or deactivation.");
        }
    }

#pragma warning disable IDISP007 // The owner or lease owns this activation scope.
    private void DisposeScope(IComActivationScope? scope)
    {
        _inScopeTransition = true;
        try
        {
            scope?.Dispose();
        }
        finally
        {
            _inScopeTransition = false;
        }
    }
#pragma warning restore IDISP007

    private void VerifyThread()
    {
        if (_owner != Thread.CurrentThread)
        {
            throw new InvalidOperationException("The COM instance and its leases belong to the creating thread.");
        }
    }

    /// <summary>Provides access to an instance for a call on the owning thread.</summary>
    public sealed class Lease : IDisposable
    {
        private readonly ScopedComInstance<T> _owner;
        private readonly IComActivationScope? _scope;
        private readonly int _depth;
        private T? _instance;

        internal Lease(ScopedComInstance<T> owner, T instance, IComActivationScope? scope, int depth)
        {
            _owner = owner;
            _instance = instance;
            _scope = scope;
            _depth = depth;
        }

        /// <summary>Gets the instance while this lease is active. Do not retain it beyond the lease.</summary>
        public T Instance
        {
            get
            {
                _owner.VerifyThread();
                return _instance ?? throw new ObjectDisposedException(nameof(Lease));
            }
        }

        /// <summary>Releases this lease, including the transient scope if the instance was not retained.</summary>
        /// <exception cref="InvalidOperationException">Called from the wrong thread or out of order.</exception>
#pragma warning disable CA1065, S3877 // An out-of-order COM deactivation must be rejected, not silently ignored.
        public void Dispose()
        {
            _owner.VerifyLifetimeOperation();
            if (_instance is null)
            {
                return;
            }

            if (_owner._activeLeases != _depth)
            {
                throw new InvalidOperationException("COM instance leases must be disposed in reverse acquisition order.");
            }

            _owner.DisposeScope(_scope);
            _instance = null;
            _owner._activeLeases--;
        }
#pragma warning restore CA1065, S3877
    }
}
