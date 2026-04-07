using System;
using System.Collections.Generic;

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Represents a COM object together with the activation-context state that keeps it alive after
/// <see cref="Executor.Create{T}(string,string,Func{T})"/> or
/// <see cref="Executor.Create{T}(ICollection{ComPathDescriptor},Func{T})"/> returns.
/// Release it through <see cref="Executor.Free{T}(ComObjectHandle{T})"/>.
/// </summary>
/// <typeparam name="T">The COM object type.</typeparam>
/// <remarks>Instances are thread-affine and should be used and released on the same thread that created them.</remarks>
public sealed class ComObjectHandle<T>
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
    /// Gets the COM object instance. It becomes <see langword="null"/> after <see cref="Executor.Free{T}(ComObjectHandle{T})"/>
    /// is called, and the handle should not be used afterwards.
    /// </summary>
    public T? ComObject { get; internal set; }

    internal IReadOnlyList<IntPtr> ActivationContextHandles { get; private set; }

    internal IReadOnlyList<IntPtr> ActivationCookies { get; private set; }

    internal bool IsReleased { get; private set; }

    internal void MarkReleased()
    {
        IsReleased = true;
        ComObject = default;
        ActivationContextHandles = Array.Empty<IntPtr>();
        ActivationCookies = Array.Empty<IntPtr>();
    }
}
