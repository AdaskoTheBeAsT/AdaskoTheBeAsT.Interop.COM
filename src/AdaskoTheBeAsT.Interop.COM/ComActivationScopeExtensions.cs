using System;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1121 // Use built-in type alias
/// <summary>
/// Convenience extensions for direct <see cref="ComClassFactory"/> instantiation within an active
/// <see cref="IComActivationScope"/>.
/// </summary>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public static class ComActivationScopeExtensions
{
    /// <summary>
    /// Instantiates <paramref name="clsid"/> requesting a specific native <paramref name="interfaceId"/> directly
    /// from the loaded module at <paramref name="moduleIndex"/> in <see cref="IComActivationScope.LoadedModuleHandles"/>.
    /// </summary>
    /// <param name="scope">The active COM scope.</param>
    /// <param name="clsid">Coclass identifier (<c>CLSID</c>) to instantiate.</param>
    /// <param name="interfaceId">Specific native interface <c>IID</c> to request from <c>IClassFactory::CreateInstance</c>.</param>
    /// <param name="moduleIndex">Zero-based index into <see cref="IComActivationScope.LoadedModuleHandles"/>.</param>
    /// <param name="trackObject">When <see langword="true"/> (default), tracks the created RCW in <paramref name="scope"/>.</param>
    /// <returns>A <see cref="ComCreateResult{T}"/> containing the created RCW on success.</returns>
    public static ComCreateResult<object> CreateInstance(
        this IComActivationScope scope,
        Guid clsid,
        Guid interfaceId,
        int moduleIndex = 0,
        bool trackObject = true)
    {
        ThrowHelper.ThrowIfNull(scope, nameof(scope));
        return CreateInstanceWithOptionalInterface(scope, clsid, interfaceId, moduleIndex, trackObject);
    }

    /// <summary>
    /// Instantiates a strongly typed <typeparamref name="TInterface"/> wrapper for <paramref name="clsid"/> directly
    /// from an explicit <paramref name="moduleHandle"/> and optionally tracks it in <paramref name="scope"/>.
    /// </summary>
    /// <typeparam name="TInterface">Managed COM interface or coclass wrapper type.</typeparam>
    /// <param name="scope">The active COM scope that will track the created RCW when <paramref name="trackObject"/> is <see langword="true"/>.</param>
    /// <param name="moduleHandle">Explicit native <c>HMODULE</c> exporting <c>DllGetClassObject</c>.</param>
    /// <param name="clsid">Coclass identifier (<c>CLSID</c>) to instantiate.</param>
    /// <param name="trackObject">When <see langword="true"/> (default), tracks the created RCW in <paramref name="scope"/>.</param>
    /// <returns>A <see cref="ComCreateResult{T}"/> containing the typed RCW on success.</returns>
    public static ComCreateResult<TInterface> CreateInstance<TInterface>(
        this IComActivationScope scope,
        IntPtr moduleHandle,
        Guid clsid,
        bool trackObject = true)
        where TInterface : class
    {
        ThrowHelper.ThrowIfNull(scope, nameof(scope));
        var creation = ComClassFactory.CreateInstance<TInterface>(moduleHandle, clsid);
        if (!creation.Success || creation.Value is null || !trackObject)
        {
            return creation;
        }

        try
        {
            scope.TrackObject(creation.Value);
            return creation;
        }
        catch (Exception ex)
        {
            ReleaseTrackedCandidate(creation.Value);
            return new ComCreateResult<TInterface>
            {
                Success = false,
                Exception = ex,
            };
        }
    }

    internal static ComCreateResult<object> CreateInstanceWithOptionalInterface(
        IComActivationScope scope,
        Guid clsid,
        Guid? interfaceId,
        int moduleIndex,
        bool trackObject)
    {
        IntPtr moduleHandle;
        if (scope is ComActivationScope concreteScope)
        {
            if (!concreteScope.TryResolveScopeModuleHandle(moduleIndex, out moduleHandle, out var error))
            {
                return new ComCreateResult<object>
                {
                    Success = false,
                    Exception = error,
                };
            }
        }
        else if (!TryResolveGenericScopeModuleHandle(scope, moduleIndex, out moduleHandle, out var genericError))
        {
            return new ComCreateResult<object>
            {
                Success = false,
                Exception = genericError,
            };
        }

        var creation = ComClassFactory.CreateInstance(moduleHandle, clsid, interfaceId);
        if (!creation.Success || creation.Value is null || !trackObject)
        {
            return creation;
        }

        try
        {
            scope.TrackObject(creation.Value);
            return creation;
        }
        catch (Exception ex)
        {
            ReleaseTrackedCandidate(creation.Value);
            return new ComCreateResult<object>
            {
                Success = false,
                Exception = ex,
            };
        }
    }

    private static bool TryResolveGenericScopeModuleHandle(
        IComActivationScope scope,
        int moduleIndex,
        out IntPtr moduleHandle,
        out Exception? error)
    {
        moduleHandle = IntPtr.Zero;
        error = null;
        if (scope.IsReleased)
        {
            error = new ObjectDisposedException(nameof(IComActivationScope));
            return false;
        }

        if (moduleIndex < 0)
        {
            error = new ArgumentOutOfRangeException(nameof(moduleIndex), "Module index cannot be negative.");
            return false;
        }

        var handles = scope.LoadedModuleHandles;
        if (handles.Count == 0)
        {
            error = new InvalidOperationException(
                "No loaded module handles are available in this activation scope. "
                + "Set ComActivationOptions.PreloadModules = true (or ForceInProcessFallback = true) when activating the scope.");
            return false;
        }

        if (moduleIndex >= handles.Count)
        {
            error = new ArgumentOutOfRangeException(
                nameof(moduleIndex),
                "Module index must be less than LoadedModuleHandles.Count (" + handles.Count + ").");
            return false;
        }

        moduleHandle = handles[moduleIndex];
        return true;
    }

    private static void ReleaseTrackedCandidate(object instance)
    {
        if (Marshal.IsComObject(instance))
        {
            _ = Marshal.FinalReleaseComObject(instance);
        }
    }
}
#pragma warning restore SA1121
