using System;
using System.Collections.Generic;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1121 // Use built-in type alias
/// <summary>
/// Represents an active registration-free COM scope on the calling thread.
/// Multiple coclasses declared in the manifest(s) can be instantiated while the scope remains active.
/// </summary>
/// <remarks>
/// Scopes are thread-affine and must be used and disposed on the thread that activated them,
/// in strict last-in-first-out (LIFO) order relative to nested scopes, handles, and callbacks.
/// </remarks>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public interface IComActivationScope : IDisposable
{
    /// <summary>
    /// Gets a value indicating whether this scope has been deactivated and released.
    /// </summary>
    bool IsReleased { get; }

    /// <summary>
    /// Gets a value indicating whether this scope is backed by the in-process Nano Server
    /// class-factory fallback rather than native Win32 Side-by-Side (<c>CreateActCtxW</c>).
    /// </summary>
    bool IsUsingNanoServerFallback { get; }

    /// <summary>
    /// Gets the native module handles (<c>HMODULE</c>) currently loaded for the manifest DLLs in this scope.
    /// </summary>
    IReadOnlyList<IntPtr> LoadedModuleHandles { get; }

    /// <summary>
    /// Transfers exclusive runtime callable wrapper (RCW) ownership of <paramref name="comObject"/> to this scope
    /// so it is deterministically released via <c>Marshal.FinalReleaseComObject</c> before factory registrations
    /// and module leases are torn down.
    /// </summary>
    /// <param name="comObject">The COM object wrapper to release when the scope deactivates.</param>
    void TrackObject(object comObject);

    /// <summary>
    /// Instantiates <paramref name="clsid"/> directly from the loaded module at <paramref name="moduleIndex"/>
    /// in <see cref="LoadedModuleHandles"/> via <see cref="ComClassFactory"/> without consulting global COM class tables.
    /// </summary>
    /// <param name="clsid">Coclass identifier (<c>CLSID</c>) to instantiate.</param>
    /// <param name="moduleIndex">Zero-based index into <see cref="LoadedModuleHandles"/> (requires <see cref="ComActivationOptions.PreloadModules"/> in native WinSxS mode; automatically populated in fallback mode).</param>
    /// <param name="trackObject">When <see langword="true"/> (default), automatically registers the created RCW with <see cref="TrackObject"/>.</param>
    /// <returns>A <see cref="ComCreateResult{T}"/> containing the created RCW on success.</returns>
    ComCreateResult<object> CreateInstance(Guid clsid, int moduleIndex = 0, bool trackObject = true);

    /// <summary>
    /// Instantiates a strongly typed <typeparamref name="TInterface"/> wrapper for <paramref name="clsid"/> directly
    /// from the loaded module at <paramref name="moduleIndex"/> in <see cref="LoadedModuleHandles"/> via <see cref="ComClassFactory"/>.
    /// </summary>
    /// <typeparam name="TInterface">Managed COM interface or coclass wrapper type.</typeparam>
    /// <param name="clsid">Coclass identifier (<c>CLSID</c>) to instantiate.</param>
    /// <param name="moduleIndex">Zero-based index into <see cref="LoadedModuleHandles"/>.</param>
    /// <param name="trackObject">When <see langword="true"/> (default), automatically registers the created RCW with <see cref="TrackObject"/>.</param>
    /// <returns>A <see cref="ComCreateResult{T}"/> containing the typed RCW on success.</returns>
    ComCreateResult<TInterface> CreateInstance<TInterface>(Guid clsid, int moduleIndex = 0, bool trackObject = true)
        where TInterface : class;

    /// <summary>
    /// Pumps up to <paramref name="maxBatch"/> pending Windows messages on the owning thread.
    /// </summary>
    /// <param name="maxBatch">Maximum number of queued messages to dispatch in this call.</param>
    /// <returns>The outcome of the message pump operation.</returns>
    Result PumpPendingMessages(int maxBatch = 64);
}
#pragma warning restore SA1121
