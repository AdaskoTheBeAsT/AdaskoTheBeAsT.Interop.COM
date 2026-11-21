using System;
using System.Runtime.CompilerServices;
using System.Threading;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable MA0158 // Use System.Threading.Lock
#pragma warning disable SA1121 // Use built-in type alias
/// <summary>
/// Configuration options and startup preparation entry point for Windows Nano Server registration-free COM,
/// STA message-queue, Legacy OLE2, and <c>CLSID_DispenserManager</c> compatibility shims.
/// </summary>
public static class NanoServerOptions
{
    /// <summary>
    /// Environment variable that disables the in-process SxS class-factory fallback when set to <c>1</c> or <see langword="true"/>.
    /// </summary>
    public const string DisableSxsFallbackEnvironmentVariable = "ADASKO_NANOSERVER_DISABLE_SXS_FALLBACK";

    /// <summary>
    /// Environment variable that disables the library's in-process STA and message-queue shim when set to <c>1</c> or <see langword="true"/>.
    /// </summary>
    public const string DisableStaShimEnvironmentVariable = "ADASKO_NANOSERVER_DISABLE_STA_SHIM";

    /// <summary>
    /// Environment variable that disables the OS-level Nano Server STA shim when set to <c>1</c> or <see langword="true"/>.
    /// </summary>
    public const string DisableOsStaShimEnvironmentVariable = "NANOSERVER_DISABLE_STA_SHIM";

    /// <summary>
    /// Environment variable that opts in to the older-image Nano Server Legacy OLE2 IAT/delay-load shim when set to <c>1</c> or <see langword="true"/>.
    /// </summary>
    public const string EnableLegacyOle2ShimEnvironmentVariable = "ADASKO_NANOSERVER_ENABLE_LEGACY_OLE2_SHIM";

    /// <summary>
    /// Environment variable that opts in to the pure W^X native <c>CLSID_DispenserManager</c> fallback when set to <c>1</c> or <see langword="true"/>.
    /// </summary>
    public const string EnableDispenserManagerFallbackEnvironmentVariable = "ADASKO_NANOSERVER_ENABLE_DISPENSER_MANAGER_FALLBACK";

#if NET10_0_OR_GREATER
    private static readonly Lock SyncRoot = new();
#endif
#if NET472_OR_GREATER
    private static readonly object SyncRoot = new();
#endif

    private static int _initializedState;
    private static int _preparedState;
    private static int _enableSxsFallback = !IsEnvironmentFlagTrue(DisableSxsFallbackEnvironmentVariable) ? 1 : 0;
    private static int _enableStaShim = !IsEnvironmentFlagTrue(DisableStaShimEnvironmentVariable) ? 1 : 0;
    private static int _forceInProcessStaShim;
    private static int _enableLegacyOle2Shim = IsEnvironmentFlagTrue(EnableLegacyOle2ShimEnvironmentVariable) ? 1 : 0;
    private static int _enableDispenserManagerFallback = IsEnvironmentFlagTrue(EnableDispenserManagerFallbackEnvironmentVariable) ? 1 : 0;

    [ThreadStatic]
    private static bool _staThreadWarmed;

    /// <summary>
    /// Gets or sets a value indicating whether the manifest-driven in-process class-factory fallback
    /// (<c>DllGetClassObject</c> + <c>CoRegisterClassObject</c>) is enabled when Win32 Side-by-Side (<c>CreateActCtxW</c>)
    /// is unavailable on Windows Nano Server.
    /// </summary>
    public static bool EnableSxsFallback
    {
        get
        {
            EnsureInitialized(waitForHostThread: false);
            return Volatile.Read(ref _enableSxsFallback) != 0;
        }

        set
        {
            EnsureInitialized(waitForHostThread: false);
            lock (SyncRoot)
            {
                Volatile.Write(ref _enableSxsFallback, value ? 1 : 0);
            }
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether the library's in-process STA (<c>COINIT_APARTMENTTHREADED</c>)
    /// and user-mode message-queue shim is enabled on Windows Nano Server.
    /// </summary>
    public static bool EnableStaShim
    {
        get
        {
            EnsureInitialized(waitForHostThread: false);
            return Volatile.Read(ref _enableStaShim) != 0;
        }

        set
        {
            EnsureInitialized(waitForHostThread: false);
            lock (SyncRoot)
            {
                Volatile.Write(ref _enableStaShim, value ? 1 : 0);
                Volatile.Write(ref _preparedState, 0);
                NanoServerStaShim.RefreshState();
            }
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether the library's in-process STA and message-queue shim
    /// should be forced active even when an OS-level STA shim is present.
    /// </summary>
    public static bool ForceInProcessStaShim
    {
        get
        {
            EnsureInitialized(waitForHostThread: false);
            return Volatile.Read(ref _forceInProcessStaShim) != 0;
        }

        set
        {
            EnsureInitialized(waitForHostThread: false);
            lock (SyncRoot)
            {
                Volatile.Write(ref _forceInProcessStaShim, value ? 1 : 0);
                Volatile.Write(ref _preparedState, 0);
                NanoServerStaShim.RefreshState();
            }
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether the opt-in Legacy OLE2 compatibility shim
    /// (<c>OleInitialize</c>, <c>OleUninitialize</c>, <c>CoInitialize</c>, <c>OleRun</c>, <c>OleSaveToStream</c>, <c>OleLoadFromStream</c>)
    /// is enabled when neither the host OS nor <c>nanoserver_sta_shim.dll</c> natively exports functional OLE2 APIs.
    /// Defaults to <see langword="false"/> unless enabled via <see cref="EnableLegacyOle2ShimEnvironmentVariable"/>.
    /// </summary>
    public static bool EnableLegacyOle2Shim
    {
        get
        {
            EnsureInitialized(waitForHostThread: false);
            return Volatile.Read(ref _enableLegacyOle2Shim) != 0;
        }

        set
        {
            EnsureInitialized(waitForHostThread: false);
            lock (SyncRoot)
            {
                Volatile.Write(ref _enableLegacyOle2Shim, value ? 1 : 0);
                NanoServerOle2Shim.RefreshState();
            }
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether the opt-in pure W^X native <c>CLSID_DispenserManager</c>
    /// fallback is enabled when neither <c>comsvcs.dll</c> nor <c>nanoserver_sta_shim.dll</c> provides <c>CLSID_DispenserManager</c>.
    /// Defaults to <see langword="false"/> unless enabled via <see cref="EnableDispenserManagerFallbackEnvironmentVariable"/>.
    /// </summary>
    public static bool EnableDispenserManagerFallback
    {
        get
        {
            EnsureInitialized(waitForHostThread: false);
            return Volatile.Read(ref _enableDispenserManagerFallback) != 0;
        }

        set
        {
            EnsureInitialized(waitForHostThread: false);
            lock (SyncRoot)
            {
                Volatile.Write(ref _enableDispenserManagerFallback, value ? 1 : 0);
                NanoServerDispenserManager.RefreshState();
            }
        }
    }

    /// <summary>
    /// Gets a value indicating whether the library's in-process STA and message-queue shim is currently installed.
    /// </summary>
    public static bool IsInProcessStaShimActive
    {
        get
        {
            EnsureInitialized();
            return NanoServerStaShim.IsInstalled;
        }
    }

    /// <summary>
    /// Gets a value indicating whether the in-process Legacy OLE2 fallback provider is currently active.
    /// </summary>
    public static bool IsLegacyOle2ShimActive
    {
        get
        {
            EnsureInitialized(waitForHostThread: false);
            return NanoServerOle2Shim.IsInProcessFallbackActive;
        }
    }

    /// <summary>
    /// Gets a value indicating whether the in-process pure W^X <c>CLSID_DispenserManager</c> fallback provider is currently registered.
    /// </summary>
    public static bool IsDispenserManagerFallbackActive
    {
        get
        {
            EnsureInitialized(waitForHostThread: false);
            return NanoServerDispenserManager.IsFallbackRegistered;
        }
    }

    /// <summary>
    /// Idempotently prepares Nano Server forwarders, capability-queried native/in-process providers, and the
    /// main STA host thread on the process startup thread before concurrent <see cref="ApartmentState.STA"/>
    /// worker threads are started.
    /// </summary>
    /// <remarks>
    /// Call this method once during application startup after configuring any <see cref="NanoServerOptions"/>
    /// flags and before spawning STA worker pools so partially initialized STA threads do not contend on
    /// first-use module discovery or forwarder initialization.
    /// </remarks>
    public static void PrepareProcess()
    {
        if (Volatile.Read(ref _preparedState) != 0)
        {
            return;
        }

        EnsureInitialized(waitForHostThread: true);
        lock (SyncRoot)
        {
            NanoServerOle2Shim.RefreshState();
            NanoServerDispenserManager.RefreshState();
            Volatile.Write(ref _preparedState, 1);
        }
    }

    /// <summary>
    /// Ensures the current <see cref="ApartmentState.STA"/> thread has completed full Win32 OLE/STA window-class
    /// initialization (<c>OleMainThreadWndClass</c>) via a balanced <c>CoInitializeEx</c>/<c>CoUninitialize</c> pair.
    /// </summary>
    /// <remarks>
    /// CoreCLR initializes <see cref="ApartmentState.STA"/> threads via internal <c>RoInitialize</c> flags that defer
    /// registering <c>OleMainThreadWndClass</c> until an explicit <c>CoInitializeEx(0, COINIT_APARTMENTTHREADED)</c> runs.
    /// Calling this method once per STA worker thread before entering contended managed locks prevents concurrent
    /// <c>CoWaitForMultipleHandles</c>/<c>GetOrCreateSTAWindow</c> races on headless hosts. Repeated calls on the same
    /// STA thread are an O(1) thread-local no-op. Non-STA threads are left untouched.
    /// </remarks>
    public static void EnsureCurrentStaThreadInitialized()
    {
        if (_staThreadWarmed)
        {
            return;
        }

        EnsureInitialized(waitForHostThread: true);
        if (!IsCurrentThreadSta())
        {
            return;
        }

        var hr = NativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_APARTMENTTHREADED);
        if (hr >= 0)
        {
            NativeMethods.CoUninitialize();
        }

        _staThreadWarmed = true;
    }

    /// <summary>
    /// Re-evaluates the current STA shim configuration and environment variables and resets cached
    /// API-set presence flags in loaded system modules.
    /// </summary>
    public static void RefreshStaState()
    {
        EnsureInitialized(waitForHostThread: false);
        lock (SyncRoot)
        {
            Volatile.Write(ref _preparedState, 0);
            NanoServerStaShim.RefreshState();
            NanoServerOle2Shim.RefreshState();
            NanoServerDispenserManager.RefreshState();
            Volatile.Write(ref _preparedState, 1);
        }
    }

    /// <summary>
    /// Resets all Nano Server options to their environment variable defaults and refreshes provider state.
    /// </summary>
    public static void ResetToEnvironmentDefaults()
    {
        EnsureInitialized(waitForHostThread: false);
        lock (SyncRoot)
        {
            Volatile.Write(ref _enableSxsFallback, !IsEnvironmentFlagTrue(DisableSxsFallbackEnvironmentVariable) ? 1 : 0);
            Volatile.Write(ref _enableStaShim, !IsEnvironmentFlagTrue(DisableStaShimEnvironmentVariable) ? 1 : 0);
            Volatile.Write(ref _forceInProcessStaShim, 0);
            Volatile.Write(ref _enableLegacyOle2Shim, IsEnvironmentFlagTrue(EnableLegacyOle2ShimEnvironmentVariable) ? 1 : 0);
            Volatile.Write(ref _enableDispenserManagerFallback, IsEnvironmentFlagTrue(EnableDispenserManagerFallbackEnvironmentVariable) ? 1 : 0);
            Volatile.Write(ref _preparedState, 0);
            NanoServerStaShim.RefreshState();
            NanoServerOle2Shim.RefreshState();
            NanoServerDispenserManager.RefreshState();
        }
    }

    internal static bool ReadEnableStaShimRaw()
        => Volatile.Read(ref _enableStaShim) != 0;

    internal static bool ReadForceInProcessStaShimRaw()
        => Volatile.Read(ref _forceInProcessStaShim) != 0;

    internal static bool ReadEnableLegacyOle2ShimRaw()
        => Volatile.Read(ref _enableLegacyOle2Shim) != 0;

    internal static bool ReadEnableDispenserManagerFallbackRaw()
        => Volatile.Read(ref _enableDispenserManagerFallback) != 0;

    internal static void EnsureInitialized(bool waitForHostThread = true)
    {
        if (Volatile.Read(ref _initializedState) == 0)
        {
            lock (SyncRoot)
            {
                if (_initializedState == 0)
                {
                    InitializeUnderLock();
                }
            }
        }

        if (waitForHostThread)
        {
            NanoServerStaShim.WaitForMainStaHostReady();
        }
    }

    internal static bool IsEnvironmentFlagTrue(string variableName)
    {
        var value = Environment.GetEnvironmentVariable(variableName);
        if (value is not { Length: > 0 })
        {
            return false;
        }

        var trimmed = value.Trim();
        return string.Equals(trimmed, "1", StringComparison.OrdinalIgnoreCase)
            || string.Equals(trimmed, "true", StringComparison.OrdinalIgnoreCase)
            || string.Equals(trimmed, "yes", StringComparison.OrdinalIgnoreCase);
    }

#pragma warning disable CA2255
    [ModuleInitializer]
    internal static void InitializeModule() => EnsureInitialized(waitForHostThread: false);
#pragma warning restore CA2255

    private static void InitializeUnderLock()
    {
        Volatile.Write(ref _enableSxsFallback, !IsEnvironmentFlagTrue(DisableSxsFallbackEnvironmentVariable) ? 1 : 0);
        Volatile.Write(ref _enableStaShim, !IsEnvironmentFlagTrue(DisableStaShimEnvironmentVariable) ? 1 : 0);
        Volatile.Write(ref _enableLegacyOle2Shim, IsEnvironmentFlagTrue(EnableLegacyOle2ShimEnvironmentVariable) ? 1 : 0);
        Volatile.Write(ref _enableDispenserManagerFallback, IsEnvironmentFlagTrue(EnableDispenserManagerFallbackEnvironmentVariable) ? 1 : 0);
        Volatile.Write(ref _initializedState, 1);
        NanoServerStaShim.Initialize();
        NanoServerOle2Shim.RefreshState();
        NanoServerDispenserManager.RefreshState();
    }

    private static bool IsCurrentThreadSta()
    {
        if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
        {
            return true;
        }

        var hr = NativeMethods.CoGetApartmentType(out var aptType, out _);
        return hr >= 0 && aptType is NativeMethods.APTTYPE_STA or NativeMethods.APTTYPE_MAINSTA;
    }
}
#pragma warning restore SA1121
#pragma warning restore MA0158
