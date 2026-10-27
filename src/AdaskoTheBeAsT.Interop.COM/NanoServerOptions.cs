using System;
using System.Runtime.CompilerServices;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable MA0158 // Use System.Threading.Lock
/// <summary>
/// Configuration options for Windows Nano Server registration-free COM and STA compatibility shims.
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

    private static readonly object SyncRoot = new();
    private static bool _initialized;
    private static bool _enableSxsFallback = !IsEnvironmentFlagTrue(DisableSxsFallbackEnvironmentVariable);
    private static bool _enableStaShim = !IsEnvironmentFlagTrue(DisableStaShimEnvironmentVariable);
    private static bool _forceInProcessStaShim;

    /// <summary>
    /// Gets or sets a value indicating whether the manifest-driven in-process class-factory fallback
    /// (<c>DllGetClassObject</c> + <c>CoRegisterClassObject</c>) is enabled when Win32 Side-by-Side (<c>CreateActCtxW</c>)
    /// is unavailable on Windows Nano Server.
    /// </summary>
    public static bool EnableSxsFallback
    {
        get
        {
            EnsureInitialized();
            lock (SyncRoot)
            {
                return _enableSxsFallback;
            }
        }

        set
        {
            EnsureInitialized();
            lock (SyncRoot)
            {
                _enableSxsFallback = value;
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
            EnsureInitialized();
            lock (SyncRoot)
            {
                return _enableStaShim;
            }
        }

        set
        {
            EnsureInitialized();
            lock (SyncRoot)
            {
                _enableStaShim = value;
                NanoServerStaShim.RefreshState();
            }
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether the library's in-process STA and message-queue shim
    /// should be forced active on Nano Server even when an OS-level STA shim is present.
    /// </summary>
    public static bool ForceInProcessStaShim
    {
        get
        {
            EnsureInitialized();
            lock (SyncRoot)
            {
                return _forceInProcessStaShim;
            }
        }

        set
        {
            EnsureInitialized();
            lock (SyncRoot)
            {
                _forceInProcessStaShim = value;
                NanoServerStaShim.RefreshState();
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
    /// Re-evaluates the current STA shim configuration and environment variables and resets cached
    /// API-set presence flags in loaded system modules.
    /// </summary>
    public static void RefreshStaState()
    {
        EnsureInitialized();
        lock (SyncRoot)
        {
            NanoServerStaShim.RefreshState();
        }
    }

    /// <summary>
    /// Resets all Nano Server options to their environment variable defaults and refreshes the STA shim state.
    /// </summary>
    public static void ResetToEnvironmentDefaults()
    {
        EnsureInitialized();
        lock (SyncRoot)
        {
            _enableSxsFallback = !IsEnvironmentFlagTrue(DisableSxsFallbackEnvironmentVariable);
            _enableStaShim = !IsEnvironmentFlagTrue(DisableStaShimEnvironmentVariable);
            _forceInProcessStaShim = false;
            NanoServerStaShim.RefreshState();
        }
    }

    internal static void EnsureInitialized(bool waitForHostThread = true)
    {
        lock (SyncRoot)
        {
            if (!_initialized)
            {
                _enableSxsFallback = !IsEnvironmentFlagTrue(DisableSxsFallbackEnvironmentVariable);
                _enableStaShim = !IsEnvironmentFlagTrue(DisableStaShimEnvironmentVariable);
                _initialized = true;
                NanoServerStaShim.Initialize();
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
}
#pragma warning restore MA0158
