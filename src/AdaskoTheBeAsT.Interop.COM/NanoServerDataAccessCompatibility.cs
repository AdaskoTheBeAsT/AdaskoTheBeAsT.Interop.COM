#nullable enable
using System;
using System.IO;
#if NET10_0_OR_GREATER
using System.Runtime.Versioning;
#endif
using System.Threading;
using Microsoft.Win32;

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Seeds Microsoft Data Access Components (WDAC/MDAC) and Jet 4.0 OLE DB registry settings under
/// <c>HKEY_CURRENT_USER</c> strictly inside Windows Nano Server containers where <c>HKEY_LOCAL_MACHINE</c>
/// WDAC/Jet keys are absent. Never writes to <c>HKEY_CURRENT_USER</c> on Desktop Windows or Server Core hosts.
/// </summary>
#if NET10_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public static class NanoServerDataAccessCompatibility
{
#if NET10_0_OR_GREATER
    private static readonly Lock SyncRoot = new();
#else
    private static readonly object SyncRoot = new();
#endif
    private static bool _dataAccessSeeded;
    private static bool _jetSeeded;

    /// <summary>Checks for the x86 compatibility forwarder on Windows Nano Server.</summary>
    /// <returns>Whether the required Nano Server forwarder is present.</returns>
    public static bool IsAvailable()
    {
        if (Environment.OSVersion.Platform != PlatformID.Win32NT || Environment.Is64BitProcess)
        {
            return false;
        }

#pragma warning disable SCS0018, SEC0116 // OS SysWOW64 forwarder path check on Windows Nano Server.
        var systemDirectory = Environment.SystemDirectory;
        var windowsDirectory = Path.GetDirectoryName(systemDirectory);
        return windowsDirectory is not null
            && !File.Exists(Path.Combine(systemDirectory, "win32u.dll"))
            && File.Exists(Path.Combine(windowsDirectory, "SysWOW64", "forwarders", "ole32.dll"));
#pragma warning restore SCS0018, SEC0116
    }

    /// <summary>Seeds missing data access metadata for the current user on x86 Nano Server.</summary>
    public static void EnsureDataAccessSeeded()
    {
        if (Volatile.Read(ref _dataAccessSeeded)
            || !IsAvailable())
        {
            return;
        }

        lock (SyncRoot)
        {
            if (_dataAccessSeeded)
            {
                return;
            }

            using var hkcu32 = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32);
            SeedDataAccessVersionIfMissing(hkcu32);
            Volatile.Write(ref _dataAccessSeeded, true);
        }
    }

    /// <summary>Seeds missing Jet 4.0 and data access metadata for the current user on x86 Nano Server.</summary>
    /// <param name="sharedComDirectory">Trusted directory containing the private 32-bit Jet provider.</param>
    public static void EnsureJet40AndDataAccessSeeded(string sharedComDirectory)
    {
        if (string.IsNullOrWhiteSpace(sharedComDirectory))
        {
            throw new ArgumentException("A private provider directory is required.", nameof(sharedComDirectory));
        }

        if (Volatile.Read(ref _jetSeeded)
            || !IsAvailable())
        {
            return;
        }

        lock (SyncRoot)
        {
            if (_jetSeeded)
            {
                return;
            }

            using var hkcu32 = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32);
            SeedDataAccessVersionIfMissing(hkcu32);
            SeedJetEngineSettingsIfMissing(hkcu32);
            SeedJetOleDbProviderMetadataIfMissing(hkcu32, sharedComDirectory);
            Volatile.Write(ref _dataAccessSeeded, true);
            Volatile.Write(ref _jetSeeded, true);
        }
    }

    private static void SeedDataAccessVersionIfMissing(RegistryKey hkcu32)
    {
        using var machineDataAccess = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\DataAccess");
        if (machineDataAccess?.GetValue("FullInstallVer") is string { Length: > 0 })
        {
            return;
        }

        using var existingUser = hkcu32.OpenSubKey(@"Software\Microsoft\DataAccess");
        if (existingUser?.GetValue("FullInstallVer") is string { Length: > 0 })
        {
            return;
        }

        using var da = hkcu32.CreateSubKey(@"Software\Microsoft\DataAccess");
        SetIfMissing(da, "Full21Install", -1, RegistryValueKind.DWord);
        SetIfMissing(da, "Version", "6.3.9600.16384", RegistryValueKind.String);
        SetIfMissing(da, "FullInstallVer", "6.3.9600.16384", RegistryValueKind.String);
    }

    private static void SeedJetEngineSettingsIfMissing(RegistryKey hkcu32)
    {
        using var machineJet = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Jet\4.0\Engines");
        if (machineJet?.GetValue("SandBoxMode") is not null)
        {
            return;
        }

        using var eng = hkcu32.CreateSubKey(@"Software\Microsoft\Jet\4.0\Engines");
        SetIfMissing(eng, "SandBoxMode", 2, RegistryValueKind.DWord);
        SetIfMissing(eng, "SystemDB", "system.mdb", RegistryValueKind.String);
        using var j4 = hkcu32.CreateSubKey(@"Software\Microsoft\Jet\4.0\Engines\Jet 4.0");
        SetIfMissing(j4, "PageTimeout", 5000, RegistryValueKind.DWord);
        SetIfMissing(j4, "LockRetry", 20, RegistryValueKind.DWord);
        SetIfMissing(j4, "MaxBufferSize", 0, RegistryValueKind.DWord);
        SetIfMissing(j4, "Threads", 3, RegistryValueKind.DWord);
        SetIfMissing(j4, "UserCommitSync", "yes", RegistryValueKind.String);
        SetIfMissing(j4, "ImplicitCommitSync", "no", RegistryValueKind.String);
    }

    private static void SetIfMissing(RegistryKey? key, string? name, object value, RegistryValueKind kind)
    {
        if (key is not null && key.GetValue(name) is null)
        {
            key.SetValue(name, value, kind);
        }
    }

#pragma warning disable SCS0018, SEC0116 // Fixed Jet 4.0 provider DLL filenames inside validated shared COM directory.
    private static void SeedJetOleDbProviderMetadataIfMissing(
        RegistryKey hkcu32,
        string sharedComDirectory)
    {
        const string jetProviderClsid = "{DEE35070-506B-11CF-B1AA-00AA00B8DE95}";
        const string jetErrorClsid = "{DEE35071-506B-11CF-B1AA-00AA00B8DE95}";

        using var existingRoot = Registry.ClassesRoot.OpenSubKey($@"CLSID\{jetProviderClsid}\InprocServer32");
        if (existingRoot?.GetValue(name: null) is string { Length: > 0 })
        {
            return;
        }

        var fullJetDll = Path.GetFullPath(Path.Combine(sharedComDirectory, "msjetoledb40.dll"));
        if (!File.Exists(fullJetDll))
        {
            throw new FileNotFoundException("The private Jet 4.0 provider was not found.", fullJetDll);
        }

        using var ck = hkcu32.CreateSubKey($@"Software\Classes\CLSID\{jetProviderClsid}");
        SetIfMissing(ck, name: null, "Microsoft.Jet.OLEDB", RegistryValueKind.String);
        SetIfMissing(ck, "OLEDB_SERVICES", unchecked((int)0xfffffffe), RegistryValueKind.DWord);

        using var op = hkcu32.CreateSubKey($@"Software\Classes\CLSID\{jetProviderClsid}\OLE DB Provider");
        SetIfMissing(op, name: null, "Microsoft Jet 4.0 OLE DB Provider", RegistryValueKind.String);

        using var ext = hkcu32.CreateSubKey(
            $@"Software\Classes\CLSID\{jetProviderClsid}\ExtendedErrors\{jetErrorClsid}");
        SetIfMissing(ext, name: null, "Microsoft Jet 4.0 OLE DB Provider Error Lookup", RegistryValueKind.String);

        using var ips = hkcu32.CreateSubKey($@"Software\Classes\CLSID\{jetProviderClsid}\InprocServer32");
        SetIfMissing(ips, "ThreadingModel", "Both", RegistryValueKind.String);
        SetIfMissing(ips, name: null, fullJetDll, RegistryValueKind.String);
    }
#pragma warning restore SCS0018, SEC0116
}
