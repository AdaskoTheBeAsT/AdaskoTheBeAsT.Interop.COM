using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using AwesomeAssertions;
using Microsoft.Win32;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

public sealed class NanoServerDataAccessCompatibilityTest : IDisposable
{
    private const string DataAccessPath = @"Software\Microsoft\DataAccess";
    private const string EnginesPath = @"Software\Microsoft\Jet\4.0\Engines";
    private const string ProviderPath = @"CLSID\{DEE35070-506B-11CF-B1AA-00AA00B8DE95}";
    private const string UserProviderPath = @"Software\Classes\" + ProviderPath;

    // Each hive is represented by a separate disposable subtree under HKCU, never a real registration.
    private readonly string _registryPath = $@"Software\AdaskoTheBeAsT.Interop.COM.Tests-{Guid.NewGuid():N}";

    [Fact]
    public void DesktopOrX64Process_DoesNotEnableNanoServerCompatibility()
    {
        if (Environment.Is64BitProcess || File.Exists(Path.Combine(Environment.SystemDirectory, "win32u.dll")))
        {
            NanoServerDataAccessCompatibility.IsAvailable().Should().BeFalse();
            NanoServerDataAccessCompatibility.EnsureDataAccessSeeded();
            NanoServerDataAccessCompatibility.EnsureJet40AndDataAccessSeeded("unused");
        }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" \t ")]
    public void EnsureJet40AndDataAccessSeeded_RejectsMissingDirectoryEvenWhenUnavailable(string? directory)
    {
        Action seed = () => NanoServerDataAccessCompatibility.EnsureJet40AndDataAccessSeeded(directory!);

        seed.Should().ThrowExactly<ArgumentException>().WithParameterName("sharedComDirectory");
    }

#pragma warning disable SCS0018, SEC0116 // Only fixed filenames under this test's unique temporary directory.
    [Fact]
    public void HasNanoServerForwarder_RequiresForwarderAndNoDesktopDll()
    {
        var windowsDirectory = CreateTemporaryDirectory();
        try
        {
            var systemDirectory = Directory.CreateDirectory(Path.Combine(windowsDirectory, "System32")).FullName;
            var forwarders = Directory.CreateDirectory(Path.Combine(windowsDirectory, "SysWOW64", "forwarders")).FullName;
            NanoServerDataAccessCompatibility.HasNanoServerForwarder(Path.GetPathRoot(windowsDirectory)!).Should().BeFalse();
            NanoServerDataAccessCompatibility.HasNanoServerForwarder(systemDirectory).Should().BeFalse();

            File.WriteAllBytes(Path.Combine(forwarders, "ole32.dll"), []);
            NanoServerDataAccessCompatibility.HasNanoServerForwarder(systemDirectory).Should().BeTrue();

            File.WriteAllBytes(Path.Combine(systemDirectory, "win32u.dll"), []);
            NanoServerDataAccessCompatibility.HasNanoServerForwarder(systemDirectory).Should().BeFalse();
        }
        finally
        {
            Directory.Delete(windowsDirectory, recursive: true);
        }
    }
#pragma warning restore SCS0018, SEC0116

    [Fact]
    public void UnavailableEnvironment_NeverOpensRegistry()
    {
        var probes = 0;
        var seeder = new NanoServerDataAccessCompatibility.Seeder(
            () =>
            {
                probes++;
                return false;
            },
            (_, _) => throw new InvalidOperationException("Unexpected registry access."));

        seeder.SeedDataAccess();
        seeder.SeedJet40AndDataAccess("unused");

        probes.Should().Be(2);
    }

    [Theory]
    [InlineData(RegistryHive.LocalMachine)]
    [InlineData(RegistryHive.CurrentUser)]
    public void EnsureDataAccessSeeded_LeavesAnExistingVersionAlone(RegistryHive hive)
    {
        using var existing = CreateKey(hive, DataAccessPath);
        existing.SetValue("FullInstallVer", "installed-version");

        CreateSeeder().SeedDataAccess();

        existing.GetValue("FullInstallVer").Should().Be("installed-version");
        existing.GetValue("Full21Install").Should().BeNull();
        existing.GetValue("Version").Should().BeNull();
        if (hive == RegistryHive.LocalMachine)
        {
            AssertKeyAbsent(DataAccessPath);
        }
    }

    [Fact]
    public void EnsureDataAccessSeeded_FillsMissingValuesWithoutOverwritingOtherValues()
    {
        using var machine = CreateKey(RegistryHive.LocalMachine, DataAccessPath);
        machine.SetValue("FullInstallVer", string.Empty);
        using var user = CreateKey(RegistryHive.CurrentUser, DataAccessPath);
        user.SetValue("Version", "custom");
        user.SetValue("Full21Install", 1, RegistryValueKind.DWord);

        CreateSeeder().SeedDataAccess();

        AssertUserValues(
            DataAccessPath,
            ("FullInstallVer", "6.3.9600.16384", RegistryValueKind.String),
            ("Version", "custom", RegistryValueKind.String),
            ("Full21Install", 1, RegistryValueKind.DWord));
        machine.GetValue("FullInstallVer").Should().Be(string.Empty);
    }

    [Fact]
    public void EnsureJet40AndDataAccessSeeded_SkipsInstalledEngineAndProvider()
    {
        using var machineEngines = CreateKey(RegistryHive.LocalMachine, EnginesPath);
        machineEngines.SetValue("SandBoxMode", 0, RegistryValueKind.DWord);
        using var registeredProvider = CreateKey(RegistryHive.ClassesRoot, ProviderPath + @"\InprocServer32");
        registeredProvider.SetValue(name: null, "installed-provider.dll");
        var seeder = CreateSeeder();

        seeder.SeedJet40AndDataAccess("unused");

        AssertKeyAbsent(EnginesPath);
        AssertKeyAbsent(UserProviderPath);
        AssertUserValues(DataAccessPath, ("FullInstallVer", "6.3.9600.16384", RegistryValueKind.String));
        machineEngines.GetValue("SandBoxMode").Should().Be(0);
        registeredProvider.GetValue(name: null).Should().Be("installed-provider.dll");
    }

#pragma warning disable SCS0018, SEC0116 // Only fixed filenames under this test's unique temporary directory.
    [Fact]
    public void EnsureJet40AndDataAccessSeeded_RetriesWhenProviderBecomesAvailable()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            var dll = Path.Combine(directory, "msjetoledb40.dll");
            var seeder = CreateSeeder();
            Action seed = () => seeder.SeedJet40AndDataAccess(Path.Combine(directory, "."));
            seed.Should().Throw<FileNotFoundException>().Which.FileName.Should().Be(dll);
            AssertKeyAbsent(UserProviderPath);
            using (var engines = CreateKey(RegistryHive.CurrentUser, EnginesPath + @"\Jet 4.0"))
            {
                engines.SetValue("Threads", 8, RegistryValueKind.DWord);
            }

            File.WriteAllBytes(dll, []);
            seed();
            File.Delete(dll);
            seed();

            AssertUserValues(
                DataAccessPath,
                ("Full21Install", -1, RegistryValueKind.DWord),
                ("Version", "6.3.9600.16384", RegistryValueKind.String),
                ("FullInstallVer", "6.3.9600.16384", RegistryValueKind.String));
            AssertUserValues(
                EnginesPath,
                ("SandBoxMode", 2, RegistryValueKind.DWord),
                ("SystemDB", "system.mdb", RegistryValueKind.String));
            AssertUserValues(
                EnginesPath + @"\Jet 4.0",
                ("PageTimeout", 5000, RegistryValueKind.DWord),
                ("LockRetry", 20, RegistryValueKind.DWord),
                ("MaxBufferSize", 0, RegistryValueKind.DWord),
                ("Threads", 8, RegistryValueKind.DWord),
                ("UserCommitSync", "yes", RegistryValueKind.String),
                ("ImplicitCommitSync", "no", RegistryValueKind.String));
            AssertUserValues(
                UserProviderPath,
                (null, "Microsoft.Jet.OLEDB", RegistryValueKind.String),
                ("OLEDB_SERVICES", -2, RegistryValueKind.DWord));
            AssertUserValues(
                UserProviderPath + @"\OLE DB Provider",
                (null, "Microsoft Jet 4.0 OLE DB Provider", RegistryValueKind.String));
            AssertUserValues(
                UserProviderPath + @"\ExtendedErrors\{DEE35071-506B-11CF-B1AA-00AA00B8DE95}",
                (null, "Microsoft Jet 4.0 OLE DB Provider Error Lookup", RegistryValueKind.String));
            AssertUserValues(
                UserProviderPath + @"\InprocServer32",
                ("ThreadingModel", "Both", RegistryValueKind.String),
                (null, dll, RegistryValueKind.String));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
#pragma warning restore SCS0018, SEC0116

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ConcurrentSeeding_OnlyOpensTheRegistryOnceAsync(bool includeJet)
    {
        using var registeredProvider = CreateKey(RegistryHive.ClassesRoot, ProviderPath + @"\InprocServer32");
        registeredProvider.SetValue(name: null, "installed-provider.dll");
        using var firstProbeStarted = new ManualResetEventSlim();
        using var resumeFirstProbe = new ManualResetEventSlim();
        var timeout = TimeSpan.FromSeconds(10);
        var probes = 0;
        var userOpens = 0;
        var seeder = new NanoServerDataAccessCompatibility.Seeder(
            () =>
            {
                if (Interlocked.Increment(ref probes) == 1)
                {
                    firstProbeStarted.Set();
                    if (!resumeFirstProbe.Wait(timeout, TestContext.Current.CancellationToken))
                    {
                        throw new TimeoutException("The first availability check was not resumed.");
                    }
                }

                return true;
            },
            (hive, view) =>
            {
                if (hive == RegistryHive.CurrentUser)
                {
                    Interlocked.Increment(ref userOpens);
                }

                return OpenBaseKey(hive, view);
            });
        Action seed = includeJet ? () => seeder.SeedJet40AndDataAccess("unused") : seeder.SeedDataAccess;

        var first = Task.Run(seed, TestContext.Current.CancellationToken);
        try
        {
            firstProbeStarted.Wait(timeout, TestContext.Current.CancellationToken).Should().BeTrue();
            seed();
        }
        finally
        {
            resumeFirstProbe.Set();
            await first;
        }

        seed();
        seeder.SeedDataAccess();
        userOpens.Should().Be(1);
        probes.Should().Be(2);
        AssertUserValues(DataAccessPath, ("FullInstallVer", "6.3.9600.16384", RegistryValueKind.String));
    }

    public void Dispose() => Registry.CurrentUser.DeleteSubKeyTree(_registryPath, throwOnMissingSubKey: false);

    private static string CreateTemporaryDirectory() =>
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"AdaskoDataAccess_{Guid.NewGuid():N}")).FullName;

    private NanoServerDataAccessCompatibility.Seeder CreateSeeder() => new(() => true, OpenBaseKey);

    private RegistryKey OpenBaseKey(RegistryHive hive, RegistryView view)
    {
        view.Should().Be(hive == RegistryHive.CurrentUser ? RegistryView.Registry32 : RegistryView.Default);
        return Registry.CurrentUser.CreateSubKey($@"{_registryPath}\{hive}");
    }

    private RegistryKey CreateKey(RegistryHive hive, string relativePath) =>
        Registry.CurrentUser.CreateSubKey($@"{_registryPath}\{hive}\{relativePath}");

    private void AssertKeyAbsent(string relativePath)
    {
        using var key = Registry.CurrentUser.OpenSubKey($@"{_registryPath}\{nameof(RegistryHive.CurrentUser)}\{relativePath}");
        key.Should().BeNull();
    }

    private void AssertUserValues(string relativePath, params (string? Name, object Value, RegistryValueKind Kind)[] values)
    {
        using var key = Registry.CurrentUser.OpenSubKey($@"{_registryPath}\{nameof(RegistryHive.CurrentUser)}\{relativePath}");
        var userKey = key.Should().BeOfType<RegistryKey>().Which;
        foreach (var (name, value, kind) in values)
        {
            userKey.GetValue(name).Should().Be(value);
            userKey.GetValueKind(name).Should().Be(kind);
        }
    }
}
