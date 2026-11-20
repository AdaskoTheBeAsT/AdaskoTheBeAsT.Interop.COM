using AwesomeAssertions;
using Microsoft.Win32;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable IDISP007 // Tests explicitly verify scope teardown.
#pragma warning disable MA0051 // Comprehensive manifest & HKCU lease validation tests.
#pragma warning disable S8969 // Multi-targeted net472..net11.0 nullability annotations differ on BCL APIs.
#pragma warning disable SCS0018, SEC0116 // Controlled test temp directory paths.
public class ComManifestValidationAndFallbackOptionsTest
{
    [Theory]
    [InlineData("""<file name="NativeLifetimeProbe.dll"><comClass clsid="not-a-guid" /></file>""")]
    [InlineData("""<file name="NativeLifetimeProbe.dll"><comClass clsid="{00000000-0000-0000-0000-000000000000}" /></file>""")]
    [InlineData("""<file name="..\NativeLifetimeProbe.dll"><comClass clsid="{3a930e11-62b4-41d5-8c0a-510d92a70101}" /></file>""")]
    [InlineData("""<file name="\\server\share\NativeLifetimeProbe.dll"><comClass clsid="{3a930e11-62b4-41d5-8c0a-510d92a70101}" /></file>""")]
    [InlineData("""<file name="\\?\C:\NativeLifetimeProbe.dll"><comClass clsid="{3a930e11-62b4-41d5-8c0a-510d92a70101}" /></file>""")]
    [InlineData("""<file name="\Windows\System32\NativeLifetimeProbe.dll"><comClass clsid="{3a930e11-62b4-41d5-8c0a-510d92a70101}" /></file>""")]
    [InlineData("""<file name="C:NativeLifetimeProbe.dll"><comClass clsid="{3a930e11-62b4-41d5-8c0a-510d92a70101}" /></file>""")]
    [InlineData("""<file name="NativeLifetimeProbe.dll"><comClass clsid="{3a930e11-62b4-41d5-8c0a-510d92a70101}" threadingModel="Apartment" /><comClass clsid="{3a930e11-62b4-41d5-8c0a-510d92a70101}" threadingModel="Free" /></file>""")]
    [InlineData("""<file name="NativeLifetimeProbe.dll"><comClass clsid="{3a930e11-62b4-41d5-8c0a-510d92a70101}" progid="Adasko.DupProgId" /><comClass clsid="{3a930e11-62b4-41d5-8c0a-510d92a70201}" progid="Adasko.DupProgId" /></file>""")]
    [InlineData("""<file name="NativeLifetimeProbe.dll"><comClass clsid="{3a930e11-62b4-41d5-8c0a-510d92a70101}" progid="..\BadProgId" /></file>""")]
    [InlineData("""<file name="NativeLifetimeProbe.dll"><typelib tlbid="bad-tlbid" version="1.0" helpdir="" /></file>""")]
    public void InvalidManifestShouldBeRejectedBeforeLoadingModulesOrTouchingRegistry(string innerXml)
    {
        TestWindow.OnStaThread(() =>
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "AdaskoInvalidManifest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try
            {
                var dllPath = Path.Combine(tempDir, "NativeLifetimeProbe.dll");
                File.Copy(ProbeNativeInterop.ProbeDllSourcePath, dllPath, overwrite: true);
                var manifestPath = Path.Combine(tempDir, "invalid.manifest");
                File.WriteAllText(
                    manifestPath,
                    $"""
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0" description="Licensee: TopSecretVendorLicense-XYZ">
                      {innerXml}
                    </assembly>
                    """);

                var api = new ProbeNativeInterop.SimulatedFallbackActivationContextApi();
                var activation = Executor.Activate(
                    [new ComPathDescriptor(dllPath, manifestPath)],
                    new ComActivationOptions
                    {
                        UnloadModulesOnDeactivate = true,
                        RegisterCurrentUserProgIdsOnFallback = true,
                        PumpPendingMessages = false,
                    },
                    api);

                activation.Success.Should().BeFalse();
                activation.Exception.Should().NotBeNull();
                activation.Exception!.ToString().Should().NotContain("TopSecretVendorLicense-XYZ");
                NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);
            }
            finally
            {
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, recursive: true);
                }
            }
        });
    }

    [Fact]
    public void ProxyStubClassUnavailableShouldSkipOnlyForKnownProxyStubWhileOrdinaryUnavailableCoclassHardFails()
    {
        TestWindow.OnStaThread(() =>
        {
            var (workerDir, dllPath, _) = ProbeNativeInterop.CreateIsolatedProbeWorker("proxystub");
            try
            {
                var proxyManifest = Path.Combine(workerDir, "with_proxystub.manifest");
                File.WriteAllText(
                    proxyManifest,
                    $$"""
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
                      <file name="NativeLifetimeProbe.dll">
                        <comClass clsid="{{ProbeNativeInterop.AlphaClsidString}}" threadingModel="Apartment" description="Licensee: SecretCustomer-777" />
                        <comClass clsid="{{ProbeNativeInterop.ProxyStubClsidString}}" threadingModel="Both" description="PSFactoryBuffer" />
                        <comInterfaceProxyStub iid="{3a930e11-62b4-41d5-8c0a-510d92a70103}" name="IProbeSecondary" proxyStubClsid32="{{ProbeNativeInterop.ProxyStubClsidString}}" />
                      </file>
                    </assembly>
                    """);

                var api = new ProbeNativeInterop.SimulatedFallbackActivationContextApi();

                // 1. IgnoreUnavailableProxyStubClasses = true (default) succeeds for known PSFactoryBuffer.
                var okResult = Executor.Activate(
                    [new ComPathDescriptor(dllPath, proxyManifest)],
                    new ComActivationOptions { UnloadModulesOnDeactivate = true, PumpPendingMessages = false },
                    api);
                okResult.Success.Should().BeTrue();
                Executor.Deactivate(okResult.Value!).Success.Should().BeTrue();

                // 2. IgnoreUnavailableProxyStubClasses = false fails when proxy/stub CLSID is unavailable.
                var strictResult = Executor.Activate(
                    [new ComPathDescriptor(dllPath, proxyManifest)],
                    new ComActivationOptions
                    {
                        IgnoreUnavailableProxyStubClasses = false,
                        UnloadModulesOnDeactivate = true,
                        PumpPendingMessages = false,
                    },
                    api);
                strictResult.Success.Should().BeFalse();
                strictResult.Exception!.ToString().Should().NotContain("SecretCustomer-777");

                // 3. An ordinary unavailable coclass MUST ALWAYS fail even when IgnoreUnavailableProxyStubClasses = true!
                var ordinaryUnavailableManifest = Path.Combine(workerDir, "ordinary_unavailable.manifest");
                File.WriteAllText(
                    ordinaryUnavailableManifest,
                    $"""
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
                      <file name="NativeLifetimeProbe.dll">
                        <comClass clsid="{ProbeNativeInterop.AlphaClsidString}" threadingModel="Apartment" />
                        <comClass clsid="{ProbeNativeInterop.UnavailableClsidString}" threadingModel="Apartment" description="Licensee: NeverLeakThis" />
                      </file>
                    </assembly>
                    """);

                var ordinaryFail = Executor.Activate(
                    [new ComPathDescriptor(dllPath, ordinaryUnavailableManifest)],
                    new ComActivationOptions
                    {
                        IgnoreUnavailableProxyStubClasses = true,
                        UnloadModulesOnDeactivate = true,
                        PumpPendingMessages = false,
                    },
                    api);
                ordinaryFail.Success.Should().BeFalse();
                ordinaryFail.Exception!.ToString().Should().NotContain("NeverLeakThis");
                NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);
            }
            finally
            {
                if (Directory.Exists(workerDir))
                {
                    Directory.Delete(workerDir, recursive: true);
                }
            }
        });
    }

    [Fact]
    public void OptInCurrentUserProgIdRegistrationShouldReferenceCountPreservePriorValuesAndAvoidClobberingExternalEdits()
    {
        ProbeNativeInterop.RunWithCrossProcessRegistryGate(() => TestWindow.OnStaThread(() =>
        {
            var progId = "AdaskoTest.Probe." + Guid.NewGuid().ToString("N");
            var clsidSubKeyPath = @"Software\Classes\" + progId + @"\CLSID";
            var progIdKeyPath = @"Software\Classes\" + progId;
            var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("hkcu_progid", alphaProgId: progId);
            try
            {
                var api = new ProbeNativeInterop.SimulatedFallbackActivationContextApi();

                // 1. Default (RegisterCurrentUserProgIdsOnFallback = false) must NOT write to HKCU.
                var defaultAct = Executor.Activate(
                    [new ComPathDescriptor(dllPath, manifestPath)],
                    new ComActivationOptions { UnloadModulesOnDeactivate = true, PumpPendingMessages = false },
                    api);
                defaultAct.Success.Should().BeTrue();
                using (var key = Registry.CurrentUser.OpenSubKey(progIdKeyPath))
                {
                    key.Should().BeNull();
                }

                Executor.Deactivate(defaultAct.Value!).Success.Should().BeTrue();

                // 2. Opt-in with overlapping scopes reference-counts HKCU registration and resolves CLSIDFromProgID!
                var optInOptions = new ComActivationOptions
                {
                    RegisterCurrentUserProgIdsOnFallback = true,
                    UnloadModulesOnDeactivate = true,
                    PumpPendingMessages = false,
                };

                var scope1Res = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], optInOptions, api);
                scope1Res.Success.Should().BeTrue();
                var scope1 = scope1Res.Value!;

                var scope2Res = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], optInOptions, api);
                scope2Res.Success.Should().BeTrue();
                var scope2 = scope2Res.Value!;

                NativeMethods.CLSIDFromProgID(progId, out var resolvedClsid).Should().Be(0);
                resolvedClsid.Should().Be(ProbeNativeInterop.AlphaClsid);

                // Deactivating scope2 keeps the ProgID registered while scope1 is still active.
                Executor.Deactivate(scope2).Success.Should().BeTrue();
                using (var keyAfterScope2 = Registry.CurrentUser.OpenSubKey(clsidSubKeyPath))
                {
                    keyAfterScope2.Should().NotBeNull();
                }

                // Deactivating scope1 removes the owned transient ProgID key completely.
                Executor.Deactivate(scope1).Success.Should().BeTrue();
                using (var keyAfterScope1 = Registry.CurrentUser.OpenSubKey(progIdKeyPath))
                {
                    keyAfterScope1.Should().BeNull();
                }

                // 3. Pre-existing user registry value must be restored when the opt-in scope ends!
                const string priorValue = "{11111111-2222-3333-4444-555555555555}";
                using (var preKey = Registry.CurrentUser.CreateSubKey(clsidSubKeyPath))
                {
                    preKey!.SetValue(null, priorValue, RegistryValueKind.String);
                }

                var restoreScopeRes = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], optInOptions, api);
                restoreScopeRes.Success.Should().BeTrue();
                Executor.Deactivate(restoreScopeRes.Value!).Success.Should().BeTrue();

                using (var restoredKey = Registry.CurrentUser.OpenSubKey(clsidSubKeyPath))
                {
                    restoredKey.Should().NotBeNull();
                    restoredKey!.GetValue(null).Should().Be(priorValue);
                }

                // 4. External modification while scope is active must NOT be clobbered on rollback!
                const string externalEditValue = "{99999999-8888-7777-6666-555555555555}";
                var externalScopeRes = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], optInOptions, api);
                externalScopeRes.Success.Should().BeTrue();
                using (var editKey = Registry.CurrentUser.OpenSubKey(clsidSubKeyPath, writable: true))
                {
                    editKey!.SetValue(null, externalEditValue, RegistryValueKind.String);
                }

                Executor.Deactivate(externalScopeRes.Value!).Success.Should().BeTrue();
                using (var postExternalKey = Registry.CurrentUser.OpenSubKey(clsidSubKeyPath))
                {
                    postExternalKey!.GetValue(null).Should().Be(externalEditValue);
                }
            }
            finally
            {
                Registry.CurrentUser.DeleteSubKeyTree(progIdKeyPath, throwOnMissingSubKey: false);
                if (Directory.Exists(workerDir))
                {
                    Directory.Delete(workerDir, recursive: true);
                }
            }
        }));
    }

    [Fact]
    public void FallbackShouldLoadEmbeddedTypeLibAndRollbackOptionalCurrentUserTypeLibKey()
    {
        TestWindow.OnStaThread(() =>
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "AdaskoTypeLib_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var uniqueTlbGuid = Guid.NewGuid();
            var tlbKeyPath = @"Software\Classes\TypeLib\" + uniqueTlbGuid.ToString("B").ToUpperInvariant();
            try
            {
                var dllPath = Path.Combine(tempDir, "NativeCOM.dll");
                File.Copy(TestPaths.AssemblyPath, dllPath, overwrite: true);
                var manifestPath = Path.Combine(tempDir, "NativeCOM.typelib.manifest");
                File.WriteAllText(
                    manifestPath,
                    $$"""
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
                      <file name="NativeCOM.dll">
                        <comClass clsid="{03d0aedd-2013-418a-8b92-e57c30f2ab26}" threadingModel="Apartment" tlbid="{{uniqueTlbGuid:B}}" />
                        <typelib tlbid="{{uniqueTlbGuid:B}}" version="1.0" helpdir="" flags="HASDISKIMAGE" />
                      </file>
                    </assembly>
                    """);

                var api = new ProbeNativeInterop.SimulatedFallbackActivationContextApi();
                var options = new ComActivationOptions
                {
                    LoadTypeLibsOnFallback = true,
                    RegisterCurrentUserTypeLibsOnFallback = true,
                    UnloadModulesOnDeactivate = true,
                    PumpPendingMessages = false,
                };

                var act = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], options, api);
                act.Success.Should().BeTrue();

                using (var activeTlbKey = Registry.CurrentUser.OpenSubKey(tlbKeyPath))
                {
                    activeTlbKey.Should().NotBeNull();
                }

                var concat = new NativeCOM.StringConcatenatorClass();
                act.Value!.TrackObject(concat);
                concat.ConcatStrings("TypeLib ", "Loaded").Should().Be("TypeLib Loaded");

                Executor.Deactivate(act.Value).Success.Should().BeTrue();
                using (var removedTlbKey = Registry.CurrentUser.OpenSubKey(tlbKeyPath))
                {
                    removedTlbKey.Should().BeNull();
                }
            }
            finally
            {
                Registry.CurrentUser.DeleteSubKeyTree(tlbKeyPath, throwOnMissingSubKey: false);
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, recursive: true);
                }
            }
        });
    }
}
#pragma warning restore SCS0018, SEC0116
#pragma warning restore S8969
#pragma warning restore MA0051
#pragma warning restore IDISP007
#pragma warning restore SA1121
