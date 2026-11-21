using System.Runtime.InteropServices;
using AwesomeAssertions;
using Microsoft.Win32;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable IDISP007 // Tests explicitly verify scope teardown.
#pragma warning disable MA0051 // Comprehensive IErrorInfo and bidirectional HKCU/HKCR registry scenarios.
#pragma warning disable S8969 // Multi-targeted net472..net11.0 nullability annotations differ on BCL APIs.
#pragma warning disable SCS0018, SEC0116 // Controlled test temp directory paths.
public class ComErrorDiagnosticsAndBidirectionalRegistryTest
{
    [Fact]
    public void PeekConsumeAndCreateExceptionForHrShouldCaptureThreadErrorInfoWithoutLeaks()
    {
        TestWindow.OnStaThread(() =>
        {
            // Ensure clean initial thread error slot.
            _ = ComErrorDiagnostics.ConsumeCurrentErrorInfo();
            ComErrorDiagnostics.PeekCurrentErrorInfo().Should().BeNull();
            ComErrorDiagnostics.ConsumeCurrentErrorInfo().Should().BeNull();

            const int customHr = unchecked((int)0x800A03E9);
            var alphaIid = ProbeNativeInterop.AlphaIid;
            var iidBuffer = Marshal.AllocHGlobal(16);
            int returnedHr;
            try
            {
                Marshal.Copy(alphaIid.ToByteArray(), 0, iidBuffer, 16);
                returnedHr = ProbeNativeInterop.SafeNativeMethods.PopulateProbeComErrorInfo(
                    iidBuffer,
                    "Probe.CalcEngine",
                    "Psychrometric enthalpy out of range",
                    @"C:\Docs\ProbeHelp.chm",
                    42u,
                    customHr);
            }
            finally
            {
                Marshal.FreeHGlobal(iidBuffer);
            }

            returnedHr.Should().Be(customHr);

            // 1. Non-destructive PeekCurrentErrorInfo preserves the thread's IErrorInfo slot across repeated peeks.
            var peek1 = ComErrorDiagnostics.PeekCurrentErrorInfo();
            var peek2 = ComErrorDiagnostics.PeekCurrentErrorInfo();
            peek1.Should().NotBeNull();
            peek2.Should().NotBeNull();
            peek1!.InterfaceId.Should().Be(alphaIid);
            peek1.Source.Should().Be("Probe.CalcEngine");
            peek1.Description.Should().Be("Psychrometric enthalpy out of range");
            peek1.HelpFile.Should().Be(@"C:\Docs\ProbeHelp.chm");
            peek1.HelpContext.Should().Be(42u);
            peek1.ToString().Should().Be(peek1.ToSummaryString());
            peek1.ToSummaryString().Should().Be("Probe.CalcEngine: Psychrometric enthalpy out of range");
            peek2!.Description.Should().Be(peek1.Description);

            // 2. Non-consuming CreateExceptionForHr enriches COMException while leaving IErrorInfo intact.
            var nonConsumingEx = ComErrorDiagnostics.CreateExceptionForHr(customHr, consumeErrorInfo: false)
                .Should().BeOfType<COMException>().Which;
            nonConsumingEx.ErrorCode.Should().Be(customHr);
            nonConsumingEx.Source.Should().Be("Probe.CalcEngine");
            nonConsumingEx.HelpLink.Should().Be(@"C:\Docs\ProbeHelp.chm#42");
            nonConsumingEx.Message.Should().Contain("0x800A03E9");
            nonConsumingEx.Message.Should().Contain("Psychrometric enthalpy out of range");
            nonConsumingEx.Message.Should().Contain("Probe.CalcEngine");

            // 3. Consuming CreateExceptionForHr enriches COMException and clears the thread's IErrorInfo slot.
            var consumingEx = ComErrorDiagnostics.CreateExceptionForHr(customHr, consumeErrorInfo: true)
                .Should().BeOfType<COMException>().Which;
            consumingEx.ErrorCode.Should().Be(customHr);
            consumingEx.Message.Should().Contain("Psychrometric enthalpy out of range");

            ComErrorDiagnostics.PeekCurrentErrorInfo().Should().BeNull();
            ComErrorDiagnostics.ConsumeCurrentErrorInfo().Should().BeNull();

            // 4. Empty-slot CreateExceptionForHr falls back cleanly; non-failure HRESULT throws ArgumentOutOfRangeException.
            var fallbackEx = ComErrorDiagnostics.CreateExceptionForHr(NativeMethods.REGDB_E_CLASSNOTREG)
                .Should().BeOfType<COMException>().Which;
            fallbackEx.ErrorCode.Should().Be(NativeMethods.REGDB_E_CLASSNOTREG);

            var nonFailureAct = () => ComErrorDiagnostics.CreateExceptionForHr(NativeMethods.S_OK);
            nonFailureAct.Should().Throw<ArgumentOutOfRangeException>();
        });
    }

    [Fact]
    public void VisualBasicErrorNumberAndInterfaceErrorOffsetHelpersShouldDecodeExpectedHresults()
    {
        ComErrorDiagnostics.TryGetVisualBasicErrorNumber(unchecked((int)0x800A000D), out var vb13).Should().BeTrue();
        vb13.Should().Be(13);

        ComErrorDiagnostics.TryGetVisualBasicErrorNumber(unchecked((int)0x800A03E9), out var vb1001).Should().BeTrue();
        vb1001.Should().Be(1001);

        ComErrorDiagnostics.TryGetVisualBasicErrorNumber(NativeMethods.DISP_E_TYPEMISMATCH, out var dispMismatch).Should().BeTrue();
        dispMismatch.Should().Be(13);

        ComErrorDiagnostics.TryGetVisualBasicErrorNumber(NativeMethods.DISP_E_OVERFLOW, out var dispOverflow).Should().BeTrue();
        dispOverflow.Should().Be(6);

        ComErrorDiagnostics.TryGetVisualBasicErrorNumber(NativeMethods.DISP_E_BADINDEX, out var dispBadIndex).Should().BeTrue();
        dispBadIndex.Should().Be(9);

        ComErrorDiagnostics.TryGetVisualBasicErrorNumber(NativeMethods.DISP_E_DIVBYZERO, out var dispDivZero).Should().BeTrue();
        dispDivZero.Should().Be(11);

        ComErrorDiagnostics.TryGetVisualBasicErrorNumber(NativeMethods.E_OUTOFMEMORY, out var oomErr).Should().BeTrue();
        oomErr.Should().Be(7);

        ComErrorDiagnostics.TryGetVisualBasicErrorNumber(NativeMethods.S_OK, out _).Should().BeFalse();
        ComErrorDiagnostics.TryGetVisualBasicErrorNumber(unchecked((int)0x80004005), out _).Should().BeFalse();
        ComErrorDiagnostics.TryGetVisualBasicErrorNumber(unchecked((int)0x80040200), out _).Should().BeFalse();

        ComErrorDiagnostics.TryGetInterfaceErrorOffset(unchecked((int)0x80040205), out var defaultOffset).Should().BeTrue();
        defaultOffset.Should().Be(0x0205);

        ComErrorDiagnostics.TryGetInterfaceErrorOffset(
            unchecked((int)0x80040205),
            out var customOffset,
            baseCode: unchecked((int)0x80040200)).Should().BeTrue();
        customOffset.Should().Be(5);

        ComErrorDiagnostics.TryGetInterfaceErrorOffset(
            unchecked((int)0x800401FF),
            out _,
            baseCode: unchecked((int)0x80040200)).Should().BeFalse();
        ComErrorDiagnostics.TryGetInterfaceErrorOffset(unchecked((int)0x80070005), out _).Should().BeFalse();
        ComErrorDiagnostics.TryGetInterfaceErrorOffset(NativeMethods.S_OK, out _).Should().BeFalse();
    }

    [Fact]
    public void BidirectionalProgIdAndVersionIndependentProgIdLeasesShouldRoundTripAndCleanUpOnTeardown()
    {
        ProbeNativeInterop.RunWithCrossProcessRegistryGate(() => TestWindow.OnStaThread(() =>
        {
            var token = Guid.NewGuid().ToString("N").Substring(0, 12);
            var alphaVersioned = "Adasko.BidirAlpha." + token + ".1";
            var alphaVersionIndependent = "Adasko.BidirAlpha." + token;
            var betaVersioned = "Adasko.BidirBeta." + token + ".1";
            var betaVersionIndependent = "Adasko.BidirBeta." + token;

            var (workerDir, dllPath, _) = ProbeNativeInterop.CreateIsolatedProbeWorker("bidir_reg");
            try
            {
                var manifestPath = Path.Combine(workerDir, "bidirectional.manifest");
                File.WriteAllText(
                    manifestPath,
                    $"""
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
                      <file name="NativeLifetimeProbe.dll">
                        <comClass clsid="{ProbeNativeInterop.AlphaClsidString}" threadingModel="Apartment" progid="{alphaVersioned}" versionIndependentProgId="{alphaVersionIndependent}" />
                        <comClass clsid="{ProbeNativeInterop.BetaClsidString}" threadingModel="Apartment">
                          <progid>{betaVersioned}</progid>
                          <progid>{betaVersionIndependent}</progid>
                        </comClass>
                      </file>
                    </assembly>
                    """);

                var options = new ComActivationOptions
                {
                    ForceInProcessFallback = true,
                    RegisterCurrentUserProgIdsOnFallback = true,
                    UnloadModulesOnDeactivate = true,
                    PumpPendingMessages = false,
                };

                var act1 = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], options);
                act1.Success.Should().BeTrue();
                var scope1 = act1.Value!;

                var act2 = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], options);
                act2.Success.Should().BeTrue();
                var scope2 = act2.Value!;

                // Forward lookups via Win32 CLSIDFromProgID:
                NativeMethods.CLSIDFromProgID(alphaVersioned, out var resolvedAlpha1).Should().Be(0);
                resolvedAlpha1.Should().Be(ProbeNativeInterop.AlphaClsid);
                NativeMethods.CLSIDFromProgID(alphaVersionIndependent, out var resolvedAlphaVi).Should().Be(0);
                resolvedAlphaVi.Should().Be(ProbeNativeInterop.AlphaClsid);

                NativeMethods.CLSIDFromProgID(betaVersioned, out var resolvedBeta1).Should().Be(0);
                resolvedBeta1.Should().Be(ProbeNativeInterop.BetaClsid);
                NativeMethods.CLSIDFromProgID(betaVersionIndependent, out var resolvedBetaVi).Should().Be(0);
                resolvedBetaVi.Should().Be(ProbeNativeInterop.BetaClsid);

                // Reverse lookups via Win32 combase!ProgIDFromCLSID:
                ReadProgIdFromClsid(ProbeNativeInterop.AlphaClsid).Should().Be(alphaVersioned);
                ReadProgIdFromClsid(ProbeNativeInterop.BetaClsid).Should().Be(betaVersioned);

                ReadHkcuDefaultString($@"Software\Classes\CLSID\{ProbeNativeInterop.AlphaClsid:B}\VersionIndependentProgID")
                    .Should().Be(alphaVersionIndependent);
                ReadHkcuDefaultString($@"Software\Classes\CLSID\{ProbeNativeInterop.BetaClsid:B}\VersionIndependentProgID")
                    .Should().Be(betaVersionIndependent);

                // Deactivate inner overlapping scope first: both forward and reverse registrations stay alive!
                Executor.Deactivate(scope2).Success.Should().BeTrue();
                ReadProgIdFromClsid(ProbeNativeInterop.AlphaClsid).Should().Be(alphaVersioned);

                // Deactivate outer scope: all HKCU keys (forward and reverse) are removed cleanly!
                Executor.Deactivate(scope1).Success.Should().BeTrue();

                AssertHkcuSubKeyDoesNotExist($@"Software\Classes\{alphaVersioned}");
                AssertHkcuSubKeyDoesNotExist($@"Software\Classes\{alphaVersionIndependent}");
                AssertHkcuSubKeyDoesNotExist($@"Software\Classes\{betaVersioned}");
                AssertHkcuSubKeyDoesNotExist($@"Software\Classes\{betaVersionIndependent}");
                AssertHkcuSubKeyDoesNotExist($@"Software\Classes\CLSID\{ProbeNativeInterop.AlphaClsid:B}");
                AssertHkcuSubKeyDoesNotExist($@"Software\Classes\CLSID\{ProbeNativeInterop.BetaClsid:B}");
            }
            finally
            {
                CleanupTestRegistryKeys(alphaVersioned, alphaVersionIndependent, betaVersioned, betaVersionIndependent);
                ProbeNativeInterop.CleanupIsolatedProbeWorker(workerDir);
            }
        }));
    }

    [Fact]
    public void SkipCurrentUserRegistrationWhenPresentInClassesRootShouldSkipHkcuWritesAndPromoteWhenRequested()
    {
        ProbeNativeInterop.RunWithCrossProcessRegistryGate(() =>
        {
            var token = Guid.NewGuid().ToString("N").Substring(0, 12);
            var versionedProgId = "Adasko.HkcrSkip." + token + ".1";
            var versionIndependentProgId = "Adasko.HkcrSkip." + token;
            var (workerDir, dllPath, _) = ProbeNativeInterop.CreateIsolatedProbeWorker("hkcr_skip");
            try
            {
                var manifestPath = Path.Combine(workerDir, "hkcr_skip.manifest");
                File.WriteAllText(
                    manifestPath,
                    $"""
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
                      <file name="NativeLifetimeProbe.dll">
                        <comClass clsid="{ProbeNativeInterop.AlphaClsidString}" threadingModel="Apartment" progid="{versionedProgId}" versionIndependentProgId="{versionIndependentProgId}" />
                      </file>
                    </assembly>
                    """);

                ComManifestParser.TryParse(manifestPath, workerDir, dllPath, out var parsed, out var err).Should().BeTrue();
                err.Should().Be(0);

                var alphaBracedLower = ProbeNativeInterop.AlphaClsid.ToString("B").ToLowerInvariant();
                var alphaBracedUpper = ProbeNativeInterop.AlphaClsid.ToString("B").ToUpperInvariant();
                var simulatedHkcr = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    [$@"{versionedProgId}\CLSID|"] = alphaBracedLower,
                    [$@"{versionIndependentProgId}\CLSID|"] = alphaBracedLower,
                    [$@"CLSID\{alphaBracedUpper}\ProgID|"] = versionedProgId,
                    [$@"CLSID\{alphaBracedUpper}\VersionIndependentProgID|"] = versionIndependentProgId,
                };

                object? Reader(string subKey, string valueName)
                    => simulatedHkcr.TryGetValue(subKey + "|" + valueName, out var val) ? val : null;

                // 1. With skipWhenPresentInClassesRoot = true, matching HKCR entries prevent any HKCU creation!
                CurrentUserRegistryLeaseManager.TryAcquireFallbackRegistryLease(
                    parsed!,
                    registerProgIds: true,
                    registerTypeLibs: false,
                    skipWhenPresentInClassesRoot: true,
                    out var skippedLease,
                    Reader).Should().BeTrue();

                AssertHkcuSubKeyDoesNotExist($@"Software\Classes\{versionedProgId}");
                AssertHkcuSubKeyDoesNotExist($@"Software\Classes\{versionIndependentProgId}");

                // 2. Acquiring an overlapping lease with skipWhenPresentInClassesRoot = false promotes the lease to HKCU!
                CurrentUserRegistryLeaseManager.TryAcquireFallbackRegistryLease(
                    parsed!,
                    registerProgIds: true,
                    registerTypeLibs: false,
                    skipWhenPresentInClassesRoot: false,
                    out var forcedLease,
                    Reader).Should().BeTrue();

                ReadHkcuDefaultString($@"Software\Classes\{versionedProgId}\CLSID")
                    .Should().BeEquivalentTo(alphaBracedUpper);

                // 3. Releasing both leases leaves HKCU completely clean!
                skippedLease!.Rollback();
                ReadHkcuDefaultString($@"Software\Classes\{versionedProgId}\CLSID")
                    .Should().BeEquivalentTo(alphaBracedUpper);

                forcedLease!.Rollback();
                AssertHkcuSubKeyDoesNotExist($@"Software\Classes\{versionedProgId}");
                AssertHkcuSubKeyDoesNotExist($@"Software\Classes\{versionIndependentProgId}");
                AssertHkcuSubKeyDoesNotExist($@"Software\Classes\CLSID\{ProbeNativeInterop.AlphaClsid:B}");
            }
            finally
            {
                CleanupTestRegistryKeys(versionedProgId, versionIndependentProgId);
                ProbeNativeInterop.CleanupIsolatedProbeWorker(workerDir);
            }
        });
    }

    private static string ReadProgIdFromClsid(Guid clsid)
    {
        var local = clsid;
        var hr = NativeMethods.ProgIDFromCLSID(ref local, out var rawPtr);
        hr.Should().Be(0);
        rawPtr.Should().NotBe(IntPtr.Zero);
        try
        {
            return Marshal.PtrToStringUni(rawPtr)!;
        }
        finally
        {
            Marshal.FreeCoTaskMem(rawPtr);
        }
    }

    private static string? ReadHkcuDefaultString(string subKeyPath)
    {
        using var key = Registry.CurrentUser.OpenSubKey(subKeyPath, writable: false);
        return key?.GetValue(string.Empty) as string;
    }

    private static void AssertHkcuSubKeyDoesNotExist(string subKeyPath)
    {
        using var key = Registry.CurrentUser.OpenSubKey(subKeyPath, writable: false);
        key.Should().BeNull();
    }

    private static void CleanupTestRegistryKeys(params string[] progIds)
    {
        foreach (var progId in progIds)
        {
            Registry.CurrentUser.DeleteSubKeyTree($@"Software\Classes\{progId}", throwOnMissingSubKey: false);
        }

        Registry.CurrentUser.DeleteSubKeyTree($@"Software\Classes\CLSID\{ProbeNativeInterop.AlphaClsid:B}", throwOnMissingSubKey: false);
        Registry.CurrentUser.DeleteSubKeyTree($@"Software\Classes\CLSID\{ProbeNativeInterop.BetaClsid:B}", throwOnMissingSubKey: false);
    }
}
