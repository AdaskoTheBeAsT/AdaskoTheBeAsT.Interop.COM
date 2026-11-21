using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable SA1202 // Helper placement in test class
#pragma warning disable SA1204 // Static helper placement in test class
#pragma warning disable IDISP007 // Tests explicitly verify scope teardown.
#pragma warning disable IDISP016, IDISP017 // Explicit double-Dispose idempotency verification.
#pragma warning disable MA0051 // Comprehensive native integration scenarios for COM-13..COM-17.
#pragma warning disable MA0158 // Multi-TFM test uses System.Object monitor across net472..net11.0.
#pragma warning disable S1215 // Explicit GC collection verifies deterministic RCW finalization before module unload.
#pragma warning disable SCS0018, SEC0116 // Controlled test temp directory paths.
public class StaPreWarmingLocaleQuirksDiagnosticsAndManifestRewriterTest
{
    private const uint PolishLocaleId = 0x0415u;

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate void ExportedVoidProcDelegate();

    [Fact]
    public void EightFreshStaThreadsContendingOnLockAndActivateSimultaneouslyShouldSucceedWithBalancedApartment()
    {
        const int workerCount = 8;
        var cancellationToken = TestContext.Current.CancellationToken;
        var workers = new (string WorkerDirectory, string DllPath, string ManifestPath)[workerCount];
        for (var i = 0; i < workerCount; i++)
        {
            workers[i] = ProbeNativeInterop.CreateIsolatedProbeWorker("sta_warm_" + i.ToString(CultureInfo.InvariantCulture));
        }

        try
        {
            using var startBarrier = new Barrier(workerCount);
            var syncGate = new object();
            var failures = new List<ExceptionDispatchInfo>();
            var threads = new Thread[workerCount];

            for (var index = 0; index < workerCount; index++)
            {
                var (_, dllPath, manifestPath) = workers[index];
                threads[index] = new Thread(() =>
                {
                    try
                    {
                        startBarrier.SignalAndWait(TimeSpan.FromSeconds(20), cancellationToken).Should().BeTrue();

                        NanoServerOptions.EnsureCurrentStaThreadInitialized();
                        NanoServerOptions.EnsureCurrentStaThreadInitialized(); // O(1) thread-local fast path

                        lock (syncGate)
                        {
                            Thread.SpinWait(250);
                        }

                        NativeMethods.CoGetApartmentType(out var aptBefore, out _).Should().Be(0);
                        (aptBefore is NativeMethods.APTTYPE_STA or NativeMethods.APTTYPE_MAINSTA).Should().BeTrue();

                        var act = Executor.Activate(
                            [new ComPathDescriptor(dllPath, manifestPath)],
                            new ComActivationOptions
                            {
                                PreloadModules = true,
                                UnloadModulesOnDeactivate = true,
                                PumpPendingMessages = true,
                            });
                        act.Success.Should().BeTrue();
                        var scope = act.Value!;

                        ComModuleDiagnostics.VerifyRegisteredInprocServer(
                            ProbeNativeInterop.AlphaClsid,
                            scope.LoadedModuleHandles[0],
                            dllPath)
                            .Status.Should().Be(ComModuleVerificationStatus.Match);

                        Executor.Deactivate(scope).Success.Should().BeTrue();

                        NativeMethods.CoGetApartmentType(out var aptAfter, out _).Should().Be(0);
                        aptAfter.Should().Be(aptBefore);
                        NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);
                    }
                    catch (Exception ex)
                    {
                        lock (failures)
                        {
                            failures.Add(ExceptionDispatchInfo.Capture(ex));
                        }

                        try
                        {
                            startBarrier.RemoveParticipant();
                        }
                        catch (InvalidOperationException barrierEx)
                        {
                            _ = barrierEx;
                        }
                    }
                })
                {
                    IsBackground = true,
                };
                threads[index].SetApartmentState(ApartmentState.STA);
            }

            foreach (var thread in threads)
            {
                thread.Start();
            }

            foreach (var thread in threads)
            {
                thread.Join(TimeSpan.FromSeconds(30)).Should().BeTrue();
            }

            if (failures.Count > 0)
            {
                failures[0].Throw();
            }

            // Also confirm EnsureCurrentStaThreadInitialized is a safe no-op on an MTA thread.
            ExceptionDispatchInfo? mtaFailure = null;
            var mtaThread = new Thread(() =>
            {
                try
                {
                    NanoServerOptions.EnsureCurrentStaThreadInitialized();
                    NativeMethods.CoGetApartmentType(out var aptType, out _).Should().Be(0);
                    aptType.Should().Be(NativeMethods.APTTYPE_MTA);
                }
                catch (Exception ex)
                {
                    mtaFailure = ExceptionDispatchInfo.Capture(ex);
                }
            })
            {
                IsBackground = true,
            };
            mtaThread.SetApartmentState(ApartmentState.MTA);
            mtaThread.Start();
            mtaThread.Join(TimeSpan.FromSeconds(15)).Should().BeTrue();
            mtaFailure?.Throw();
        }
        finally
        {
            for (var i = 0; i < workerCount; i++)
            {
                ProbeNativeInterop.CleanupIsolatedProbeWorker(workers[i].WorkerDirectory);
            }
        }
    }

    [Fact]
    public void ComAutomationLocaleScopeShouldGuardOleAut32ParsingRestoreLocaleAndRejectWrongThreadDispose()
    {
        TestWindow.OnStaThread(() =>
        {
            var initialLcid = NativeMethods.GetThreadLocale();
            var initialCulture = CultureInfo.CurrentCulture;
            var initialUiCulture = CultureInfo.CurrentUICulture;
            var polishCulture = new CultureInfo("pl-PL");

            try
            {
                NativeMethods.SetThreadLocale(PolishLocaleId).Should().BeTrue();
                CultureInfo.CurrentCulture = polishCulture;
                CultureInfo.CurrentUICulture = polishCulture;

                // Under Polish (0x0415, comma decimal separator), "12,34" parses to 12.34 while "12.34" either fails or misparses!
                ProbeNativeInterop.SafeNativeMethods.ParseDoubleWithThreadLocale(
                    "12,34",
                    0u,
                    out var commaParsed,
                    out var observedPolishLcid).Should().Be(0);
                observedPolishLcid.Should().Be(PolishLocaleId);
                BitConverter.DoubleToInt64Bits(commaParsed).Should().Be(BitConverter.DoubleToInt64Bits(12.34d));

                var dotOutsideHr = ProbeNativeInterop.SafeNativeMethods.ParseDoubleWithThreadLocale(
                    "12.34",
                    0u,
                    out var dotOutsideValue,
                    out _);
                (dotOutsideHr < 0 || BitConverter.DoubleToInt64Bits(dotOutsideValue) != BitConverter.DoubleToInt64Bits(12.34d))
                    .Should().BeTrue();

                // Enter US English (0x0409) + InvariantCulture scope:
                var scope = ComAutomationLocaleScope.Enter();
                scope.OwnerThreadId.Should().Be(Environment.CurrentManagedThreadId);
                scope.PreviousThreadLocaleId.Should().Be(PolishLocaleId);
                CultureInfo.CurrentCulture.Should().Be(CultureInfo.InvariantCulture);
                CultureInfo.CurrentUICulture.Should().Be(CultureInfo.InvariantCulture);

                ProbeNativeInterop.SafeNativeMethods.ParseDoubleWithThreadLocale(
                    "12.34",
                    0u,
                    out var dotInsideValue,
                    out var observedUsLcid).Should().Be(0);
                observedUsLcid.Should().Be(ComAutomationLocaleScope.UsEnglishLocaleId);
                BitConverter.DoubleToInt64Bits(dotInsideValue).Should().Be(BitConverter.DoubleToInt64Bits(12.34d));

                ProbeNativeInterop.SafeNativeMethods.FormatDoubleWithThreadLocale(
                    12.34d,
                    0u,
                    out var formattedBstr,
                    out _).Should().Be(0);
                try
                {
                    Marshal.PtrToStringBSTR(formattedBstr).Should().Be("12.34");
                }
                finally
                {
                    if (formattedBstr != IntPtr.Zero)
                    {
                        Marshal.FreeBSTR(formattedBstr);
                    }
                }

                // Wrong-thread Dispose() throws InvalidOperationException without mutating either thread's locale!
                Exception? wrongThreadException = null;
                uint helperBefore = 0;
                uint helperAfter = 0;
                var helper = new Thread(() =>
                {
                    _ = NativeMethods.SetThreadLocale(ComAutomationLocaleScope.UkEnglishLocaleId);
                    helperBefore = NativeMethods.GetThreadLocale();
                    try
                    {
                        scope.Dispose();
                    }
                    catch (Exception ex)
                    {
                        wrongThreadException = ex;
                    }

                    helperAfter = NativeMethods.GetThreadLocale();
                })
                {
                    IsBackground = true,
                };
                helper.Start();
                helper.Join(TimeSpan.FromSeconds(15)).Should().BeTrue();

                wrongThreadException.Should().BeOfType<InvalidOperationException>();
                helperAfter.Should().Be(helperBefore);
                NativeMethods.GetThreadLocale().Should().Be(ComAutomationLocaleScope.UsEnglishLocaleId);

                // Owner-thread Dispose() restores 0x0415 and pl-PL, and repeated Dispose() is idempotent.
                scope.Dispose();
                scope.Dispose();

                NativeMethods.GetThreadLocale().Should().Be(PolishLocaleId);
                CultureInfo.CurrentCulture.Name.Should().Be("pl-PL");
                CultureInfo.CurrentUICulture.Name.Should().Be("pl-PL");

                // Also verify EnterInvariant() (0x007F) and UkEnglishLocaleId (0x0809).
                using (ComAutomationLocaleScope.EnterInvariant())
                {
                    NativeMethods.GetThreadLocale().Should().Be(ComAutomationLocaleScope.InvariantLocaleId);
                }

                NativeMethods.GetThreadLocale().Should().Be(PolishLocaleId);
            }
            finally
            {
                CultureInfo.CurrentCulture = initialCulture;
                CultureInfo.CurrentUICulture = initialUiCulture;
                _ = NativeMethods.SetThreadLocale(initialLcid);
            }
        });
    }

    [Fact]
    public void SuppressUninitializeImportsShouldKeepStaApartmentAndActiveRcwsIntactWhileEnforcingWxAndRejectingSystemModules()
    {
        var (unpatchedDir, unpatchedDllPath, _) = ProbeNativeInterop.CreateIsolatedProbeWorker("uninit_unpatched");
        var (patchedDir, patchedDllPath, patchedManifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("uninit_patched");
        try
        {
            // 1. Prove the real native pathology on an unpatched worker DLL: an extra CoUninitialize tears down a 1-refcount STA apartment
            // (dropping the thread out of STA into either CO_E_NOTINITIALIZED or the process-wide implicit MTA)!
            var unpatchedStaSurvived = true;
            ExceptionDispatchInfo? unpatchedThreadFailure = null;
            var rawThread = new Thread(() =>
            {
                var rawModule = IntPtr.Zero;
                try
                {
                    NativeMethods.CoGetApartmentType(out var aptBefore, out _).Should().Be(0);
                    (aptBefore is NativeMethods.APTTYPE_STA or NativeMethods.APTTYPE_MAINSTA).Should().BeTrue();

                    rawModule = NativeMethods.LoadLibraryEx(unpatchedDllPath, IntPtr.Zero, NativeMethods.LOAD_WITH_ALTERED_SEARCH_PATH);
                    rawModule.Should().NotBe(IntPtr.Zero);

                    var proc = NativeMethods.GetProcAddress(rawModule, "InvokeImportedCoUninitializeTwice");
                    proc.Should().NotBe(IntPtr.Zero);
                    var fn = Marshal.GetDelegateForFunctionPointer<ExportedVoidProcDelegate>(proc);
                    fn();

                    var postHr = NativeMethods.CoGetApartmentType(out var aptAfter, out _);
                    unpatchedStaSurvived = postHr == 0
                        && (aptAfter is NativeMethods.APTTYPE_STA or NativeMethods.APTTYPE_MAINSTA);
                }
                catch (Exception ex)
                {
                    unpatchedThreadFailure = ExceptionDispatchInfo.Capture(ex);
                }
                finally
                {
                    if (rawModule != IntPtr.Zero)
                    {
                        _ = NativeMethods.FreeLibrary(rawModule);
                    }
                }
            })
            {
                IsBackground = true,
            };
            rawThread.SetApartmentState(ApartmentState.STA);
            rawThread.Start();
            rawThread.Join(TimeSpan.FromSeconds(15)).Should().BeTrue();
            unpatchedThreadFailure?.Throw();
            unpatchedStaSurvived.Should().BeFalse();

            // 2. Prove on an identical STA thread that SuppressUninitializeImports keeps the STA apartment alive!
            var patchedStaSurvived = false;
            ExceptionDispatchInfo? patchedRawFailure = null;
            var patchedRawThread = new Thread(() =>
            {
                var rawModule = IntPtr.Zero;
                try
                {
                    NativeMethods.CoGetApartmentType(out var aptBefore, out _).Should().Be(0);
                    (aptBefore is NativeMethods.APTTYPE_STA or NativeMethods.APTTYPE_MAINSTA).Should().BeTrue();

                    rawModule = NativeMethods.LoadLibraryEx(patchedDllPath, IntPtr.Zero, NativeMethods.LOAD_WITH_ALTERED_SEARCH_PATH);
                    rawModule.Should().NotBe(IntPtr.Zero);

                    ComModuleQuirks.SuppressUninitializeImports(rawModule).Should().BeGreaterThanOrEqualTo(2);

                    var coProc = NativeMethods.GetProcAddress(rawModule, "InvokeImportedCoUninitializeTwice");
                    coProc.Should().NotBe(IntPtr.Zero);
                    Marshal.GetDelegateForFunctionPointer<ExportedVoidProcDelegate>(coProc)();

                    var oleProc = NativeMethods.GetProcAddress(rawModule, "InvokeImportedOleUninitializeTwice");
                    oleProc.Should().NotBe(IntPtr.Zero);
                    Marshal.GetDelegateForFunctionPointer<ExportedVoidProcDelegate>(oleProc)();

                    var postHr = NativeMethods.CoGetApartmentType(out var aptAfter, out _);
                    patchedStaSurvived = postHr == 0
                        && (aptAfter is NativeMethods.APTTYPE_STA or NativeMethods.APTTYPE_MAINSTA);

                    _ = ComModuleQuirks.RestoreUninitializeImports(rawModule);
                }
                catch (Exception ex)
                {
                    patchedRawFailure = ExceptionDispatchInfo.Capture(ex);
                }
                finally
                {
                    if (rawModule != IntPtr.Zero)
                    {
                        _ = NativeMethods.FreeLibrary(rawModule);
                        NanoServerOle2Shim.OnModuleUnloaded(rawModule);
                    }
                }
            })
            {
                IsBackground = true,
            };
            patchedRawThread.SetApartmentState(ApartmentState.STA);
            patchedRawThread.Start();
            patchedRawThread.Join(TimeSpan.FromSeconds(15)).Should().BeTrue();
            patchedRawFailure?.Throw();
            patchedStaSurvived.Should().BeTrue();

            // 3. Now prove ComModuleQuirks.SuppressUninitializeImports protects Executor activation scopes, active RCWs, and W^X!
            TestWindow.OnStaThread(() =>
            {
                // Reject zero/invalid handles and protected OS/COM/CLR modules!
                var zeroAct = () => ComModuleQuirks.SuppressUninitializeImports(IntPtr.Zero);
                zeroAct.Should().Throw<ArgumentException>();

                var combaseHandle = NativeMethods.GetModuleHandleW("combase.dll");
                combaseHandle.Should().NotBe(IntPtr.Zero);
                var combaseAct = () => ComModuleQuirks.SuppressUninitializeImports(combaseHandle);
                combaseAct.Should().Throw<InvalidOperationException>();

                var ole32Handle = NativeMethods.GetModuleHandleW("ole32.dll");
                if (ole32Handle != IntPtr.Zero)
                {
                    var ole32Act = () => ComModuleQuirks.SuppressUninitializeImports(ole32Handle);
                    ole32Act.Should().Throw<InvalidOperationException>();
                }

                var activation = Executor.Activate(
                    [new ComPathDescriptor(patchedDllPath, patchedManifestPath)],
                    new ComActivationOptions
                    {
                        PreloadModules = true,
                        UnloadModulesOnDeactivate = true,
                        PumpPendingMessages = false,
                    });
                activation.Success.Should().BeTrue();
                var scope = activation.Value!;
                var workerHandle = scope.LoadedModuleHandles[0];

                var patchedCount = ComModuleQuirks.SuppressUninitializeImports(workerHandle);

                // Both CoUninitialize and OleUninitialize imports should be patched.
                patchedCount.Should().BeGreaterThanOrEqualTo(2);
                ComModuleQuirks.SuppressUninitializeImports(workerHandle).Should().Be(patchedCount); // Idempotent

                // Verify strict W^X on the native no-op stub page: allocated RW, transitioned to RX, never RWX!
                var stubPage = NanoServerOle2Shim.UninitializeNoOpStubAddress;
                stubPage.Should().NotBe(IntPtr.Zero);
                var mbiSize = new UIntPtr((uint)Marshal.SizeOf<NativeMethods.MEMORY_BASIC_INFORMATION>());
                NativeMethods.VirtualQuery(stubPage, out var mbi, mbiSize).Should().NotBe(UIntPtr.Zero);
                mbi.AllocationProtect.Should().Be(NativeMethods.PAGE_READWRITE);
                mbi.Protect.Should().Be(NativeMethods.PAGE_EXECUTE_READ);

                ExerciseSuppressedUninitializeWithLiveRcw(scope, workerHandle);

                ComModuleQuirks.RestoreUninitializeImports(workerHandle).Should().Be(patchedCount);
                ComModuleQuirks.RestoreUninitializeImports(workerHandle).Should().Be(0);

                Executor.Deactivate(scope).Success.Should().BeTrue();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                _ = Executor.TryUnloadUnusedModules();
                NativeMethods.GetModuleHandleW(patchedDllPath).Should().Be(IntPtr.Zero);
            });
        }
        finally
        {
            ProbeNativeInterop.CleanupIsolatedProbeWorker(patchedDir);
            ProbeNativeInterop.CleanupIsolatedProbeWorker(unpatchedDir);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void VerifyRegisteredInprocServerShouldInspectModuleWithoutRcwAndAllowImmediateUnloadWithoutGc(bool forceFallback)
    {
        TestWindow.OnStaThread(() =>
        {
            var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker(
                forceFallback ? "rcwfree_fb" : "rcwfree_sxs");
            try
            {
                var act = Executor.Activate(
                    [new ComPathDescriptor(dllPath, manifestPath)],
                    new ComActivationOptions
                    {
                        ForceInProcessFallback = forceFallback,
                        PreloadModules = true,
                        UnloadModulesOnDeactivate = true,
                        PumpPendingMessages = false,
                    });
                act.Success.Should().BeTrue();
                var scope = act.Value!;

                var resolvedPath = ComModuleDiagnostics.GetRegisteredInprocServerPath(ProbeNativeInterop.AlphaClsid);
                resolvedPath.Should().Be(Path.GetFullPath(dllPath));

                var matchResult = ComModuleDiagnostics.VerifyRegisteredInprocServer(
                    ProbeNativeInterop.AlphaClsid,
                    IntPtr.Zero,
                    expectedDllPath: dllPath,
                    interfaceId: ProbeNativeInterop.AlphaIid,
                    methodSlotIndex: 3);
                matchResult.Status.Should().Be(ComModuleVerificationStatus.Match);
                matchResult.ModuleHandle.Should().Be(scope.LoadedModuleHandles[0]);

                var mismatchResult = ComModuleDiagnostics.InspectRegisteredInprocServer(
                    ProbeNativeInterop.AlphaClsid,
                    new IntPtr(0x12340000),
                    expectedDllPath: Path.Combine(workerDir, "DifferentServer.dll"));
                mismatchResult.Status.Should().Be(ComModuleVerificationStatus.Mismatch);

                var verifyMismatchAct = () => ComModuleDiagnostics.VerifyRegisteredInprocServer(
                    ProbeNativeInterop.AlphaClsid,
                    new IntPtr(0x12340000),
                    expectedDllPath: Path.Combine(workerDir, "DifferentServer.dll"));
                verifyMismatchAct.Should().Throw<InvalidOperationException>();

                // Deactivate and unload immediately with ZERO GC.Collect() calls — proving zero managed RCWs were created!
                var deact = Executor.Deactivate(scope);
                deact.Success.Should().BeTrue();
                NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);
                File.Delete(dllPath);
            }
            finally
            {
                ProbeNativeInterop.CleanupIsolatedProbeWorker(workerDir);
            }
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ComManifestRewriterShouldRewriteWorkerDllPruneFilesStripPiiSaveIdempotentlyAndActivate(bool forceFallback)
    {
        TestWindow.OnStaThread(() =>
        {
            var rootDir = Path.Combine(Path.GetTempPath(), "AdaskoRewriter_" + Guid.NewGuid().ToString("N"));
            var templateDir = Path.Combine(rootDir, "template");
            var workerDir = Path.Combine(rootDir, "worker_00");
            Directory.CreateDirectory(templateDir);
            Directory.CreateDirectory(workerDir);

            try
            {
                var sourceManifestPath = Path.Combine(templateDir, "TemplateProbe.manifest");
                const string orphanTlbId = "{99999999-8888-7777-6666-555555555555}";
                const string unusedIid = "{11111111-2222-3333-4444-555555555555}";
                const string psOaIid = "{00020424-0000-0000-C000-000000000046}";
                File.WriteAllText(
                    sourceManifestPath,
                    $"""
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
                      <!-- Generated by ManifestGenerator on BuildHost-Secret-01 -->
                      <assemblyIdentity type="win32" name="TemplateProbe" version="1.0.0.0" />
                      <description>Licensee: TopSecretCustomerPII-2026</description>
                      <file name="NativeLifetimeProbe.dll">
                        <comClass clsid="{ProbeNativeInterop.AlphaClsidString}" threadingModel="Apartment" />
                        <comClass clsid="{ProbeNativeInterop.BetaClsidString}" threadingModel="Apartment" />
                      </file>
                      <file name="UnusedSatellite.dll">
                        <typelib tlbid="{orphanTlbId}" version="1.0" helpdir="" />
                      </file>
                      <comInterfaceExternalProxyStub iid="{unusedIid}" name="IUnusedSatellite" tlbid="{orphanTlbId}" proxyStubClsid32="{psOaIid}" />
                    </assembly>
                    """);

                var workerDllPath = Path.Combine(workerDir, "NativeLifetimeProbe_00.dll");
                File.Copy(ProbeNativeInterop.ProbeDllSourcePath, workerDllPath, overwrite: true);
                var workerManifestPath = Path.Combine(workerDir, "NativeLifetimeProbe_00.manifest");

                var rewriter = ComManifestRewriter
                    .FromFile(sourceManifestPath)
                    .WithAssemblyIdentityName("NativeLifetimeProbe.Worker00")
                    .RenameFiles(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["NativeLifetimeProbe.dll"] = "NativeLifetimeProbe_00.dll",
                    })
                    .RenameFile("NativeLifetimeProbe.dll", "NativeLifetimeProbe_00.dll") // Idempotent second rename!
                    .RemoveFile("UnusedSatellite.dll", removeOrphanExternalProxyStubs: true)
                    .RemoveFile("UnusedSatellite.dll") // Idempotent second remove!
                    .StripDescriptions();

                // Safety rejections:
                var inPlaceOverwriteAct = () => rewriter.SaveIfChanged(sourceManifestPath);
                inPlaceOverwriteAct.Should().Throw<InvalidOperationException>();

                var unknownFileAct = () => rewriter.RenameFile("TypoNonExistent.dll", "Worker00.dll");
                unknownFileAct.Should().Throw<InvalidOperationException>();

                var traversalFileAct = () => rewriter.RenameFile(@"..\Evil.dll", "Worker00.dll");
                traversalFileAct.Should().Throw<ArgumentException>();

                var traversalIdentityAct = () => rewriter.WithAssemblyIdentityName(@"..\EvilIdentity");
                traversalIdentityAct.Should().Throw<ArgumentException>();

                // SaveIfChanged writes UTF-8 without BOM on first write (true) and skips identical rewrite (false):
                rewriter.SaveIfChanged(workerManifestPath).Should().BeTrue();
                rewriter.SaveIfChanged(workerManifestPath).Should().BeFalse();

                var savedBytes = File.ReadAllBytes(workerManifestPath);
                savedBytes.Should().NotBeEmpty();
                savedBytes[0].Should().Be((byte)'<'); // No UTF-8 BOM!

                var savedXml = rewriter.ToXmlString();
                savedXml.Should().Contain("NativeLifetimeProbe.Worker00");
                savedXml.Should().Contain("NativeLifetimeProbe_00.dll");
                savedXml.Should().NotContain("TopSecretCustomerPII-2026");
                savedXml.Should().NotContain("BuildHost-Secret-01");
                savedXml.Should().NotContain("UnusedSatellite.dll");
                savedXml.Should().NotContain("99999999-8888-7777-6666-555555555555");

                // Removing the sole remaining <comClass> <file> must fail structural validation on ToUtf8Bytes!
                var brokenRewriter = ComManifestRewriter.FromXml(savedXml).RemoveFile("NativeLifetimeProbe_00.dll");
                var brokenBytesAct = () => brokenRewriter.ToUtf8Bytes();
                brokenBytesAct.Should().Throw<InvalidOperationException>();

                // Activate the rewritten worker manifest end-to-end under both Desktop WinSxS and ForceInProcessFallback!
                var act = Executor.Activate(
                    [new ComPathDescriptor(workerDllPath, workerManifestPath)],
                    new ComActivationOptions
                    {
                        ForceInProcessFallback = forceFallback,
                        PreloadModules = true,
                        UnloadModulesOnDeactivate = true,
                        PumpPendingMessages = false,
                    });
                act.Success.Should().BeTrue();
                var scope = act.Value!;

                ComModuleDiagnostics.VerifyRegisteredInprocServer(
                    ProbeNativeInterop.AlphaClsid,
                    scope.LoadedModuleHandles[0],
                    workerDllPath,
                    ProbeNativeInterop.AlphaIid,
                    methodSlotIndex: 3)
                    .Status.Should().Be(ComModuleVerificationStatus.Match);

                Executor.Deactivate(scope).Success.Should().BeTrue();
                NativeMethods.GetModuleHandleW(workerDllPath).Should().Be(IntPtr.Zero);
            }
            finally
            {
                if (Directory.Exists(rootDir))
                {
                    Directory.Delete(rootDir, recursive: true);
                }
            }
        });
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void ExerciseSuppressedUninitializeWithLiveRcw(IComActivationScope scope, IntPtr workerHandle)
    {
        var created = scope.CreateInstance<ProbeNativeInterop.IProbeAlpha>(ProbeNativeInterop.AlphaClsid);
        created.Success.Should().BeTrue();

        var coUninitProc = NativeMethods.GetProcAddress(workerHandle, "InvokeImportedCoUninitializeTwice");
        coUninitProc.Should().NotBe(IntPtr.Zero);
        var invokeCoUninitTwice = Marshal.GetDelegateForFunctionPointer<ExportedVoidProcDelegate>(coUninitProc);
        invokeCoUninitTwice();

        var oleUninitProc = NativeMethods.GetProcAddress(workerHandle, "InvokeImportedOleUninitializeTwice");
        oleUninitProc.Should().NotBe(IntPtr.Zero);
        var invokeOleUninitTwice = Marshal.GetDelegateForFunctionPointer<ExportedVoidProcDelegate>(oleUninitProc);
        invokeOleUninitTwice();

        // Apartment and live RCW remain completely healthy!
        NativeMethods.CoGetApartmentType(out var aptType, out _).Should().Be(0);
        (aptType is NativeMethods.APTTYPE_STA or NativeMethods.APTTYPE_MAINSTA).Should().BeTrue();

        created.Value!.ComputeAlpha(33, out var result).Should().Be(0);
        result.Should().Be(331);
    }
}
