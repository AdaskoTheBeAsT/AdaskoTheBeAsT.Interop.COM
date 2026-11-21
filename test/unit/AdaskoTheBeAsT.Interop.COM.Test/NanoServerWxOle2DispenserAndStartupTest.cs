using System.Globalization;
using System.Runtime.InteropServices;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable IDISP007 // Tests explicitly verify scope teardown.
#pragma warning disable MA0051 // Comprehensive W^X, OLE2, DispenserManager, and startup scenarios.
#pragma warning disable SCS0018, SEC0116 // Controlled repository source path inspection.
public class NanoServerWxOle2DispenserAndStartupTest
{
    [Fact]
    public void LibrarySourceShouldNeverReferencePageExecuteReadWriteAndNativeDispenserPagesMustEnforceWx()
    {
        var currentDir = new DirectoryInfo(AppContext.BaseDirectory);
        while (currentDir is not null && !File.Exists(Path.Combine(currentDir.FullName, "AdaskoTheBeAsT.Interop.COM.slnx")))
        {
            currentDir = currentDir.Parent;
        }

        currentDir.Should().NotBeNull();
        var repoSrcDir = Path.Combine(currentDir.FullName, "src", "AdaskoTheBeAsT.Interop.COM");
        Directory.Exists(repoSrcDir).Should().BeTrue();

        var sourceFiles = Directory.GetFiles(repoSrcDir, "*.cs", SearchOption.AllDirectories);
        sourceFiles.Should().NotBeEmpty();

        foreach (var file in sourceFiles)
        {
            var text = File.ReadAllText(file);
            text.Should().NotContain("PageExecuteReadWrite", because: $"{Path.GetFileName(file)} must strictly enforce W^X");
            text.Should().NotContain("PAGE_EXECUTE_READWRITE", because: $"{Path.GetFileName(file)} must strictly enforce W^X");
        }

        var factoryPtr = NanoServerDispenserManager.GetOrCreateNativeClassFactoryForTesting();
        factoryPtr.Should().NotBe(IntPtr.Zero);

        var codePage = NanoServerDispenserManager.CodePageAddress;
        var dataPage = NanoServerDispenserManager.DataPageAddress;
        codePage.Should().NotBe(IntPtr.Zero);
        dataPage.Should().NotBe(IntPtr.Zero);
        codePage.Should().NotBe(dataPage);

        var infoSize = new UIntPtr((uint)Marshal.SizeOf<NativeMethods.MEMORY_BASIC_INFORMATION>());
        NativeMethods.VirtualQuery(codePage, out var codeMbi, infoSize).Should().NotBe(UIntPtr.Zero);
        codeMbi.Protect.Should().Be(NativeMethods.PAGE_EXECUTE_READ);
        codeMbi.AllocationProtect.Should().Be(NativeMethods.PAGE_READWRITE);

        NativeMethods.VirtualQuery(dataPage, out var dataMbi, infoSize).Should().NotBe(UIntPtr.Zero);
        dataMbi.Protect.Should().Be(NativeMethods.PAGE_READWRITE);
    }

    [Fact]
    public void NativeWxDispenserManagerStubShouldExecuteEndToEndAcrossComSvcsVtablesAndRecycleSlotsUnderConcurrency()
    {
        var factoryPtr = NanoServerDispenserManager.GetOrCreateNativeClassFactoryForTesting();
        factoryPtr.Should().NotBe(IntPtr.Zero);

        // Single-threaded full ComSvcs.h ABI verification in C++ (x86 stdcall / x64 ms_abi).
        var hr = ProbeNativeInterop.SafeNativeMethods.ExerciseNativeDispenserFactory(
            factoryPtr,
            out var created,
            out var destroyed,
            out var destroyedStrings);
        hr.Should().Be(0);
        created.Should().Be(2);
        destroyed.Should().Be(2);
        destroyedStrings.Should().Be(1);

        // Exceed 64-slot holder pool sequentially and concurrently across 8 worker threads to verify atomic slot recycling!
        var failures = new List<Exception>();
        var threads = new Thread[8];
        for (var t = 0; t < threads.Length; t++)
        {
            threads[t] = new Thread(() =>
            {
                try
                {
                    for (var iter = 0; iter < 16; iter++)
                    {
                        var iterHr = ProbeNativeInterop.SafeNativeMethods.ExerciseNativeDispenserFactory(
                            factoryPtr,
                            out var c,
                            out var d,
                            out var ds);
                        if (iterHr != 0 || c != 2 || d != 2 || ds != 1)
                        {
                            throw new InvalidOperationException(string.Format(
                                CultureInfo.InvariantCulture,
                                "Dispenser iteration failed: hr=0x{0:X8}, c={1}, d={2}, ds={3}",
                                iterHr,
                                c,
                                d,
                                ds));
                        }
                    }
                }
                catch (Exception ex)
                {
                    lock (failures)
                    {
                        failures.Add(ex);
                    }
                }
            });
            threads[t].SetApartmentState(ApartmentState.MTA);
            threads[t].Start();
        }

        foreach (var thread in threads)
        {
            thread.Join();
        }

        failures.Should().BeEmpty();
    }

    [Fact]
    public void NanoServerOle2ShimShouldSupportInitializeRunSaveLoadAndPeImportPatching()
    {
        // 1. CoInitialize / OleInitialize / OleUninitialize on MTA (RPC_E_CHANGED_MODE) and STA (balanced init/uninit) threads.
        Exception? mtaInitError = null;
        var mtaInitThread = new Thread(() =>
        {
            try
            {
                NanoServerOle2Shim.OleInitializeImpl(IntPtr.Zero).Should().Be(NativeMethods.RPC_E_CHANGED_MODE);
            }
            catch (Exception ex)
            {
                mtaInitError = ex;
            }
        });
        mtaInitThread.SetApartmentState(ApartmentState.MTA);
        mtaInitThread.Start();
        mtaInitThread.Join();
        mtaInitError.Should().BeNull();

        Exception? initThreadError = null;
        var initThread = new Thread(() =>
        {
            try
            {
                NanoServerOle2Shim.CoInitializeImpl(new IntPtr(1)).Should().Be(NativeMethods.E_INVALIDARG);
                NanoServerOle2Shim.OleInitializeImpl(new IntPtr(1)).Should().Be(NativeMethods.E_INVALIDARG);

                var oleHr1 = NanoServerOle2Shim.OleInitializeImpl(IntPtr.Zero);
                oleHr1.Should().Be(NativeMethods.S_OK);
                var oleHr2 = NanoServerOle2Shim.OleInitializeImpl(IntPtr.Zero);
                oleHr2.Should().Be(NativeMethods.S_FALSE);

                NanoServerOle2Shim.OleUninitializeImpl();
                NanoServerOle2Shim.OleUninitializeImpl();
            }
            catch (Exception ex)
            {
                initThreadError = ex;
            }
        });
        initThread.SetApartmentState(ApartmentState.STA);
        initThread.Start();
        initThread.Join();
        initThreadError.Should().BeNull();

        // 2. OleRunImpl, OleSaveToStreamImpl, OleLoadFromStreamImpl, and PE IAT patching on STA thread!
        TestWindow.OnStaThread(() =>
        {
            var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("ole2_shim");
            try
            {
                var api = new ProbeNativeInterop.SimulatedFallbackActivationContextApi();
                var act = Executor.Activate(
                    [new ComPathDescriptor(dllPath, manifestPath)],
                    new ComActivationOptions { UnloadModulesOnDeactivate = true, PumpPendingMessages = false },
                    api);
                act.Success.Should().BeTrue();
                var scope = act.Value!;

                Marshal.ThrowExceptionForHR(ProbeNativeInterop.SafeNativeMethods.CreateNativePersistRunnableProbe(777, out var probeUnk));
                Marshal.ThrowExceptionForHR(ProbeNativeInterop.SafeNativeMethods.CreateNativeMemoryStream(out var streamPtr));
                try
                {
                    // Verify PE IAT patching of NativeLifetimeProbe.dll's OLE32.dll!OleRun import slot and clean restoration!
                    var loadedProbeModule = NativeMethods.GetModuleHandleW("NativeLifetimeProbe.dll");
                    loadedProbeModule.Should().NotBe(IntPtr.Zero);
                    var patchedCount = NanoServerOle2Shim.PatchModuleForTesting(loadedProbeModule);
                    patchedCount.Should().BeGreaterThanOrEqualTo(1);
                    try
                    {
                        // Calling InvokeImportedOleRun in C++ jumps through the patched IAT into NanoServerOle2Shim.OleRunImpl,
                        // incrementing PersistRunnableProbe.payloadValue by +1000 (777 -> 1777)!
                        ProbeNativeInterop.SafeNativeMethods.InvokeImportedOleRun(probeUnk).Should().Be(0);
                    }
                    finally
                    {
                        NanoServerOle2Shim.RestoreModuleForTesting(loadedProbeModule);
                    }

                    // Query IPersistStream to save via OleSaveToStreamImpl.
                    var iidPersistStream = new Guid(0x00000109, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46);
                    Marshal.ThrowExceptionForHR(NativeMethods.QueryInterface(probeUnk, ref iidPersistStream, out var persistStreamPtr));
                    try
                    {
                        NanoServerOle2Shim.OleSaveToStreamImpl(persistStreamPtr, streamPtr).Should().Be(0);
                    }
                    finally
                    {
                        _ = Marshal.Release(persistStreamPtr);
                    }

                    Marshal.ThrowExceptionForHR(ProbeNativeInterop.SafeNativeMethods.RewindNativeMemoryStream(streamPtr));

                    // Load brand-new instance via OleLoadFromStreamImpl (CoCreateInstance + IPersistStream::Load)!
                    var iidAlpha = ProbeNativeInterop.AlphaIid;
                    NanoServerOle2Shim.OleLoadFromStreamImpl(streamPtr, ref iidAlpha, out var loadedAlphaPtr).Should().Be(0);
                    loadedAlphaPtr.Should().NotBe(IntPtr.Zero);
                    try
                    {
                        var loadedAlpha = (ProbeNativeInterop.IProbeAlpha)Marshal.GetObjectForIUnknown(loadedAlphaPtr);
                        try
                        {
                            loadedAlpha.ComputeAlpha(23, out var total).Should().Be(0);
                            total.Should().Be(1800); // 1777 saved after OleRun + 23
                        }
                        finally
                        {
                            _ = Marshal.FinalReleaseComObject(loadedAlpha);
                        }
                    }
                    finally
                    {
                        _ = Marshal.Release(loadedAlphaPtr);
                    }
                }
                finally
                {
                    _ = Marshal.Release(streamPtr);
                    _ = Marshal.Release(probeUnk);
                    Executor.Deactivate(scope).Success.Should().BeTrue();
                }
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
    public void PrepareProcessAndCapabilityOptionsShouldBeIdempotentBeforeConcurrentStaWorkersStart()
    {
        try
        {
            NanoServerOptions.EnableLegacyOle2Shim = true;
            NanoServerOptions.EnableDispenserManagerFallback = true;

            // Idempotent repeated startup calls before launching concurrent STA workers.
            NanoServerOptions.PrepareProcess();
            NanoServerOptions.PrepareProcess();

            // On full Windows Desktop, capability query finds native OS providers rather than activating fallback shims unnecessarily.
            if (File.Exists(Path.Combine(Environment.SystemDirectory, "comsvcs.dll")))
            {
                NanoServerDispenserManager.ActiveProvider.Should().Be(NanoServerDispenserManager.DispenserProviderKind.NativeOs);
                NanoServerOptions.IsDispenserManagerFallbackActive.Should().BeFalse();
            }

            NanoServerOle2Shim.ActiveProvider.Should().NotBe(NanoServerOle2Shim.Ole2ProviderKind.None);

            var errors = new List<Exception>();
            var workers = new Thread[6];
            for (var i = 0; i < workers.Length; i++)
            {
                workers[i] = new Thread(() =>
                {
                    try
                    {
                        var res = Executor.Execute(
                            TestPaths.AssemblyPath,
                            TestPaths.ManifestPath,
                            () =>
                            {
                                var concat = new NativeCOM.StringConcatenatorClass();
                                concat.ConcatStrings("Concurrent ", "STA").Should().Be("Concurrent STA");
                            });
                        if (!res.Success)
                        {
                            throw res.Exception ?? new InvalidOperationException("Concurrent STA execution failed.");
                        }
                    }
                    catch (Exception ex)
                    {
                        lock (errors)
                        {
                            errors.Add(ex);
                        }
                    }
                });
                workers[i].SetApartmentState(ApartmentState.STA);
                workers[i].Start();
            }

            foreach (var worker in workers)
            {
                worker.Join();
            }

            errors.Should().BeEmpty();
        }
        finally
        {
            NanoServerOptions.ResetToEnvironmentDefaults();
        }
    }
}
#pragma warning restore SCS0018, SEC0116
#pragma warning restore MA0051
#pragma warning restore IDISP007
#pragma warning restore SA1121
