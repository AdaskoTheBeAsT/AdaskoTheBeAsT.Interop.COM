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
#pragma warning disable MA0051 // Comprehensive end-to-end activation and class-factory scenarios.
#pragma warning disable S1215 // Explicit GC collection verifies deterministic RCW finalization before module unload.
#pragma warning disable S8969 // Multi-targeted net472..net11.0 nullability annotations differ on BCL APIs.
#pragma warning disable SCS0018, SEC0116 // Controlled test temp directory paths.
public class ForceInProcessFallbackAndClassFactoryTest
{
    [Fact]
    public void ForceInProcessFallbackShouldBypassCreateActCtxOnDesktopAndRespectEnableSxsFallbackGuard()
    {
        TestWindow.OnStaThread(() =>
        {
            var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("force_fb");
            try
            {
                var spyApi = new DesktopSpyActivationContextApi();
                var options = new ComActivationOptions
                {
                    ForceInProcessFallback = true,
                    UnloadModulesOnDeactivate = true,
                    PumpPendingMessages = false,
                };

                var activation = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], options, spyApi);
                activation.Success.Should().BeTrue();
                spyApi.CreateActCtxCallCount.Should().Be(0);

                var scope = activation.Value!;
                try
                {
                    scope.IsUsingNanoServerFallback.Should().BeTrue();
                    scope.LoadedModuleHandles.Should().ContainSingle();

                    ExerciseIsolatedAlpha(scope, dllPath);
                }
                finally
                {
                    var deact = Executor.Deactivate(scope);
                    deact.Success.Should().BeTrue();
                }

                NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);

                // When EnableSxsFallback is explicitly disabled, ForceInProcessFallback fails closed with InvalidOperationException.
                try
                {
                    NanoServerOptions.EnableSxsFallback = false;
                    var disabledResult = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], options, spyApi);
                    disabledResult.Success.Should().BeFalse();
                    disabledResult.Exception.Should().BeOfType<InvalidOperationException>();
                    spyApi.CreateActCtxCallCount.Should().Be(0);
                }
                finally
                {
                    NanoServerOptions.ResetToEnvironmentDefaults();
                }
            }
            finally
            {
                ProbeNativeInterop.CleanupIsolatedProbeWorker(workerDir);
            }
        });
    }

    [Fact]
    public void NestedStaFallbackScopesShouldReferenceCountClassObjectRegistrationAndRejectConflictingDllPath()
    {
        TestWindow.OnStaThread(() =>
        {
            var (worker1Dir, dll1Path, manifest1Path) = ProbeNativeInterop.CreateIsolatedProbeWorker("nested_sta1");
            var (worker2Dir, dll2Path, manifest2Path) = ProbeNativeInterop.CreateIsolatedProbeWorker("nested_sta2");
            try
            {
                var spyApi = new DesktopSpyActivationContextApi();
                var options = new ComActivationOptions
                {
                    ForceInProcessFallback = true,
                    UnloadModulesOnDeactivate = true,
                    PumpPendingMessages = false,
                };

                var outerAct = Executor.Activate([new ComPathDescriptor(dll1Path, manifest1Path)], options, spyApi);
                outerAct.Success.Should().BeTrue();
                var outerScope = outerAct.Value!;
                var initialRegisterCount = spyApi.RegisterCallCount;
                initialRegisterCount.Should().Be(3); // Alpha, Beta, PersistProbe coclasses

                // 1. Nested scope on same STA thread for same (CLSID, DllPath) reuses existing CoRegisterClassObject tokens.
                var innerSameAct = Executor.Activate([new ComPathDescriptor(dll1Path, manifest1Path)], options, spyApi);
                innerSameAct.Success.Should().BeTrue();
                var innerSameScope = innerSameAct.Value!;
                spyApi.RegisterCallCount.Should().Be(initialRegisterCount);
                spyApi.RevokeCallCount.Should().Be(0);

                // 2. Conflicting DLL path for the same CLSID in the same STA apartment fails closed without clobbering outer/inner registrations.
                var conflictAct = Executor.Activate([new ComPathDescriptor(dll2Path, manifest2Path)], options, spyApi);
                conflictAct.Success.Should().BeFalse();
                NativeMethods.GetModuleHandleW(dll2Path).Should().Be(IntPtr.Zero);

                // 3. Deactivating inner scope decrements lease count without revoking outer scope's active registration.
                Executor.Deactivate(innerSameScope).Success.Should().BeTrue();
                spyApi.RevokeCallCount.Should().Be(0);

                // Outer scope remains fully usable via CoCreateInstance!
                ComModuleDiagnostics.VerifyRegisteredInprocServer(
                    ProbeNativeInterop.AlphaClsid,
                    outerScope.LoadedModuleHandles[0],
                    dll1Path)
                    .Status.Should().Be(ComModuleVerificationStatus.Match);

                // 4. Deactivating outer scope revokes all 3 registrations and unloads dll1Path.
                Executor.Deactivate(outerScope).Success.Should().BeTrue();
                spyApi.RevokeCallCount.Should().Be(initialRegisterCount);
                NativeMethods.GetModuleHandleW(dll1Path).Should().Be(IntPtr.Zero);

                var iidUnknown = new Guid("00000000-0000-0000-C000-000000000046");
                var alphaClsid = ProbeNativeInterop.AlphaClsid;
                var postDeactivateHr = NativeMethods.CoCreateInstance(
                    ref alphaClsid,
                    IntPtr.Zero,
                    NativeMethods.CLSCTX_INPROC_SERVER,
                    ref iidUnknown,
                    out var leakedPtr);
                if (leakedPtr != IntPtr.Zero)
                {
                    _ = Marshal.Release(leakedPtr);
                }

                postDeactivateHr.Should().Be(NativeMethods.REGDB_E_CLASSNOTREG);
            }
            finally
            {
                ProbeNativeInterop.CleanupIsolatedProbeWorker(worker2Dir);
                ProbeNativeInterop.CleanupIsolatedProbeWorker(worker1Dir);
            }
        });
    }

    [Fact]
    public void TwoParallelStaWorkersWithForceInProcessFallbackShouldIsolateSameClsidToDistinctModulesSimultaneously()
    {
        var (workerADir, dllAPath, manifestAPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("par_fb_a");
        var (workerBDir, dllBPath, manifestBPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("par_fb_b");
        try
        {
            using var barrier = new Barrier(2);
            var baseA = IntPtr.Zero;
            var baseB = IntPtr.Zero;
            ExceptionDispatchInfo? failureA = null;
            ExceptionDispatchInfo? failureB = null;

            var threadA = new Thread(() => RunParallelStaFallbackWorker(dllAPath, manifestAPath, barrier, out baseA, out failureA))
            {
                IsBackground = true,
            };
            threadA.SetApartmentState(ApartmentState.STA);

            var threadB = new Thread(() => RunParallelStaFallbackWorker(dllBPath, manifestBPath, barrier, out baseB, out failureB))
            {
                IsBackground = true,
            };
            threadB.SetApartmentState(ApartmentState.STA);

            threadA.Start();
            threadB.Start();

            threadA.Join(TimeSpan.FromSeconds(30)).Should().BeTrue();
            threadB.Join(TimeSpan.FromSeconds(30)).Should().BeTrue();

            failureA?.Throw();
            failureB?.Throw();

            baseA.Should().NotBe(IntPtr.Zero);
            baseB.Should().NotBe(IntPtr.Zero);
            baseA.Should().NotBe(baseB);
        }
        finally
        {
            ProbeNativeInterop.CleanupIsolatedProbeWorker(workerBDir);
            ProbeNativeInterop.CleanupIsolatedProbeWorker(workerADir);
        }
    }

    [Fact]
    public void ConcurrentMtaFallbackScopesForSameDllShouldRefCountUntilLastScopeDeactivates()
    {
        var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("mta_refcnt", threadingModel: "Both");
        try
        {
            using var bothActiveBarrier = new Barrier(2);
            using var firstDeactivatedEvent = new ManualResetEventSlim(false);
            using var secondVerifiedEvent = new ManualResetEventSlim(false);
            var spyApi = new DesktopSpyActivationContextApi();
            ExceptionDispatchInfo? mta1Failure = null;
            ExceptionDispatchInfo? mta2Failure = null;

            var options = new ComActivationOptions
            {
                ForceInProcessFallback = true,
                UnloadModulesOnDeactivate = true,
                PumpPendingMessages = false,
            };

            var mta1 = new Thread(() =>
            {
                try
                {
                    var act = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], options, spyApi);
                    act.Success.Should().BeTrue();
                    var scope = act.Value!;
                    bothActiveBarrier.SignalAndWait(TimeSpan.FromSeconds(15)).Should().BeTrue();

                    Executor.Deactivate(scope).Success.Should().BeTrue();
                    firstDeactivatedEvent.Set();
                    secondVerifiedEvent.Wait(TimeSpan.FromSeconds(15)).Should().BeTrue();
                }
                catch (Exception ex)
                {
                    mta1Failure = ExceptionDispatchInfo.Capture(ex);
                    firstDeactivatedEvent.Set();
                }
            })
            {
                IsBackground = true,
            };
            mta1.SetApartmentState(ApartmentState.MTA);

            var mta2 = new Thread(() =>
            {
                try
                {
                    var act = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], options, spyApi);
                    act.Success.Should().BeTrue();
                    var scope = act.Value!;
                    bothActiveBarrier.SignalAndWait(TimeSpan.FromSeconds(15)).Should().BeTrue();

                    firstDeactivatedEvent.Wait(TimeSpan.FromSeconds(15)).Should().BeTrue();
                    spyApi.RevokeCallCount.Should().Be(0);

                    // Even after MTA thread 1 deactivated, MTA thread 2's shared registration lease remains active!
                    ComModuleDiagnostics.VerifyRegisteredInprocServer(
                        ProbeNativeInterop.AlphaClsid,
                        scope.LoadedModuleHandles[0],
                        dllPath)
                        .Status.Should().Be(ComModuleVerificationStatus.Match);

                    secondVerifiedEvent.Set();
                    Executor.Deactivate(scope).Success.Should().BeTrue();
                }
                catch (Exception ex)
                {
                    mta2Failure = ExceptionDispatchInfo.Capture(ex);
                    secondVerifiedEvent.Set();
                }
            })
            {
                IsBackground = true,
            };
            mta2.SetApartmentState(ApartmentState.MTA);

            mta1.Start();
            mta2.Start();

            mta1.Join(TimeSpan.FromSeconds(30)).Should().BeTrue();
            mta2.Join(TimeSpan.FromSeconds(30)).Should().BeTrue();

            mta1Failure?.Throw();
            mta2Failure?.Throw();

            spyApi.RegisterCallCount.Should().Be(3);
            spyApi.RevokeCallCount.Should().Be(3);
            NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);
        }
        finally
        {
            ProbeNativeInterop.CleanupIsolatedProbeWorker(workerDir);
        }
    }

    [Fact]
    public void ComClassFactoryAndScopeExtensionsShouldInstantiateDirectlyWithoutSxsOrRegistryAndEnforceLifetimes()
    {
        TestWindow.OnStaThread(() =>
        {
            var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("direct_cf");
            try
            {
                // 1. Direct CreateInstanceFromDll<TInterface> with zero SxS manifest or registry entries.
                var unavailableResult = ComClassFactory.CreateInstanceFromDll<ProbeNativeInterop.IProbeAlpha>(
                    dllPath,
                    ProbeNativeInterop.UnavailableClsid,
                    out var failedHandle);
                unavailableResult.Success.Should().BeFalse();
                failedHandle.Should().Be(IntPtr.Zero);
                NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);

                var directFromDll = ComClassFactory.CreateInstanceFromDll<ProbeNativeInterop.IProbeAlpha>(
                    dllPath,
                    ProbeNativeInterop.AlphaClsid,
                    out var loadedHandle);
                directFromDll.Success.Should().BeTrue();
                loadedHandle.Should().NotBe(IntPtr.Zero);
                try
                {
                    ExerciseDirectModuleHandleCreation(loadedHandle, dllPath, directFromDll.Value!);
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    _ = NativeMethods.FreeLibrary(loadedHandle);
                }

                // 2. Scope without PreloadModules fails closed when calling scope.CreateInstance.
                var nonPreloadedAct = Executor.Activate(
                    [new ComPathDescriptor(dllPath, manifestPath)],
                    new ComActivationOptions { PreloadModules = false, PumpPendingMessages = false });
                nonPreloadedAct.Success.Should().BeTrue();
                var nonPreloadedScope = nonPreloadedAct.Value!;
                var noHandlesCreate = nonPreloadedScope.CreateInstance<ProbeNativeInterop.IProbeAlpha>(ProbeNativeInterop.AlphaClsid);
                noHandlesCreate.Success.Should().BeFalse();
                noHandlesCreate.Exception.Should().BeOfType<InvalidOperationException>();
                Executor.Deactivate(nonPreloadedScope).Success.Should().BeTrue();

                // 3. Scope with PreloadModules = true + UnloadModulesOnDeactivate = true tracks RCWs and enforces thread affinity & bounds.
                var preloadedAct = Executor.Activate(
                    [new ComPathDescriptor(dllPath, manifestPath)],
                    new ComActivationOptions
                    {
                        PreloadModules = true,
                        UnloadModulesOnDeactivate = true,
                        PumpPendingMessages = false,
                    });
                preloadedAct.Success.Should().BeTrue();
                var scope = preloadedAct.Value!;

                scope.CreateInstance(ProbeNativeInterop.AlphaClsid, moduleIndex: -1).Exception.Should().BeOfType<ArgumentOutOfRangeException>();
                scope.CreateInstance(ProbeNativeInterop.AlphaClsid, moduleIndex: 5).Exception.Should().BeOfType<ArgumentOutOfRangeException>();

                ComCreateResult<ProbeNativeInterop.IProbeAlpha>? wrongThreadResult = null;
                var helperThread = new Thread(() => wrongThreadResult = scope.CreateInstance<ProbeNativeInterop.IProbeAlpha>(ProbeNativeInterop.AlphaClsid))
                {
                    IsBackground = true,
                };
                helperThread.Start();
                helperThread.Join(TimeSpan.FromSeconds(10)).Should().BeTrue();
                wrongThreadResult.Should().NotBeNull();
                wrongThreadResult!.Success.Should().BeFalse();
                wrongThreadResult.Exception.Should().BeOfType<InvalidOperationException>();

                ExerciseScopeDirectCreation(scope);

                Executor.Deactivate(scope).Success.Should().BeTrue();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                _ = Executor.TryUnloadUnusedModules();
                NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);
            }
            finally
            {
                ProbeNativeInterop.CleanupIsolatedProbeWorker(workerDir);
            }
        });
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void ExerciseIsolatedAlpha(IComActivationScope scope, string expectedDllPath)
    {
        var raw = ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.AlphaClsid);
        try
        {
            var alpha = (ProbeNativeInterop.IProbeAlpha)raw;
            alpha.ComputeAlpha(5, out var computed).Should().Be(0);
            computed.Should().Be(51);
            ComModuleDiagnostics.VerifyInstanceModule(
                raw,
                scope.LoadedModuleHandles[0],
                expectedDllPath,
                ProbeNativeInterop.AlphaIid,
                methodSlotIndex: 3)
                .Status.Should().Be(ComModuleVerificationStatus.Match);
        }
        finally
        {
            if (Marshal.IsComObject(raw))
            {
                _ = Marshal.FinalReleaseComObject(raw);
            }
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void ExerciseDirectModuleHandleCreation(
        IntPtr loadedHandle,
        string expectedDllPath,
        ProbeNativeInterop.IProbeAlpha firstAlpha)
    {
        try
        {
            firstAlpha.ComputeAlpha(12, out var val).Should().Be(0);
            val.Should().Be(121);

            ComModuleDiagnostics.VerifyInstanceModule(
                firstAlpha,
                loadedHandle,
                expectedDllPath,
                ProbeNativeInterop.AlphaIid,
                methodSlotIndex: 3)
                .Status.Should().Be(ComModuleVerificationStatus.Match);

            var secondaryCreation = ComClassFactory.CreateInstance(
                loadedHandle,
                ProbeNativeInterop.AlphaClsid,
                ProbeNativeInterop.SecondaryIid);
            secondaryCreation.Success.Should().BeTrue();
            try
            {
                var secondary = (ProbeNativeInterop.IProbeSecondary)secondaryCreation.Value!;
                secondary.GetSignature(out var sig).Should().Be(0);
                sig.Should().Be(0xA1FA);
            }
            finally
            {
                if (secondaryCreation.Value is not null && Marshal.IsComObject(secondaryCreation.Value))
                {
                    _ = Marshal.FinalReleaseComObject(secondaryCreation.Value);
                }
            }
        }
        finally
        {
            if (Marshal.IsComObject(firstAlpha))
            {
                _ = Marshal.FinalReleaseComObject(firstAlpha);
            }
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void ExerciseScopeDirectCreation(IComActivationScope scope)
    {
        var created = scope.CreateInstance<ProbeNativeInterop.IProbeAlpha>(ProbeNativeInterop.AlphaClsid);
        created.Success.Should().BeTrue();
        created.Value!.ComputeAlpha(25, out var output).Should().Be(0);
        output.Should().Be(251);
        created.Value.GetModuleBaseAddress(out var moduleBase).Should().Be(0);
        moduleBase.Should().Be(scope.LoadedModuleHandles[0]);
    }

    private static void RunParallelStaFallbackWorker(
        string dllPath,
        string manifestPath,
        Barrier barrier,
        out IntPtr observedModuleBase,
        out ExceptionDispatchInfo? failure)
    {
        observedModuleBase = IntPtr.Zero;
        failure = null;
        try
        {
            var options = new ComActivationOptions
            {
                ForceInProcessFallback = true,
                UnloadModulesOnDeactivate = true,
                PumpPendingMessages = true,
            };

            var act = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], options);
            if (!act.Success)
            {
                throw new InvalidOperationException("Parallel ForceInProcessFallback activation failed.", act.Exception);
            }

            var scope = act.Value!;
            barrier.SignalAndWait(TimeSpan.FromSeconds(15)).Should().BeTrue();
            observedModuleBase = CaptureScopeModuleBaseAndVerify(scope, dllPath);
            barrier.SignalAndWait(TimeSpan.FromSeconds(15)).Should().BeTrue();

            var deact = Executor.Deactivate(scope);
            deact.Success.Should().BeTrue();
            NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);
        }
        catch (Exception ex)
        {
            failure = ExceptionDispatchInfo.Capture(ex);
            try
            {
                barrier.RemoveParticipant();
            }
            catch (InvalidOperationException probeEx)
            {
                _ = probeEx;
            }
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static IntPtr CaptureScopeModuleBaseAndVerify(IComActivationScope scope, string dllPath)
    {
        var raw = ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.AlphaClsid);
        try
        {
            var alpha = (ProbeNativeInterop.IProbeAlpha)raw;
            alpha.GetModuleBaseAddress(out var modBase).Should().Be(0);
            modBase.Should().Be(scope.LoadedModuleHandles[0]);
            ComModuleDiagnostics.VerifyInstanceModule(
                raw,
                scope.LoadedModuleHandles[0],
                dllPath,
                ProbeNativeInterop.AlphaIid,
                methodSlotIndex: 3)
                .Status.Should().Be(ComModuleVerificationStatus.Match);
            return modBase;
        }
        finally
        {
            if (Marshal.IsComObject(raw))
            {
                _ = Marshal.FinalReleaseComObject(raw);
            }
        }
    }

    private sealed class DesktopSpyActivationContextApi : ActivationContextApi
    {
        private int _createActCtxCalls;
        private int _registerCalls;
        private int _revokeCalls;

        internal int CreateActCtxCallCount => Volatile.Read(ref _createActCtxCalls);

        internal int RegisterCallCount => Volatile.Read(ref _registerCalls);

        internal int RevokeCallCount => Volatile.Read(ref _revokeCalls);

        internal override IntPtr CreateActCtx(ref ActCtx context)
        {
            _ = Interlocked.Increment(ref _createActCtxCalls);
            return base.CreateActCtx(ref context);
        }

        internal override int RegisterClassObject(ref Guid clsid, IntPtr classFactory, out uint registrationToken)
        {
            _ = Interlocked.Increment(ref _registerCalls);
            return base.RegisterClassObject(ref clsid, classFactory, out registrationToken);
        }

        internal override int RevokeClassObject(uint registrationToken)
        {
            _ = Interlocked.Increment(ref _revokeCalls);
            return base.RevokeClassObject(registrationToken);
        }
    }
}
