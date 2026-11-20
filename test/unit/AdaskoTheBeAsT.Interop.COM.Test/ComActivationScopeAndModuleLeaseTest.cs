using System.Diagnostics.Tracing;
using System.Runtime.InteropServices;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable SA1202 // Helper placement in test class
#pragma warning disable SA1204 // Static helper placement in test class
#pragma warning disable IDISP007 // Tests intentionally test explicit Deactivate and Dispose flows.
#pragma warning disable IDISP016 // Tests inspect post-release scope properties.
#pragma warning disable IDISP017 // Tests deliberately exercise both Deactivate and Dispose.
#pragma warning disable MA0051 // Comprehensive integration scenarios.
#pragma warning disable SCS0018, SEC0116 // Controlled test temp directory paths.
public class ComActivationScopeAndModuleLeaseTest
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ActivateShouldSupportCrossOrderMultiCoclassCallsAndCleanDeactivation(bool useNanoFallback)
    {
        var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("crossorder");
        try
        {
            TestWindow.OnStaThread(() =>
            {
                var options = new ComActivationOptions
                {
                    PreloadModules = !useNanoFallback,
                    UnloadModulesOnDeactivate = true,
                    PumpPendingMessages = true,
                };
                var api = useNanoFallback
                    ? new ProbeNativeInterop.SimulatedFallbackActivationContextApi()
                    : ActivationContextApi.Instance;

                var activation = Executor.Activate(
                    [new ComPathDescriptor(dllPath, manifestPath)],
                    options,
                    api);
                activation.Success.Should().BeTrue();
                using var scope = activation.Value!;
                scope.IsReleased.Should().BeFalse();
                scope.IsUsingNanoServerFallback.Should().Be(useNanoFallback);
                scope.LoadedModuleHandles.Should().HaveCount(1);

                ExerciseCrossOrderCoclasses(scope);

                scope.PumpPendingMessages(32).Success.Should().BeTrue();

                var deactivateResult = Executor.Deactivate(scope);
                deactivateResult.Success.Should().BeTrue();
                scope.IsReleased.Should().BeTrue();
                scope.PumpPendingMessages().Success.Should().BeFalse();

                // Second Deactivate / Dispose is idempotent.
                Executor.Deactivate(scope).Success.Should().BeTrue();
                scope.Dispose();

                // Synchronously unmapped and unlocked on disk without GC or sleep.
                NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);
                File.Delete(dllPath);
                File.Exists(dllPath).Should().BeFalse();
            });
        }
        finally
        {
            ProbeNativeInterop.CleanupIsolatedProbeWorker(workerDir);
        }
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static void ExerciseCrossOrderCoclasses(IComActivationScope scope)
    {
        // Order 1: Alpha -> Beta
        var alpha1 = (ProbeNativeInterop.IProbeAlpha)ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.AlphaClsid);
        var beta1 = (ProbeNativeInterop.IProbeBeta)ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.BetaClsid);
        scope.TrackObject(alpha1);
        scope.TrackObject(beta1);

        beta1.GetBetaTag(out var tag1).Should().Be(0);
        tag1.Should().Be(0xBE7A);
        beta1.CombineWithAlpha(alpha1, 4, out var combined1).Should().Be(0);
        combined1.Should().Be(48); // (4 * 10 + 1) + 7

        // Order 2: Beta -> Alpha inside the same scope
        var beta2 = (ProbeNativeInterop.IProbeBeta)ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.BetaClsid);
        var alpha2 = (ProbeNativeInterop.IProbeAlpha)ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.AlphaClsid);
        scope.TrackObject(beta2);
        scope.TrackObject(alpha2);

        beta2.CombineWithAlpha(alpha2, 5, out var combined2).Should().Be(0);
        combined2.Should().Be(59); // (5 * 10 + 2) + 7
    }

    [Fact]
    public void SessionRecyclingShouldIsolateMutableDllStateUnloadModulesAndAllowImmediateDirectoryDeletion()
    {
        TestWindow.OnStaThread(() =>
        {
            var api = new ProbeNativeInterop.SimulatedFallbackActivationContextApi();
            var options = new ComActivationOptions
            {
                UnloadModulesOnDeactivate = true,
                PumpPendingMessages = false,
            };

            for (var cycle = 1; cycle <= 3; cycle++)
            {
                var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker($"recycle_{cycle}");
                var activation = Executor.Activate(
                    [new ComPathDescriptor(dllPath, manifestPath)],
                    options,
                    api);
                activation.Success.Should().BeTrue();

                using (var scope = activation.Value!)
                {
                    var alpha = (ProbeNativeInterop.IProbeAlpha)ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.AlphaClsid);
                    scope.TrackObject(alpha);

                    // Each freshly copied DLL must start its mutable static counter at 1.
                    alpha.ComputeAlpha(10, out var firstValue).Should().Be(0);
                    firstValue.Should().Be(101);
                }

                NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);
                Directory.Delete(workerDir, recursive: true);
                Directory.Exists(workerDir).Should().BeFalse();
            }
        });
    }

    [Fact]
    public void OverlappingScopesOnSharedDllShouldKeepModuleMappedUntilFinalScopeEnds()
    {
        TestWindow.OnStaThread(() =>
        {
            var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("shared_overlap");
            try
            {
                var api = new ProbeNativeInterop.SimulatedFallbackActivationContextApi();
                var options = new ComActivationOptions
                {
                    UnloadModulesOnDeactivate = true,
                    PumpPendingMessages = false,
                };

                var act1 = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], options, api);
                act1.Success.Should().BeTrue();
                var scope1 = act1.Value!;

                var act2 = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], options, api);
                act2.Success.Should().BeTrue();
                var scope2 = act2.Value!;

                // LIFO: deactivate inner scope2 first; shared module MUST stay loaded for scope1!
                Executor.Deactivate(scope2).Success.Should().BeTrue();
                NativeMethods.GetModuleHandleW(dllPath).Should().NotBe(IntPtr.Zero);

                // scope1 can still instantiate coclasses from the shared module.
                var alpha = (ProbeNativeInterop.IProbeAlpha)ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.AlphaClsid);
                scope1.TrackObject(alpha);
                alpha.ComputeAlpha(2, out var computed).Should().Be(0);
                computed.Should().Be(21);

                // Deactivating scope1 drops the last lease and unloads the module.
                Executor.Deactivate(scope1).Success.Should().BeTrue();
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
    public void UntrackedLiveComObjectShouldPreventPrematureFreeLibraryUntilFinalReleaseAndSweep()
    {
        TestWindow.OnStaThread(() =>
        {
            var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("deferred_unload");
            try
            {
                var api = new ProbeNativeInterop.SimulatedFallbackActivationContextApi();
                var options = new ComActivationOptions
                {
                    UnloadModulesOnDeactivate = true,
                    PumpPendingMessages = false,
                };

                var activation = Executor.Activate([new ComPathDescriptor(dllPath, manifestPath)], options, api);
                activation.Success.Should().BeTrue();
                var scope = activation.Value!;

                // Intentionally NOT tracked in scope so native activeComObjectCount > 0 during Deactivate.
                var untrackedAlpha = (ProbeNativeInterop.IProbeAlpha)ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.AlphaClsid);

                Executor.Deactivate(scope).Success.Should().BeTrue();

                // DllCanUnloadNow returned S_FALSE, so the DLL MUST still be mapped and safe to call!
                NativeMethods.GetModuleHandleW(dllPath).Should().NotBe(IntPtr.Zero);
                untrackedAlpha.ComputeAlpha(3, out var val).Should().Be(0);
                val.Should().Be(31);

                // Now release the final native reference and sweep deferred unloads.
                _ = Marshal.FinalReleaseComObject(untrackedAlpha);
                Executor.TryUnloadUnusedModules().Should().BeGreaterThanOrEqualTo(1);
                NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);

                Directory.Delete(workerDir, recursive: true);
                Directory.Exists(workerDir).Should().BeFalse();
            }
            finally
            {
                if (Directory.Exists(workerDir))
                {
                    _ = Executor.TryUnloadUnusedModules();
                    Directory.Delete(workerDir, recursive: true);
                }
            }
        });
    }

    [Fact]
    public void WrongThreadAndNonLifoDeactivationShouldFailWithoutLosingStateForOwnerRetry()
    {
        TestWindow.OnStaThread(() =>
        {
            using var listener = new CapturingEventListener(ComInteropEventSource.Log);
            var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("lifo_thread");
            try
            {
                var options = new ComActivationOptions
                {
                    PreloadModules = true,
                    UnloadModulesOnDeactivate = true,
                    PumpPendingMessages = false,
                };

                var outerRes = Executor.Activate(dllPath, manifestPath, options);
                outerRes.Success.Should().BeTrue();
                var outerScope = outerRes.Value!;

                var innerRes = Executor.Activate(dllPath, manifestPath, options);
                innerRes.Success.Should().BeTrue();
                var innerScope = innerRes.Value!;

                // 1. Non-LIFO attempt on owner thread (outer before inner) must fail and keep both scopes intact.
                var nonLifo = Executor.Deactivate(outerScope);
                nonLifo.Success.Should().BeFalse();
                nonLifo.Exception.Should().BeOfType<InvalidOperationException>();
                outerScope.IsReleased.Should().BeFalse();
                innerScope.IsReleased.Should().BeFalse();

                // 2. Wrong-thread disposal attempt on innerScope must fail (and Dispose() must not throw, logging ScopeReleaseFailed).
                Result? wrongThreadResult = null;
                var otherThread = new Thread(() =>
                {
                    wrongThreadResult = Executor.Deactivate(innerScope);
                    innerScope.Dispose();
                });
                otherThread.Start();
                otherThread.Join();

                wrongThreadResult!.Success.Should().BeFalse();
                wrongThreadResult.Exception.Should().BeOfType<InvalidOperationException>();
                innerScope.IsReleased.Should().BeFalse();
                listener.Events.Should().Contain(e => e.EventId == 4);

                // 3. Correct LIFO retry on owner thread succeeds cleanly and unloads the module!
                Executor.Deactivate(innerScope).Success.Should().BeTrue();
                Executor.Deactivate(outerScope).Success.Should().BeTrue();
                innerScope.IsReleased.Should().BeTrue();
                outerScope.IsReleased.Should().BeTrue();
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
    public void ExistingMtaThreadShouldPropagateRpcEChangedModeWithoutAlteringApartmentForStaManifest()
    {
        Exception? failure = null;
        var mtaThread = new Thread(() =>
        {
            var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("mta_conflict", threadingModel: "Apartment");
            try
            {
                var api = new ProbeNativeInterop.SimulatedFallbackActivationContextApi
                {
                    ForceNotInitializedOnFirstRegister = true,
                };

                var activation = Executor.Activate(
                    [new ComPathDescriptor(dllPath, manifestPath)],
                    new ComActivationOptions { UnloadModulesOnDeactivate = true, PumpPendingMessages = false },
                    api);

                activation.Success.Should().BeFalse();
                activation.Exception.Should().NotBeNull();
                Thread.CurrentThread.GetApartmentState().Should().Be(ApartmentState.MTA);
                api.StaInitCalls.Should().Be(1);
                api.MtaInitCalls.Should().Be(0);
                api.UninitCalls.Should().Be(0);
                NativeMethods.GetModuleHandleW(dllPath).Should().Be(IntPtr.Zero);
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            finally
            {
                if (Directory.Exists(workerDir))
                {
                    Directory.Delete(workerDir, recursive: true);
                }
            }
        });
        mtaThread.SetApartmentState(ApartmentState.MTA);
        mtaThread.Start();
        mtaThread.Join();
        failure.Should().BeNull();
    }

    [Fact]
    public void ComExecutorAndLegacyIComExecutorExtensionMethodsShouldSupportActivationScopes()
    {
        TestWindow.OnStaThread(() =>
        {
            var comExecutor = new ComExecutor();
            IComExecutor asBaseInterface = comExecutor;

            var scopeRes = asBaseInterface.Activate(TestPaths.AssemblyPath, TestPaths.ManifestPath, new ComActivationOptions { PumpPendingMessages = false });
            scopeRes.Success.Should().BeTrue();
            using (var scope = scopeRes.Value!)
            {
                var obj = new NativeCOM.StringConcatenatorClass();
                scope.TrackObject(obj);
                obj.ConcatStrings("Ext ", "Works").Should().Be("Ext Works");
                asBaseInterface.Deactivate(scope).Success.Should().BeTrue();
            }

            // Custom third-party IComExecutor without IComActivationExecutor remains usable via Executor fallback.
            IComExecutor customLegacyExecutor = new LegacyStubComExecutor();
            var legacyScopeRes = customLegacyExecutor.Activate(
                TestPaths.AssemblyPath,
                TestPaths.ManifestPath,
                new ComActivationOptions { PumpPendingMessages = false });
            legacyScopeRes.Success.Should().BeTrue();
            using var legacyScope = legacyScopeRes.Value!;
            customLegacyExecutor.Deactivate(legacyScope).Success.Should().BeTrue();
        });
    }

    private sealed class LegacyStubComExecutor : IComExecutor
    {
        public Result Execute(string comAssemblyPath, string manifestPath, Action action)
            => new() { Success = true };

        public Result Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action)
            => new() { Success = true };

        public ComObjectCreationResult<T> Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory)
            where T : class
            => new() { Success = false, Exception = new NotSupportedException() };

        public ComObjectCreationResult<T> Create<T>(ICollection<ComPathDescriptor> comPathDescriptors, Func<T> factory)
            where T : class
            => new() { Success = false, Exception = new NotSupportedException() };

        public Result Free<T>(ComObjectHandle<T> comObjectHandle)
            where T : class
            => new() { Success = true };
    }
}
#pragma warning restore SCS0018, SEC0116
#pragma warning restore MA0051
#pragma warning restore IDISP017
#pragma warning restore IDISP016
#pragma warning restore IDISP007
#pragma warning restore SA1121
