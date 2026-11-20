using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable SA1202 // Helper placement in test class
#pragma warning disable SA1204 // Static helper placement in test class
#pragma warning disable SYSLIB1054 // P/Invoke in test assembly targets net472..net11.0.
#pragma warning disable IDISP007 // Tests explicitly verify scope teardown.
#pragma warning disable MA0051 // Comprehensive module-isolation scenarios.
#pragma warning disable SCS0018, SEC0116 // Controlled test temp directory paths.
public class ComModuleDiagnosticsTest
{
    [Fact]
    public void VerifyInstanceModuleShouldMatchIsolatedDllAndDetectActualMismatchAcrossWorkers()
    {
        var (worker1Dir, dll1Path, manifest1Path) = ProbeNativeInterop.CreateIsolatedProbeWorker("diag_w1");
        var (worker2Dir, dll2Path, manifest2Path) = ProbeNativeInterop.CreateIsolatedProbeWorker("diag_w2");
        try
        {
            NanoServerOptions.EnsureInitialized();
            using var worker2Ready = new ManualResetEventSlim(false);
            using var diagnosticsCompleted = new ManualResetEventSlim(false);
            var nativeBase2 = IntPtr.Zero;
            ExceptionDispatchInfo? worker1Failure = null;
            ExceptionDispatchInfo? worker2Failure = null;

            var worker2Sta = new Thread(() =>
            {
                try
                {
                    var options = new ComActivationOptions
                    {
                        PreloadModules = true,
                        UnloadModulesOnDeactivate = true,
                        PumpPendingMessages = true,
                    };

                    var act2 = Executor.Activate([new ComPathDescriptor(dll2Path, manifest2Path)], options);
                    if (!act2.Success)
                    {
                        throw new InvalidOperationException("Worker 2 activation failed.", act2.Exception);
                    }

                    using var scope2 = act2.Value!;
                    var (base2, status2) = ExerciseWorker2Instance(scope2, dll2Path);
                    _ = Interlocked.Exchange(ref nativeBase2, base2);
                    worker2Ready.Set();
                    _ = diagnosticsCompleted.Wait(TimeSpan.FromSeconds(15));

                    var deact2 = Executor.Deactivate(scope2);
                    deact2.Success.Should().BeTrue();
                    status2.Should().Be(ComModuleVerificationStatus.Match);
                    NativeMethods.GetModuleHandleW(dll2Path).Should().Be(IntPtr.Zero);
                    File.Delete(dll2Path);
                }
                catch (Exception ex)
                {
                    Volatile.Write(ref worker2Failure, ExceptionDispatchInfo.Capture(ex));
                    worker2Ready.Set();
                }
            })
            {
                IsBackground = true,
            };
            worker2Sta.SetApartmentState(ApartmentState.STA);

            var worker1Sta = new Thread(() =>
            {
                try
                {
                    _ = worker2Ready.Wait(TimeSpan.FromSeconds(15));
                    if (Volatile.Read(ref worker2Failure) is not null)
                    {
                        return;
                    }

                    var options = new ComActivationOptions
                    {
                        PreloadModules = true,
                        UnloadModulesOnDeactivate = true,
                        PumpPendingMessages = true,
                    };

                    var act1 = Executor.Activate([new ComPathDescriptor(dll1Path, manifest1Path)], options);
                    if (!act1.Success)
                    {
                        throw new InvalidOperationException("Worker 1 activation failed.", act1.Exception);
                    }

                    using var scope1 = act1.Value!;
                    var observedBase2 = Interlocked.CompareExchange(ref nativeBase2, IntPtr.Zero, IntPtr.Zero);
                    var snapshot = CaptureWorker1DiagnosticsAgainstWorker2(scope1, dll1Path, observedBase2, dll2Path);
                    diagnosticsCompleted.Set();

                    var deact1 = Executor.Deactivate(scope1);
                    deact1.Success.Should().BeTrue();

                    snapshot.HrBase1.Should().Be(0);
                    snapshot.NativeBase1.Should().NotBe(observedBase2);
                    snapshot.InstanceModuleHandle.Should().Be(snapshot.NativeBase1);
                    snapshot.InstanceModuleHandle.Should().Be(snapshot.LeasedModuleHandle);
                    snapshot.MatchResult.Status.Should().Be(ComModuleVerificationStatus.Match);
                    snapshot.MatchResult.ModuleHandle.Should().Be(snapshot.NativeBase1);
                    snapshot.MatchResult.MethodSlotIndex.Should().Be(3);
                    snapshot.MismatchInspection.Status.Should().Be(ComModuleVerificationStatus.Mismatch);
                    snapshot.MismatchInspection.ModuleHandle.Should().Be(snapshot.NativeBase1);
                    snapshot.MismatchException.Should().NotBeNull();
                    snapshot.MismatchException.Message.Should().Contain("isolation mismatch");

                    NativeMethods.GetModuleHandleW(dll1Path).Should().Be(IntPtr.Zero);
                    File.Delete(dll1Path);
                }
                catch (Exception ex)
                {
                    Volatile.Write(ref worker1Failure, ExceptionDispatchInfo.Capture(ex));
                }
                finally
                {
                    diagnosticsCompleted.Set();
                }
            })
            {
                IsBackground = true,
            };
            worker1Sta.SetApartmentState(ApartmentState.STA);

            worker2Sta.Start();
            worker1Sta.Start();

            worker1Sta.Join(TimeSpan.FromSeconds(30)).Should().BeTrue();
            worker2Sta.Join(TimeSpan.FromSeconds(30)).Should().BeTrue();

            Volatile.Read(ref worker2Failure)?.Throw();
            Volatile.Read(ref worker1Failure)?.Throw();
        }
        finally
        {
            ProbeNativeInterop.CleanupIsolatedProbeWorker(worker2Dir);
            ProbeNativeInterop.CleanupIsolatedProbeWorker(worker1Dir);
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static (IntPtr NativeBase, ComModuleVerificationStatus Status) ExerciseWorker2Instance(
        IComActivationScope scope2,
        string dll2Path)
    {
        var obj2 = ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.AlphaClsid);
        scope2.TrackObject(obj2);
        var alpha2 = (ProbeNativeInterop.IProbeAlpha)obj2;
        _ = alpha2.GetModuleBaseAddress(out var nativeBase2);

        var match2 = ComModuleDiagnostics.VerifyInstanceModule(obj2, nativeBase2, dll2Path);
        return (nativeBase2, match2.Status);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static Worker1DiagnosticSnapshot CaptureWorker1DiagnosticsAgainstWorker2(
        IComActivationScope scope1,
        string dll1Path,
        IntPtr nativeBase2,
        string dll2Path)
    {
        var obj1 = ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.AlphaClsid);
        scope1.TrackObject(obj1);
        var alpha1 = (ProbeNativeInterop.IProbeAlpha)obj1;
        var hrBase1 = alpha1.GetModuleBaseAddress(out var nativeBase1);

        // 1. Direct IUnknown and specific secondary native interface match expected isolated worker DLL.
        var hMod1 = ComModuleDiagnostics.GetInstanceModuleHandle(obj1);
        var matchResult = ComModuleDiagnostics.VerifyInstanceModule(
            obj1,
            nativeBase1,
            dll1Path,
            ProbeNativeInterop.SecondaryIid,
            methodSlotIndex: 3);

        // 2. Comparing worker 1's instance against worker 2's expected HMODULE / path reports concrete Mismatch!
        var mismatchInspection = ComModuleDiagnostics.InspectInstanceModule(
            obj1,
            nativeBase2,
            dll2Path,
            ProbeNativeInterop.AlphaIid,
            methodSlotIndex: 3);

        InvalidOperationException? mismatchException = null;
        try
        {
            _ = ComModuleDiagnostics.VerifyInstanceModule(obj1, nativeBase2, dll2Path);
        }
        catch (InvalidOperationException ex)
        {
            mismatchException = ex;
        }

        return new Worker1DiagnosticSnapshot
        {
            HrBase1 = hrBase1,
            NativeBase1 = nativeBase1,
            LeasedModuleHandle = scope1.LoadedModuleHandles[0],
            InstanceModuleHandle = hMod1,
            MatchResult = matchResult,
            MismatchInspection = mismatchInspection,
            MismatchException = mismatchException,
        };
    }

    [Fact]
    public void CrossApartmentProxyVtableShouldReportInconclusiveRatherThanFalsePositiveMatchOrMismatch()
    {
        var marshaledStream = IntPtr.Zero;
        Exception? staError = null;
        using var marshaledReady = new ManualResetEventSlim(false);
        using var finishOwnerThread = new ManualResetEventSlim(false);

        var ownerSta = new Thread(() =>
        {
            var (workerDir, dllPath, manifestPath) = ProbeNativeInterop.CreateIsolatedProbeWorker("diag_proxy");
            try
            {
                var api = new ProbeNativeInterop.SimulatedFallbackActivationContextApi();
                var act = Executor.Activate(
                    [new ComPathDescriptor(dllPath, manifestPath)],
                    new ComActivationOptions { UnloadModulesOnDeactivate = true, PumpPendingMessages = true },
                    api);
                act.Success.Should().BeTrue();
                using var scope = act.Value!;
                var obj = ProbeNativeInterop.CreateComInstance(ProbeNativeInterop.AlphaClsid);
                scope.TrackObject(obj);

                var unk = Marshal.GetIUnknownForObject(obj);
                try
                {
                    var iidUnk = new Guid("00000000-0000-0000-C000-000000000046");
                    Marshal.ThrowExceptionForHR(SafeNativeMethods.CoMarshalInterThreadInterfaceInStream(ref iidUnk, unk, out marshaledStream));
                }
                finally
                {
                    _ = Marshal.Release(unk);
                }

                marshaledReady.Set();
                while (!finishOwnerThread.Wait(10))
                {
                    NativeMethods.PumpPendingMessages();
                }

                Executor.Deactivate(scope).Success.Should().BeTrue();
            }
            catch (Exception ex)
            {
                staError = ex;
                marshaledReady.Set();
            }
            finally
            {
                if (Directory.Exists(workerDir))
                {
                    Directory.Delete(workerDir, recursive: true);
                }
            }
        });
        ownerSta.SetApartmentState(ApartmentState.STA);
        ownerSta.Start();

        try
        {
            marshaledReady.Wait(TestContext.Current.CancellationToken);
            staError.Should().BeNull();

            // Unmarshal on a separate MTA thread so COM wraps the object in a cross-apartment proxy (combase.dll/ole32.dll).
            Exception? clientError = null;
            var clientThread = new Thread(() =>
            {
                try
                {
                    var iidUnk = new Guid("00000000-0000-0000-C000-000000000046");
                    Marshal.ThrowExceptionForHR(SafeNativeMethods.CoGetInterfaceAndReleaseStream(marshaledStream, ref iidUnk, out var proxyUnk));
                    marshaledStream = IntPtr.Zero;
                    try
                    {
                        var proxyObj = Marshal.GetObjectForIUnknown(proxyUnk);
                        try
                        {
                            var fakeExpectedHandle = new IntPtr(0x12340000);
                            var inspection = ComModuleDiagnostics.InspectInstanceModule(proxyObj, fakeExpectedHandle);
                            inspection.Status.Should().Be(ComModuleVerificationStatus.Inconclusive);

                            // Default throwOnInconclusive: false does NOT throw.
                            ComModuleDiagnostics.VerifyInstanceModule(proxyObj, fakeExpectedHandle, throwOnInconclusive: false)
                                .Status.Should().Be(ComModuleVerificationStatus.Inconclusive);

                            // throwOnInconclusive: true throws InvalidOperationException.
                            Action strictVerify = () => ComModuleDiagnostics.VerifyInstanceModule(
                                proxyObj,
                                fakeExpectedHandle,
                                throwOnInconclusive: true);
                            strictVerify.Should().Throw<InvalidOperationException>();
                        }
                        finally
                        {
                            _ = Marshal.FinalReleaseComObject(proxyObj);
                        }
                    }
                    finally
                    {
                        _ = Marshal.Release(proxyUnk);
                    }
                }
                catch (Exception ex)
                {
                    clientError = ex;
                }
            });
            clientThread.SetApartmentState(ApartmentState.MTA);
            clientThread.Start();
            clientThread.Join();
            clientError.Should().BeNull();
        }
        finally
        {
            finishOwnerThread.Set();
            ownerSta.Join();
        }
    }

    [Fact]
    public void ArgumentValidationShouldRejectInvalidInputs()
    {
        Action nullObj = () => ComModuleDiagnostics.GetInstanceModuleHandle(null!);
        nullObj.Should().Throw<ArgumentNullException>();

        Action negativeSlot = () => ComModuleDiagnostics.GetInstanceModuleHandle(new object(), methodSlotIndex: -1);
        negativeSlot.Should().Throw<ArgumentOutOfRangeException>();

        Action missingExpected = () => ComModuleDiagnostics.InspectInstanceModule(new object(), IntPtr.Zero, expectedDllPath: null);
        missingExpected.Should().Throw<ArgumentException>();
    }

    private sealed class Worker1DiagnosticSnapshot
    {
        internal int HrBase1 { get; set; }

        internal IntPtr NativeBase1 { get; set; }

        internal IntPtr LeasedModuleHandle { get; set; }

        internal IntPtr InstanceModuleHandle { get; set; }

        internal ComModuleInspectionResult MatchResult { get; set; } = null!;

        internal ComModuleInspectionResult MismatchInspection { get; set; } = null!;

        internal InvalidOperationException? MismatchException { get; set; }
    }

    private static class SafeNativeMethods
    {
        [DllImport("ole32.dll", ExactSpelling = true)]
        internal static extern int CoMarshalInterThreadInterfaceInStream(ref Guid riid, IntPtr pUnk, out IntPtr ppStm);

        [DllImport("ole32.dll", ExactSpelling = true)]
        internal static extern int CoGetInterfaceAndReleaseStream(IntPtr pStm, ref Guid riid, out IntPtr ppv);
    }
}
#pragma warning restore SCS0018, SEC0116
#pragma warning restore MA0051
#pragma warning restore IDISP007
#pragma warning restore SYSLIB1054
#pragma warning restore SA1121
