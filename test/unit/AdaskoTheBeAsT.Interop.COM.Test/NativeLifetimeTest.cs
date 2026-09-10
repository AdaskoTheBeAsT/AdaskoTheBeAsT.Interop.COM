using System.Runtime.InteropServices;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

#pragma warning disable SYSLIB1054 // Built-in COM interface marshalling is not supported by LibraryImport.
#pragma warning disable IDISP007 // Tests own the factory-created COM handles.
#pragma warning disable IDISP016 // These assertions intentionally inspect released handles.
public class NativeLifetimeTest
{
#pragma warning disable SYSLIB1096 // These tests require runtime COM wrappers to verify Marshal.FinalReleaseComObject behavior.
    [ComImport]
    [Guid("c64bcaf7-3ee8-421d-a8ee-46cb0c9238f4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ILifetimeProbe
    {
        [PreserveSig]
        int Ping();

        void SetReleaseCallback(IReleaseCallback callback);
    }
#pragma warning restore SYSLIB1096

    [ComVisible(true)]
    [Guid("a80d257f-4121-49f2-84ca-0b4b40487d94")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IReleaseCallback
    {
        void OnReleased(uint threadId);
    }

    private static string AssemblyPath => Path.Combine(AppContext.BaseDirectory, "NativeCOM.dll");

    private static string ManifestPath => Path.Combine(AppContext.BaseDirectory, "NativeCOM.manifest");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReleaseShouldDestroyNativeObjectExactlyOnceOnOwnerThread(bool dispose)
    {
        TestWindow.OnStaThread(() =>
        {
            var before = NativeMethods.GetDestroyedCount();
            var owner = global::AdaskoTheBeAsT.Interop.COM.NativeMethods.GetCurrentThreadId();
            using var handle = CreateHandle();
            handle.ComObject!.Ping().Should().Be(0);
            NativeMethods.GetDestroyedCount().Should().Be(before);

            if (dispose)
            {
                handle.Dispose();
            }
            else
            {
                Executor.Free(handle).Success.Should().BeTrue();
            }

            handle.IsReleased.Should().BeTrue();
            NativeMethods.GetDestroyedCount().Should().Be(before + 1);
            NativeMethods.GetLastReleaseThreadId().Should().Be(owner);
            Executor.Free(handle).Success.Should().BeTrue();
            NativeMethods.GetDestroyedCount().Should().Be(before + 1);
        });
    }

    [Fact]
    public void FinalReleaseShouldInvalidateEveryManagedAlias()
    {
        TestWindow.OnStaThread(() =>
        {
            var before = NativeMethods.GetDestroyedCount();
            using var handle = CreateHandle();
            var alias = handle.ComObject!;

            Executor.Free(handle).Success.Should().BeTrue();

            Action useAlias = () => _ = alias.Ping();
            useAlias.Should().Throw<InvalidComObjectException>();
            NativeMethods.GetDestroyedCount().Should().Be(before + 1);
        });
    }

    [Fact]
    public void WrongThreadReleaseShouldNotDestroyNativeObject()
    {
        TestWindow.OnStaThread(() =>
        {
            var before = NativeMethods.GetDestroyedCount();
            using var handle = CreateHandle();
            Result? release = null;
            var thread = new Thread(() => release = Executor.Free(handle));
            thread.Start();
            thread.Join();

            release!.Success.Should().BeFalse();
            NativeMethods.GetDestroyedCount().Should().Be(before);
            handle.ComObject!.Ping().Should().Be(0);
            Executor.Free(handle).Success.Should().BeTrue();
            NativeMethods.GetDestroyedCount().Should().Be(before + 1);
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeDestructionCallbackShouldRejectReentrantRelease(bool pumpPendingMessages)
    {
        TestWindow.OnStaThread(() =>
        {
            var before = NativeMethods.GetDestroyedCount();
            var owner = global::AdaskoTheBeAsT.Interop.COM.NativeMethods.GetCurrentThreadId();
            using var window = new TestWindow();
            using var handle = CreateHandle(pumpPendingMessages);
            var expectedContext = handle.ActivationContextHandles[0];
            Result? reentrantRelease = null;
            uint callbackThread = 0;
            var contextCaptured = false;
            var callbackContext = IntPtr.Zero;
            var callback = new ReleaseCallback(threadId =>
            {
                callbackThread = threadId;
                contextCaptured = global::AdaskoTheBeAsT.Interop.COM.NativeMethods.GetCurrentActCtx(out callbackContext);
                if (contextCaptured && callbackContext != IntPtr.Zero)
                {
                    global::AdaskoTheBeAsT.Interop.COM.NativeMethods.ReleaseActCtx(callbackContext);
                }

                reentrantRelease = Executor.Free(handle);
            });
            handle.ComObject!.SetReleaseCallback(callback);
            window.Post();

            Executor.Free(handle).Success.Should().BeTrue();

            callbackThread.Should().Be(owner);
            contextCaptured.Should().BeTrue();
            callbackContext.Should().Be(expectedContext);
            reentrantRelease.Should().NotBeNull();
            reentrantRelease.Success.Should().BeFalse();
            reentrantRelease.Exception.Should().BeOfType<InvalidOperationException>();
            handle.IsReleased.Should().BeTrue();
            NativeMethods.GetDestroyedCount().Should().Be(before + 1);
            window.DeliveryCount.Should().Be(pumpPendingMessages ? 1 : 0);
            global::AdaskoTheBeAsT.Interop.COM.NativeMethods.PumpPendingMessages();
            window.DeliveryCount.Should().Be(1);
            GC.KeepAlive(callback);
        });
    }

    [Fact]
    public void WindowCallbackDuringFreeShouldRejectReentrantRelease()
    {
        TestWindow.OnStaThread(() =>
        {
            var before = NativeMethods.GetDestroyedCount();
            using var window = new TestWindow();
            using var handle = CreateHandle();
            Result? reentrantRelease = null;
            window.OnDelivery = () => reentrantRelease = Executor.Free(handle);
            window.Post();

            Executor.Free(handle).Success.Should().BeTrue();

            window.DeliveryCount.Should().Be(1);
            reentrantRelease.Should().NotBeNull();
            reentrantRelease.Success.Should().BeFalse();
            reentrantRelease.Exception.Should().BeOfType<InvalidOperationException>();
            handle.IsReleased.Should().BeTrue();
            NativeMethods.GetDestroyedCount().Should().Be(before + 1);
        });
    }

    private static ComObjectHandle<ILifetimeProbe> CreateHandle(bool pumpPendingMessages = true)
    {
        var creation = Executor.Create(
            AssemblyPath,
            ManifestPath,
            () =>
            {
                Marshal.ThrowExceptionForHR(NativeMethods.CreateLifetimeProbe(out var probe));
                return probe;
            },
            pumpPendingMessages);
        creation.Exception.Should().BeNull();
        creation.Success.Should().BeTrue();
        return creation.Value!;
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class ReleaseCallback : IReleaseCallback
    {
        private readonly Action<uint> _callback;

        public ReleaseCallback(Action<uint> callback)
        {
            _callback = callback;
        }

        public void OnReleased(uint threadId) => _callback(threadId);
    }

    private static class NativeMethods
    {
        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int CreateLifetimeProbe([MarshalAs(UnmanagedType.Interface)] out ILifetimeProbe probe);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int GetDestroyedCount();

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern uint GetLastReleaseThreadId();
    }
}
#pragma warning restore IDISP016
#pragma warning restore IDISP007
#pragma warning restore SYSLIB1054
