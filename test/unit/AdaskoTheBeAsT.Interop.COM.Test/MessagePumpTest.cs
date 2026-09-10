using System.Runtime.InteropServices;
using System.Security;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

public partial class MessagePumpTest
{
    private const uint WmQuit = 0x0012;
    private const uint PmNoRemove = 0x0000;

    [Theory]
    [InlineData(0)]
    [InlineData(42)]
    [InlineData(-1)]
    public void PumpPendingMessagesShouldPreserveQuitAndExitCode(int exitCode)
    {
        global::AdaskoTheBeAsT.Interop.COM.NativeMethods.PostQuitMessage(exitCode);
        try
        {
            global::AdaskoTheBeAsT.Interop.COM.NativeMethods.PumpPendingMessages();

            var pending = global::AdaskoTheBeAsT.Interop.COM.NativeMethods.PeekMessage(
                out var message,
                IntPtr.Zero,
                0,
                0,
                PmNoRemove);
            pending.Should().BeTrue();
            message.message.Should().Be(WmQuit);
            unchecked((int)message.wParam.ToUInt64()).Should().Be(exitCode);
        }
        finally
        {
            _ = global::AdaskoTheBeAsT.Interop.COM.NativeMethods.PeekMessage(
                out _,
                IntPtr.Zero,
                WmQuit,
                WmQuit,
                global::AdaskoTheBeAsT.Interop.COM.NativeMethods.PM_REMOVE);
        }
    }

    [Fact]
    public void PumpPendingMessagesShouldLeaveMessagesBeyondBudgetQueued()
    {
        TestWindow.OnStaThread(() =>
        {
            using var window = new TestWindow();
            const int budget = global::AdaskoTheBeAsT.Interop.COM.NativeMethods.MaxMessagesPerPump;
            for (int i = 0; i <= budget; i++)
            {
                window.Post();
            }

            global::AdaskoTheBeAsT.Interop.COM.NativeMethods.PumpPendingMessages();
            window.DeliveryCount.Should().Be(budget);
            NativeMethods.PeekMessage(out _, IntPtr.Zero, 0, 0, PmNoRemove).Should().BeTrue();

            global::AdaskoTheBeAsT.Interop.COM.NativeMethods.PumpPendingMessages();
            window.DeliveryCount.Should().Be(budget + 1);
            NativeMethods.PeekMessage(out _, IntPtr.Zero, 0, 0, PmNoRemove).Should().BeFalse();
        });
    }

    [Fact]
    public void PumpPendingMessagesShouldTranslateAndDispatchQueuedMessage()
    {
        TestWindow.OnStaThread(() =>
        {
            using var window = new TestWindow();
            window.Post(wParam: 42, lParam: -7);
            var (comAssemblyPath, manifestPath) = GetPaths();

            var result = Executor.Execute(comAssemblyPath, manifestPath, () => { });

            result.Success.Should().BeTrue();
            window.DeliveryCount.Should().Be(1);
            window.LastWParam.Should().Be(new UIntPtr(42));
            window.LastLParam.Should().Be(new IntPtr(-7));
        });
    }

    [Theory]
    [InlineData(0x4E2D)]
    [InlineData(0x03A9)]
    [InlineData(0x0416)]
    public void PumpPendingMessagesShouldPreserveUnicodeCharacters(int character)
    {
        TestWindow.OnStaThread(() =>
        {
            using var window = new TestWindow();
            window.Post(TestWindow.CharacterMessage, (uint)character);

            global::AdaskoTheBeAsT.Interop.COM.NativeMethods.PumpPendingMessages();

            window.DeliveryCount.Should().Be(1);
            window.LastWParam.Should().Be(new UIntPtr((uint)character));
        });
    }

    [Fact]
    public void DisabledPumpingShouldPreserveCustomHostThreadMessages()
    {
        TestWindow.OnStaThread(() =>
        {
            const uint hostMessage = 0x8123;
            _ = NativeMethods.PeekMessage(out _, IntPtr.Zero, 0, 0, PmNoRemove);
            NativeMethods.PostThreadMessage(
                NativeMethods.GetCurrentThreadId(), hostMessage, new IntPtr(42), new IntPtr(-7)).Should().BeTrue();
            var (comAssemblyPath, manifestPath) = GetPaths();
            var executor = new ComExecutor(pumpPendingMessages: false);

            executor.Execute(comAssemblyPath, manifestPath, () => { }).Success.Should().BeTrue();
            var creation = executor.Create(comAssemblyPath, manifestPath, () => new object());
            creation.Success.Should().BeTrue();
            creation.Value!.Dispose();

            global::AdaskoTheBeAsT.Interop.COM.NativeMethods.PeekMessage(
                out var message,
                IntPtr.Zero,
                hostMessage,
                hostMessage,
                global::AdaskoTheBeAsT.Interop.COM.NativeMethods.PM_REMOVE).Should().BeTrue();
            message.wParam.Should().Be(new UIntPtr(42));
            message.lParam.Should().Be(new IntPtr(-7));
        });
    }

    private static (string ComAssemblyPath, string ManifestPath) GetPaths()
    {
        var currentPath = AppContext.BaseDirectory;
        var comAssemblyPath = Path.GetFullPath(Path.Combine(currentPath, "NativeCOM.dll"));
        var manifestPath = Path.GetFullPath(Path.Combine(currentPath, "NativeCOM.manifest"));

        File.Exists(comAssemblyPath).Should().BeTrue();
        File.Exists(manifestPath).Should().BeTrue();

        return (comAssemblyPath, manifestPath);
    }

#if NET7_0_OR_GREATER
    [SuppressUnmanagedCodeSecurity]
    internal static partial class NativeMethods
    {
        [LibraryImport("user32.dll", EntryPoint = "PeekMessageW", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static partial bool PeekMessage(
            out Msg lpMsg,
            IntPtr hWnd,
            uint wMsgFilterMin,
            uint wMsgFilterMax,
            uint wRemoveMsg);

        [LibraryImport("user32.dll", EntryPoint = "PostThreadMessageW", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static partial bool PostThreadMessage(
            uint idThread,
            uint msg,
            IntPtr wParam,
            IntPtr lParam);

        [LibraryImport("kernel32.dll")]
        internal static partial uint GetCurrentThreadId();

        [StructLayout(LayoutKind.Sequential)]
        internal struct Msg
        {
            public readonly IntPtr Hwnd;
            public readonly uint Message;
            public readonly UIntPtr WParam;
            public readonly IntPtr LParam;
            public readonly uint Time;
            public readonly int X;
            public readonly int Y;
            public readonly uint LPrivate;
        }
    }
#else
    [SuppressUnmanagedCodeSecurity]
    internal static class NativeMethods
    {
        [DllImport("user32.dll", EntryPoint = "PeekMessageW", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool PeekMessage(
            out Msg lpMsg,
            IntPtr hWnd,
            uint wMsgFilterMin,
            uint wMsgFilterMax,
            uint wRemoveMsg);

        [DllImport("user32.dll", EntryPoint = "PostThreadMessageW", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool PostThreadMessage(
            uint idThread,
            uint msg,
            IntPtr wParam,
            IntPtr lParam);

        [DllImport("kernel32.dll")]
        internal static extern uint GetCurrentThreadId();

        [StructLayout(LayoutKind.Sequential)]
        internal struct Msg
        {
            public readonly IntPtr Hwnd;
            public readonly uint Message;
            public readonly UIntPtr WParam;
            public readonly IntPtr LParam;
            public readonly uint Time;
            public readonly int X;
            public readonly int Y;
            public readonly uint LPrivate;
        }
    }
#endif
}
