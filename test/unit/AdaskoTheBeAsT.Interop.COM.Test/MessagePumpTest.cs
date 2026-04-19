using System.Runtime.InteropServices;
using System.Security;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

public class MessagePumpTest
{
    private const uint WmNull = 0x0000;
    private const uint PmNoRemove = 0x0000;

    [Fact]
    public void PumpPendingMessagesShouldTranslateAndDispatchQueuedMessage()
    {
        // Force the current thread to have a message queue.
        _ = NativeMethods.PeekMessage(out _, IntPtr.Zero, 0, 0, PmNoRemove);

        var threadId = NativeMethods.GetCurrentThreadId();
        var posted = NativeMethods.PostThreadMessage(threadId, WmNull, IntPtr.Zero, IntPtr.Zero);
        posted.Should().BeTrue(
            "PostThreadMessage must succeed so PumpPendingMessages has something to dispatch");

        var (comAssemblyPath, manifestPath) = GetPaths();

        // Executor.Execute calls NativeMethods.PumpPendingMessages at the end of its happy path,
        // so the WM_NULL we just posted is processed inside the context.
        var result = Executor.Execute(
            comAssemblyPath,
            manifestPath,
            () => { });

        result.Success.Should().BeTrue();

        // The queue should now be empty (PeekMessage with PM_NOREMOVE returns false).
        var hasLeftover = NativeMethods.PeekMessage(out _, IntPtr.Zero, 0, 0, PmNoRemove);
        hasLeftover.Should().BeFalse();
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
            public IntPtr Hwnd;
            public uint Message;
            public UIntPtr WParam;
            public IntPtr LParam;
            public uint Time;
            public int X;
            public int Y;
            public uint LPrivate;
        }
    }
}
