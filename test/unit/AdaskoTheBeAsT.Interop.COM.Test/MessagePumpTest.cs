using System.Runtime.InteropServices;
using System.Security;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

public partial class MessagePumpTest
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
