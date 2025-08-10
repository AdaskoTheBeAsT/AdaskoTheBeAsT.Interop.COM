using System;
using System.Runtime.InteropServices;
using System.Security;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable S101
#if NETSTANDARD2_0
[SuppressUnmanagedCodeSecurity]
internal static class NativeMethods
{
    internal const int ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID = 0x004;

    // Pump Windows messages for STA COM callbacks
    internal const uint PM_REMOVE = 0x0001;

    [DllImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool ActivateActCtx(IntPtr hActCtx, out IntPtr lpCookie);

    [DllImport("Kernel32.dll", EntryPoint = "CreateActCtxW", SetLastError = true, CharSet = CharSet.Unicode)]
    internal static extern IntPtr CreateActCtx(ref ActCtx pActCtx);

    [DllImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool DeactivateActCtx(int dwFlags, IntPtr lpCookie);

    [DllImport("Kernel32.dll", SetLastError = true)]
    internal static extern void ReleaseActCtx(IntPtr hActCtx);

    [DllImport("user32.dll")]
    internal static extern bool PeekMessage(
        out MSG lpMsg,
        IntPtr hWnd,
        uint wMsgFilterMin,
        uint wMsgFilterMax,
        uint wRemoveMsg);

    [DllImport("user32.dll")]
    internal static extern bool TranslateMessage([In] ref MSG lpMsg);

    [DllImport("user32.dll")]
    internal static extern IntPtr DispatchMessage([In] ref MSG lpMsg);

    internal static void PumpPendingMessages()
    {
        MSG msg;
        while (PeekMessage(out msg, IntPtr.Zero, 0, 0, PM_REMOVE))
        {
            TranslateMessage(ref msg);
            DispatchMessage(ref msg);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public UIntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public POINT pt;
        public uint lPrivate;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINT
    {
        public int x;
        public int y;
    }
}
#endif
#if NET8_0_OR_GREATER
[SuppressUnmanagedCodeSecurity]
internal static partial class NativeMethods
{
    internal const int ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID = 0x004;

    // Pump Windows messages for STA COM callbacks
    internal const uint PM_REMOVE = 0x0001;

    [LibraryImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool ActivateActCtx(IntPtr hActCtx, out IntPtr lpCookie);

    [DllImport("Kernel32.dll", EntryPoint = "CreateActCtxW", SetLastError = true)]
    internal static extern IntPtr CreateActCtx(ref ActCtx pActCtx);

    [LibraryImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool DeactivateActCtx(int dwFlags, IntPtr lpCookie);

    [LibraryImport("Kernel32.dll", SetLastError = true)]
    internal static partial void ReleaseActCtx(IntPtr hActCtx);

    [LibraryImport("user32.dll", EntryPoint = "PeekMessageW")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool PeekMessage(
        out MSG lpMsg,
        IntPtr hWnd,
        uint wMsgFilterMin,
        uint wMsgFilterMax,
        uint wRemoveMsg);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool TranslateMessage(ref MSG lpMsg);

    [LibraryImport("user32.dll", EntryPoint = "DispatchMessageW")]
    internal static partial IntPtr DispatchMessage(ref MSG lpMsg);

    internal static void PumpPendingMessages()
    {
        MSG msg;
        while (PeekMessage(out msg, IntPtr.Zero, 0, 0, PM_REMOVE))
        {
            TranslateMessage(ref msg);
            DispatchMessage(ref msg);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public UIntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public POINT pt;
        public uint lPrivate;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINT
    {
        public int x;
        public int y;
    }
}
#endif
#pragma warning restore S101
