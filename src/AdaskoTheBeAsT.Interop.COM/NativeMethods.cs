using System;
using System.Runtime.InteropServices;
using System.Security;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable S101
#pragma warning disable SA1121 // Use built-in type alias
[SuppressUnmanagedCodeSecurity]
internal static partial class NativeMethods
{
    internal const int ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID = 0x004;

    internal const int ERROR_INVALID_FUNCTION = 1;
    internal const int ERROR_FILE_NOT_FOUND = 2;
    internal const int ERROR_NOT_SUPPORTED = 50;
    internal const int ERROR_INVALID_PARAMETER = 87;
    internal const int ERROR_CALL_NOT_IMPLEMENTED = 120;
    internal const int ERROR_MOD_NOT_FOUND = 126;
    internal const int ERROR_PROC_NOT_FOUND = 127;
    internal const int ERROR_SXS_CANT_GEN_ACTCTX = 14001;

    internal const uint LOAD_WITH_ALTERED_SEARCH_PATH = 0x00000008;
    internal const uint CLSCTX_INPROC_SERVER = 0x1;
    internal const uint REGCLS_MULTIPLEUSE = 0x1;
    internal const uint COINIT_MULTITHREADED = 0x0;
    internal const uint COINIT_APARTMENTTHREADED = 0x2;

    internal const int E_NOINTERFACE = unchecked((int)0x80004002);
    internal const int CLASS_E_CLASSNOTAVAILABLE = unchecked((int)0x80040111);
    internal const int CO_E_NOTINITIALIZED = unchecked((int)0x800401F0);

    // Pump Windows messages for STA COM callbacks
    internal const uint PM_REMOVE = 0x0001;

    internal const int MaxMessagesPerPump = 256;
    private const uint WmQuit = 0x0012;

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    internal delegate int DllGetClassObjectDelegate(
        ref Guid rclsid,
        ref Guid riid,
        out IntPtr ppv);

    internal static void PumpPendingMessages()
    {
        NanoServerOptions.EnsureInitialized();
        try
        {
            for (var count = 0; count < MaxMessagesPerPump && PeekMessage(out var msg, IntPtr.Zero, 0, 0, PM_REMOVE); count++)
            {
                if (msg.message == WmQuit)
                {
                    PostQuitMessage(unchecked((int)msg.wParam.ToUInt64()));
                    break;
                }

                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }
        }
        catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException)
        {
            // Headless Windows editions such as Nano Server omit user32.dll / win32u.dll message queues.
            _ = ex;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MSG
    {
        public readonly IntPtr hwnd;
        public readonly uint message;
        public readonly UIntPtr wParam;
        public readonly IntPtr lParam;
        public readonly uint time;
        public readonly POINT pt;
        public readonly uint lPrivate;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINT
    {
        public readonly int x;
        public readonly int y;
    }
}

#if !NET8_0_OR_GREATER
internal static partial class NativeMethods
{
    [DllImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool ActivateActCtx(IntPtr hActCtx, out IntPtr lpCookie);

    [DllImport("Kernel32.dll", EntryPoint = "CreateActCtxW", SetLastError = true, CharSet = CharSet.Unicode)]
    internal static extern IntPtr CreateActCtx(ref ActCtx pActCtx);

    [DllImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool DeactivateActCtx(int dwFlags, IntPtr lpCookie);

    [DllImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool GetCurrentActCtx(out IntPtr hActCtx);

    [DllImport("Kernel32.dll")]
    internal static extern uint GetCurrentThreadId();

    [DllImport("Kernel32.dll", SetLastError = true)]
    internal static extern void ReleaseActCtx(IntPtr hActCtx);

    [DllImport("Kernel32.dll", EntryPoint = "LoadLibraryExW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true)]
    internal static extern IntPtr LoadLibraryEx(string lpLibFileName, IntPtr hFile, uint dwFlags);

    [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, BestFitMapping = false, ThrowOnUnmappableChar = true)]
    internal static extern IntPtr GetProcAddress(IntPtr hModule, [MarshalAs(UnmanagedType.LPStr)] string lpProcName);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern void CoUninitialize();

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern int CoRegisterClassObject(
        ref Guid rclsid,
        IntPtr pUnk,
        uint dwClsContext,
        uint flags,
        out uint lpdwRegister);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern int CoRevokeClassObject(uint dwRegister);

    [DllImport("user32.dll", EntryPoint = "PeekMessageW", ExactSpelling = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool PeekMessage(
        out MSG lpMsg,
        IntPtr hWnd,
        uint wMsgFilterMin,
        uint wMsgFilterMax,
        uint wRemoveMsg);

    [DllImport("user32.dll")]
    internal static extern bool TranslateMessage([In] ref MSG lpMsg);

    [DllImport("user32.dll", EntryPoint = "DispatchMessageW", ExactSpelling = true)]
    internal static extern IntPtr DispatchMessage([In] ref MSG lpMsg);

    [DllImport("user32.dll")]
    internal static extern void PostQuitMessage(int exitCode);

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool VirtualProtect(
        IntPtr lpAddress,
        UIntPtr dwSize,
        uint flNewProtect,
        out uint lpflOldProtect);

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    internal static extern void SetLastError(int dwErrCode);

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    internal static extern uint GetCurrentProcessId();

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    internal static extern IntPtr CreateEventW(
        IntPtr lpEventAttributes,
        [MarshalAs(UnmanagedType.Bool)] bool bManualReset,
        [MarshalAs(UnmanagedType.Bool)] bool bInitialState,
        IntPtr lpName);

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetEvent(IntPtr hEvent);

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool ResetEvent(IntPtr hEvent);

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    internal static extern uint WaitForSingleObject(IntPtr hHandle, uint dwMilliseconds);

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    internal static extern uint WaitForMultipleObjectsEx(
        uint nCount,
        [In] IntPtr[] lpHandles,
        [MarshalAs(UnmanagedType.Bool)] bool bWaitAll,
        uint dwMilliseconds,
        [MarshalAs(UnmanagedType.Bool)] bool bAlertable);
}
#endif
#if NET8_0_OR_GREATER
internal static partial class NativeMethods
{
    [LibraryImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool ActivateActCtx(IntPtr hActCtx, out IntPtr lpCookie);

#pragma warning disable SYSLIB1054
    [DllImport("Kernel32.dll", EntryPoint = "CreateActCtxW", SetLastError = true)]
    internal static extern IntPtr CreateActCtx(ref ActCtx pActCtx);
#pragma warning restore SYSLIB1054

    [LibraryImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool DeactivateActCtx(int dwFlags, IntPtr lpCookie);

    [LibraryImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool GetCurrentActCtx(out IntPtr hActCtx);

    [LibraryImport("Kernel32.dll")]
    internal static partial uint GetCurrentThreadId();

    [LibraryImport("Kernel32.dll", SetLastError = true)]
    internal static partial void ReleaseActCtx(IntPtr hActCtx);

    [LibraryImport("Kernel32.dll", EntryPoint = "LoadLibraryExW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    internal static partial IntPtr LoadLibraryEx(string lpLibFileName, IntPtr hFile, uint dwFlags);

    [LibraryImport("Kernel32.dll", SetLastError = true, StringMarshalling = StringMarshalling.Utf8)]
    internal static partial IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

    [LibraryImport("combase.dll")]
    internal static partial int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

    [LibraryImport("combase.dll")]
    internal static partial void CoUninitialize();

    [LibraryImport("combase.dll")]
    internal static partial int CoRegisterClassObject(
        ref Guid rclsid,
        IntPtr pUnk,
        uint dwClsContext,
        uint flags,
        out uint lpdwRegister);

    [LibraryImport("combase.dll")]
    internal static partial int CoRevokeClassObject(uint dwRegister);

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

    [LibraryImport("user32.dll")]
    internal static partial void PostQuitMessage(int exitCode);

    [LibraryImport("Kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool VirtualProtect(
        IntPtr lpAddress,
        UIntPtr dwSize,
        uint flNewProtect,
        out uint lpflOldProtect);

    [LibraryImport("Kernel32.dll")]
    internal static partial void SetLastError(int dwErrCode);

    [LibraryImport("Kernel32.dll")]
    internal static partial uint GetCurrentProcessId();

    [LibraryImport("Kernel32.dll")]
    internal static partial IntPtr CreateEventW(
        IntPtr lpEventAttributes,
        [MarshalAs(UnmanagedType.Bool)] bool bManualReset,
        [MarshalAs(UnmanagedType.Bool)] bool bInitialState,
        IntPtr lpName);

    [LibraryImport("Kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool SetEvent(IntPtr hEvent);

    [LibraryImport("Kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool ResetEvent(IntPtr hEvent);

    [LibraryImport("Kernel32.dll")]
    internal static partial uint WaitForSingleObject(IntPtr hHandle, uint dwMilliseconds);

    [LibraryImport("Kernel32.dll")]
    internal static partial uint WaitForMultipleObjectsEx(
        uint nCount,
        [In] IntPtr[] lpHandles,
        [MarshalAs(UnmanagedType.Bool)] bool bWaitAll,
        uint dwMilliseconds,
        [MarshalAs(UnmanagedType.Bool)] bool bAlertable);
}
#endif
#pragma warning restore SA1121 // Use built-in type alias
#pragma warning restore S101
