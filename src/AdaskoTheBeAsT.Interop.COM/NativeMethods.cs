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
    internal const uint LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR = 0x00000100;
    internal const uint LOAD_LIBRARY_SEARCH_DEFAULT_DIRS = 0x00001000;
    internal const uint LOAD_LIBRARY_SEARCH_DLL_AND_DEFAULT_DIRS = LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS;

    internal const uint GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT = 0x00000002;
    internal const uint GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS = 0x00000004;

    internal const uint MEM_COMMIT = 0x00001000;
    internal const uint MEM_RESERVE = 0x00002000;
    internal const uint PAGE_READWRITE = 0x04;
    internal const uint PAGE_EXECUTE_READ = 0x20;

    internal const uint CLSCTX_INPROC_SERVER = 0x1;
    internal const uint CLSCTX_INPROC_HANDLER = 0x2;
    internal const uint CLSCTX_LOCAL_SERVER = 0x4;
    internal const uint CLSCTX_SERVER = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER;
    internal const uint CLSCTX_ALL = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER | CLSCTX_LOCAL_SERVER;

    internal const uint REGCLS_MULTIPLEUSE = 0x1;
    internal const uint COINIT_MULTITHREADED = 0x0;
    internal const uint COINIT_APARTMENTTHREADED = 0x2;

    internal const int REGKIND_NONE = 2;

    internal const int APTTYPE_CURRENT = -1;
    internal const int APTTYPE_STA = 0;
    internal const int APTTYPE_MTA = 1;
    internal const int APTTYPE_NA = 2;
    internal const int APTTYPE_MAINSTA = 3;

    internal const int APTTYPEQUALIFIER_NONE = 0;
    internal const int APTTYPEQUALIFIER_IMPLICIT_MTA = 1;
    internal const int APTTYPEQUALIFIER_NA_ON_MTA = 2;
    internal const int APTTYPEQUALIFIER_NA_ON_STA = 3;
    internal const int APTTYPEQUALIFIER_NA_ON_IMPLICIT_MTA = 4;
    internal const int APTTYPEQUALIFIER_NA_ON_MAINSTA = 5;

    internal const int S_OK = 0;
    internal const int S_FALSE = 1;
    internal const int E_NOTIMPL = unchecked((int)0x80004001);
    internal const int E_NOINTERFACE = unchecked((int)0x80004002);
    internal const int E_POINTER = unchecked((int)0x80004003);
    internal const int E_UNEXPECTED = unchecked((int)0x8000FFFF);
    internal const int DISP_E_TYPEMISMATCH = unchecked((int)0x80020005);
    internal const int DISP_E_OVERFLOW = unchecked((int)0x8002000A);
    internal const int DISP_E_BADINDEX = unchecked((int)0x8002000B);
    internal const int DISP_E_DIVBYZERO = unchecked((int)0x80020012);
    internal const int E_OUTOFMEMORY = unchecked((int)0x8007000E);
    internal const int E_INVALIDARG = unchecked((int)0x80070057);
    internal const int RPC_E_CHANGED_MODE = unchecked((int)0x80010106);
    internal const int STG_E_READFAULT = unchecked((int)0x8003001E);
    internal const int CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);
    internal const int CLASS_E_CLASSNOTAVAILABLE = unchecked((int)0x80040111);
    internal const int REGDB_E_CLASSNOTREG = unchecked((int)0x80040154);
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

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    internal delegate int DllCanUnloadNowDelegate();

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    internal delegate int ClassFactoryCreateInstanceDelegate(
        IntPtr thisPtr,
        IntPtr pUnkOuter,
        ref Guid riid,
        out IntPtr ppvObject);

    internal static void PumpPendingMessages(int maxMessages = MaxMessagesPerPump)
    {
        if (maxMessages <= 0)
        {
            return;
        }

        NanoServerOptions.EnsureInitialized();
        try
        {
            for (var count = 0; count < maxMessages && PeekMessage(out var msg, IntPtr.Zero, 0, 0, PM_REMOVE); count++)
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

    internal static int QueryInterface(IntPtr unk, ref Guid iid, out IntPtr ppv)
    {
#if NET8_0_OR_GREATER
        return Marshal.QueryInterface(unk, in iid, out ppv);
#else
        return Marshal.QueryInterface(unk, ref iid, out ppv);
#endif
    }

    internal static string? TryGetModuleFilePath(IntPtr moduleHandle)
    {
        if (moduleHandle == IntPtr.Zero)
        {
            return null;
        }

        const uint maxCharacters = 4096;
        var buffer = Marshal.AllocHGlobal((int)(maxCharacters * sizeof(char)));
        try
        {
            var length = GetModuleFileNameW(moduleHandle, buffer, maxCharacters);
            return length == 0 || length >= maxCharacters
                ? null
                : Marshal.PtrToStringUni(buffer, (int)length);
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
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

    [StructLayout(LayoutKind.Sequential)]
    internal readonly struct MEMORY_BASIC_INFORMATION
    {
        public readonly IntPtr BaseAddress;
        public readonly IntPtr AllocationBase;
        public readonly uint AllocationProtect;
        public readonly UIntPtr RegionSize;
        public readonly uint State;
        public readonly uint Protect;
        public readonly uint Type;
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

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    internal static extern uint GetThreadLocale();

    [DllImport("Kernel32.dll", SetLastError = true, ExactSpelling = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetThreadLocale(uint locale);

    [DllImport("Kernel32.dll", SetLastError = true)]
    internal static extern void ReleaseActCtx(IntPtr hActCtx);

    [DllImport("Kernel32.dll", EntryPoint = "LoadLibraryExW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true)]
    internal static extern IntPtr LoadLibraryEx(string lpLibFileName, IntPtr hFile, uint dwFlags);

    [DllImport("Kernel32.dll", SetLastError = true, ExactSpelling = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool FreeLibrary(IntPtr hLibModule);

    [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true)]
    internal static extern IntPtr GetModuleHandleW(string? lpModuleName);

    [DllImport("Kernel32.dll", EntryPoint = "GetModuleHandleExW", SetLastError = true, ExactSpelling = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool GetModuleHandleExFromAddress(uint dwFlags, IntPtr lpModuleName, out IntPtr phModule);

    [DllImport("Kernel32.dll", EntryPoint = "GetModuleHandleExW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool GetModuleHandleExByName(uint dwFlags, string lpModuleName, out IntPtr phModule);

    [DllImport("Kernel32.dll", SetLastError = true, ExactSpelling = true)]
    internal static extern uint GetModuleFileNameW(IntPtr hModule, IntPtr lpFilename, uint nSize);

    [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, BestFitMapping = false, ThrowOnUnmappableChar = true)]
    internal static extern IntPtr GetProcAddress(IntPtr hModule, [MarshalAs(UnmanagedType.LPStr)] string lpProcName);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern void CoUninitialize();

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern int CoIncrementMTAUsage(out IntPtr pCookie);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern int CoRegisterClassObject(
        ref Guid rclsid,
        IntPtr pUnk,
        uint dwClsContext,
        uint flags,
        out uint lpdwRegister);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern int CoRevokeClassObject(uint dwRegister);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern int CoGetClassObject(
        ref Guid rclsid,
        uint dwClsContext,
        IntPtr pvReserved,
        ref Guid riid,
        out IntPtr ppv);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern int CoCreateInstance(
        ref Guid rclsid,
        IntPtr pUnkOuter,
        uint dwClsContext,
        ref Guid riid,
        out IntPtr ppv);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern void CoFreeUnusedLibrariesEx(uint dwUnloadDelay, uint dwReserved);

    [DllImport("combase.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
    internal static extern int CLSIDFromProgID(string lpszProgID, out Guid lpclsid);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern int ProgIDFromCLSID(ref Guid clsid, out IntPtr lpszProgID);

    [DllImport("combase.dll", ExactSpelling = true)]
    internal static extern int CoGetApartmentType(out int pAptType, out int pAptQualifier);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern int GetErrorInfo(uint dwReserved, out IntPtr pperrinfo);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern int SetErrorInfo(uint dwReserved, IntPtr perrinfo);

    [DllImport("oleaut32.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
    internal static extern int LoadTypeLibEx(string szFile, int regkind, out IntPtr pptlib);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern int LoadRegTypeLib(ref Guid rguid, ushort wVerMajor, ushort wVerMinor, uint lcid, out IntPtr pptlib);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern uint SafeArrayGetDim(IntPtr psa);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern int SafeArrayGetLBound(IntPtr psa, uint nDim, out int plLbound);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern int SafeArrayGetUBound(IntPtr psa, uint nDim, out int plUbound);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern uint SafeArrayGetElemsize(IntPtr psa);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern int SafeArrayGetVartype(IntPtr psa, out ushort pvt);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern int SafeArrayAccessData(IntPtr psa, out IntPtr ppvData);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern int SafeArrayUnaccessData(IntPtr psa);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern int SafeArrayDestroy(IntPtr psa);

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    internal static extern IntPtr SafeArrayCreateVector(ushort vt, int lLbound, uint cElements);

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

    [DllImport("Kernel32.dll", ExactSpelling = true, SetLastError = true)]
    internal static extern IntPtr VirtualAlloc(
        IntPtr lpAddress,
        UIntPtr dwSize,
        uint flAllocationType,
        uint flProtect);

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool VirtualProtect(
        IntPtr lpAddress,
        UIntPtr dwSize,
        uint flNewProtect,
        out uint lpflOldProtect);

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    internal static extern UIntPtr VirtualQuery(
        IntPtr lpAddress,
        out MEMORY_BASIC_INFORMATION lpBuffer,
        UIntPtr dwLength);

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool FlushInstructionCache(
        IntPtr hProcess,
        IntPtr lpBaseAddress,
        UIntPtr dwSize);

    [DllImport("Kernel32.dll", ExactSpelling = true)]
    internal static extern IntPtr GetCurrentProcess();

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

    [LibraryImport("Kernel32.dll")]
    internal static partial uint GetThreadLocale();

    [LibraryImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool SetThreadLocale(uint locale);

    [LibraryImport("Kernel32.dll", SetLastError = true)]
    internal static partial void ReleaseActCtx(IntPtr hActCtx);

    [LibraryImport("Kernel32.dll", EntryPoint = "LoadLibraryExW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    internal static partial IntPtr LoadLibraryEx(string lpLibFileName, IntPtr hFile, uint dwFlags);

    [LibraryImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool FreeLibrary(IntPtr hLibModule);

    [LibraryImport("Kernel32.dll", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    internal static partial IntPtr GetModuleHandleW(string? lpModuleName);

    [LibraryImport("Kernel32.dll", EntryPoint = "GetModuleHandleExW", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool GetModuleHandleExFromAddress(uint dwFlags, IntPtr lpModuleName, out IntPtr phModule);

    [LibraryImport("Kernel32.dll", EntryPoint = "GetModuleHandleExW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool GetModuleHandleExByName(uint dwFlags, string lpModuleName, out IntPtr phModule);

    [LibraryImport("Kernel32.dll", SetLastError = true)]
    internal static partial uint GetModuleFileNameW(IntPtr hModule, IntPtr lpFilename, uint nSize);

    [LibraryImport("Kernel32.dll", SetLastError = true, StringMarshalling = StringMarshalling.Utf8)]
    internal static partial IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

    [LibraryImport("combase.dll")]
    internal static partial int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

    [LibraryImport("combase.dll")]
    internal static partial void CoUninitialize();

    [LibraryImport("combase.dll")]
    internal static partial int CoIncrementMTAUsage(out IntPtr pCookie);

    [LibraryImport("combase.dll")]
    internal static partial int CoRegisterClassObject(
        ref Guid rclsid,
        IntPtr pUnk,
        uint dwClsContext,
        uint flags,
        out uint lpdwRegister);

    [LibraryImport("combase.dll")]
    internal static partial int CoRevokeClassObject(uint dwRegister);

    [LibraryImport("combase.dll")]
    internal static partial int CoGetClassObject(
        ref Guid rclsid,
        uint dwClsContext,
        IntPtr pvReserved,
        ref Guid riid,
        out IntPtr ppv);

    [LibraryImport("combase.dll")]
    internal static partial int CoCreateInstance(
        ref Guid rclsid,
        IntPtr pUnkOuter,
        uint dwClsContext,
        ref Guid riid,
        out IntPtr ppv);

    [LibraryImport("combase.dll")]
    internal static partial void CoFreeUnusedLibrariesEx(uint dwUnloadDelay, uint dwReserved);

    [LibraryImport("combase.dll", StringMarshalling = StringMarshalling.Utf16)]
    internal static partial int CLSIDFromProgID(string lpszProgID, out Guid lpclsid);

    [LibraryImport("combase.dll")]
    internal static partial int ProgIDFromCLSID(ref Guid clsid, out IntPtr lpszProgID);

    [LibraryImport("combase.dll")]
    internal static partial int CoGetApartmentType(out int pAptType, out int pAptQualifier);

    [LibraryImport("oleaut32.dll")]
    internal static partial int GetErrorInfo(uint dwReserved, out IntPtr pperrinfo);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SetErrorInfo(uint dwReserved, IntPtr perrinfo);

    [LibraryImport("oleaut32.dll", StringMarshalling = StringMarshalling.Utf16)]
    internal static partial int LoadTypeLibEx(string szFile, int regkind, out IntPtr pptlib);

    [LibraryImport("oleaut32.dll")]
    internal static partial int LoadRegTypeLib(ref Guid rguid, ushort wVerMajor, ushort wVerMinor, uint lcid, out IntPtr pptlib);

    [LibraryImport("oleaut32.dll")]
    internal static partial uint SafeArrayGetDim(IntPtr psa);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayGetLBound(IntPtr psa, uint nDim, out int plLbound);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayGetUBound(IntPtr psa, uint nDim, out int plUbound);

    [LibraryImport("oleaut32.dll")]
    internal static partial uint SafeArrayGetElemsize(IntPtr psa);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayGetVartype(IntPtr psa, out ushort pvt);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayAccessData(IntPtr psa, out IntPtr ppvData);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayUnaccessData(IntPtr psa);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayDestroy(IntPtr psa);

    [LibraryImport("oleaut32.dll")]
    internal static partial IntPtr SafeArrayCreateVector(ushort vt, int lLbound, uint cElements);

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

    [LibraryImport("Kernel32.dll", SetLastError = true)]
    internal static partial IntPtr VirtualAlloc(
        IntPtr lpAddress,
        UIntPtr dwSize,
        uint flAllocationType,
        uint flProtect);

    [LibraryImport("Kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool VirtualProtect(
        IntPtr lpAddress,
        UIntPtr dwSize,
        uint flNewProtect,
        out uint lpflOldProtect);

    [LibraryImport("Kernel32.dll")]
    internal static partial UIntPtr VirtualQuery(
        IntPtr lpAddress,
        out MEMORY_BASIC_INFORMATION lpBuffer,
        UIntPtr dwLength);

    [LibraryImport("Kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool FlushInstructionCache(
        IntPtr hProcess,
        IntPtr lpBaseAddress,
        UIntPtr dwSize);

    [LibraryImport("Kernel32.dll")]
    internal static partial IntPtr GetCurrentProcess();

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
