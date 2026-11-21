using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable MA0158 // Use System.Threading.Lock
#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable S101
#pragma warning disable SCS0018, SEC0116 // Protected system-directory check uses OS-controlled folder paths.
[SuppressUnmanagedCodeSecurity]
internal static class NanoServerOle2Shim
{
    private const ushort Pe32Magic = 0x10B;
    private const ushort Pe32PlusMagic = 0x20B;
    private const int ImportDirectoryIndex = 1;
    private const int DelayImportDirectoryIndex = 13;
    private const uint StubPageSize = 4096;
    private const string Ole32DllName = "ole32.dll";

#pragma warning disable MA0176 // Optimize guid creation
    private static readonly Guid IidIRunnableObject = new("00000126-0000-0000-C000-000000000046");
    private static readonly Guid IidIPersistStream = new("00000109-0000-0000-C000-000000000046");
#pragma warning restore MA0176

#if NET10_0_OR_GREATER
    private static readonly Lock Ole2Lock = new();
#endif
#if NET472_OR_GREATER
    private static readonly object Ole2Lock = new();
#endif

    // Static fields root delegates for the entire managed process lifetime; patched IAT slots are restored on ProcessExit.
    private static readonly CoInitializeDelegate CoInitializeInstance = CoInitializeImpl;
    private static readonly OleInitializeDelegate OleInitializeInstance = OleInitializeImpl;
    private static readonly OleUninitializeDelegate OleUninitializeInstance = OleUninitializeImpl;
    private static readonly OleRunDelegate OleRunInstance = OleRunImpl;
    private static readonly OleSaveToStreamDelegate OleSaveToStreamInstance = OleSaveToStreamImpl;
    private static readonly OleLoadFromStreamDelegate OleLoadFromStreamInstance = OleLoadFromStreamImpl;

    private static readonly Dictionary<string, IntPtr> FallbackExports = new(StringComparer.Ordinal)
    {
        ["CoInitialize"] = Marshal.GetFunctionPointerForDelegate(CoInitializeInstance),
        ["OleInitialize"] = Marshal.GetFunctionPointerForDelegate(OleInitializeInstance),
        ["OleUninitialize"] = Marshal.GetFunctionPointerForDelegate(OleUninitializeInstance),
        ["OleRun"] = Marshal.GetFunctionPointerForDelegate(OleRunInstance),
        ["OleSaveToStream"] = Marshal.GetFunctionPointerForDelegate(OleSaveToStreamInstance),
        ["OleLoadFromStream"] = Marshal.GetFunctionPointerForDelegate(OleLoadFromStreamInstance),
    };

    private static readonly Dictionary<IntPtr, List<PatchedImportSlot>> PatchedModules = [];
    private static readonly Dictionary<IntPtr, List<PatchedImportSlot>> SuppressedUninitializeModules = [];

    private static readonly HashSet<string> ProtectedModuleFileNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "combase.dll",
        Ole32DllName,
        "oleaut32.dll",
        "rpcrt4.dll",
        "ntdll.dll",
        "kernel32.dll",
        "kernelbase.dll",
        "user32.dll",
        "advapi32.dll",
        "clr.dll",
        "coreclr.dll",
        "clrjit.dll",
        "hostfxr.dll",
        "hostpolicy.dll",
        "mscorwks.dll",
        "mscoree.dll",
        "mscoreei.dll",
        "nanoserver_sta_shim.dll",
    };

    [ThreadStatic]
    private static int _threadOleInitDepth;

    private static int _activeProvider;
    private static bool _processExitHooked;
    private static int _forceInProcessForTesting;

    // Static-lifetime pure native W^X stub page (single 0xC3 `ret` instruction, valid for 0-argument stdcall/x64
    // CoUninitialize(void) and OleUninitialize(void)); never freed before ExitProcess.
    private static IntPtr _uninitializeNoOpStub;

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int CoInitializeDelegate(IntPtr pvReserved);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int OleInitializeDelegate(IntPtr pvReserved);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate void OleUninitializeDelegate();

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int OleRunDelegate(IntPtr pUnknown);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int OleSaveToStreamDelegate(IntPtr pPersistStream, IntPtr pStream);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int OleLoadFromStreamDelegate(IntPtr pStream, ref Guid iidInterface, out IntPtr ppvObj);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int RunnableObjectRunDelegate(IntPtr thisPtr, IntPtr pBindCtx);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int PersistGetClassIdDelegate(IntPtr thisPtr, out Guid clsid);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int PersistStreamLoadDelegate(IntPtr thisPtr, IntPtr pStream);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int PersistStreamSaveDelegate(IntPtr thisPtr, IntPtr pStream, int fClearDirty);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int SequentialStreamReadDelegate(IntPtr thisPtr, IntPtr pv, uint cb, out uint pcbRead);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int SequentialStreamWriteDelegate(IntPtr thisPtr, IntPtr pv, uint cb, out uint pcbWritten);

    internal enum Ole2ProviderKind
    {
        None = 0,
        NanoOsShim = 1,
        NativeOs = 2,
        InProcessFallback = 3,
    }

    internal static Ole2ProviderKind ActiveProvider
        => (Ole2ProviderKind)Volatile.Read(ref _activeProvider);

    internal static bool IsInProcessFallbackActive => ActiveProvider == Ole2ProviderKind.InProcessFallback;

    internal static IntPtr UninitializeNoOpStubAddress => Volatile.Read(ref _uninitializeNoOpStub);

    internal static bool ForceInProcessForTesting
    {
        get
        {
            return Volatile.Read(ref _forceInProcessForTesting) != 0;
        }

        set
        {
            lock (Ole2Lock)
            {
                Volatile.Write(ref _forceInProcessForTesting, value ? 1 : 0);
                RefreshStateCore();
            }
        }
    }

    internal static void RefreshState()
    {
        if (Environment.OSVersion.Platform != PlatformID.Win32NT)
        {
            return;
        }

        lock (Ole2Lock)
        {
            RefreshStateCore();
        }
    }

    internal static void OnModuleLoaded(IntPtr moduleHandle)
    {
        if (moduleHandle == IntPtr.Zero || !IsInProcessFallbackActive)
        {
            return;
        }

        lock (Ole2Lock)
        {
            if (IsInProcessFallbackActive)
            {
                PatchModuleCore(moduleHandle);
            }
        }
    }

    internal static int PatchModuleForTesting(IntPtr moduleHandle)
    {
        if (moduleHandle == IntPtr.Zero)
        {
            return 0;
        }

        lock (Ole2Lock)
        {
            return PatchModuleCore(moduleHandle);
        }
    }

    internal static void RestoreModuleForTesting(IntPtr moduleHandle)
    {
        lock (Ole2Lock)
        {
            if (PatchedModules.TryGetValue(moduleHandle, out var slots))
            {
                RestoreSlots(slots);
                PatchedModules.Remove(moduleHandle);
            }
        }
    }

    internal static int SuppressUninitializeImports(IntPtr moduleHandle)
    {
        if (moduleHandle == IntPtr.Zero)
        {
            throw new ArgumentException("Module handle cannot be zero.", nameof(moduleHandle));
        }

        if (!NativeMethods.GetModuleHandleExFromAddress(
                NativeMethods.GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | NativeMethods.GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                moduleHandle,
                out var resolvedBase)
            || resolvedBase != moduleHandle)
        {
            throw new ArgumentException(
                string.Format(CultureInfo.InvariantCulture, "Handle 0x{0:X} is not the base address of a loaded PE module.", moduleHandle.ToInt64()),
                nameof(moduleHandle));
        }

        ValidateTargetModuleNotProtected(moduleHandle);

        lock (Ole2Lock)
        {
            EnsureProcessExitHookedUnderLock();
            if (SuppressedUninitializeModules.TryGetValue(moduleHandle, out var existing))
            {
                return existing.Count;
            }

            var noOpStub = EnsureUninitializeNoOpStubUnderLock();
            var replacements = new Dictionary<string, IntPtr>(StringComparer.Ordinal)
            {
                ["CoUninitialize"] = noOpStub,
                ["OleUninitialize"] = noOpStub,
            };

            var patchedSlots = PatchModuleImportsCore(moduleHandle, IsComUninitializeImportModule, replacements);
            if (patchedSlots.Count > 0)
            {
                SuppressedUninitializeModules[moduleHandle] = patchedSlots;
            }

            return patchedSlots.Count;
        }
    }

    internal static int RestoreUninitializeImports(IntPtr moduleHandle)
    {
        if (moduleHandle == IntPtr.Zero)
        {
            throw new ArgumentException("Module handle cannot be zero.", nameof(moduleHandle));
        }

        lock (Ole2Lock)
        {
            if (!SuppressedUninitializeModules.TryGetValue(moduleHandle, out var slots))
            {
                return 0;
            }

            RestoreSlots(slots);
            SuppressedUninitializeModules.Remove(moduleHandle);
            return slots.Count;
        }
    }

    internal static void OnModuleUnloaded(IntPtr moduleHandle)
    {
        if (moduleHandle == IntPtr.Zero)
        {
            return;
        }

        if (NativeMethods.GetModuleHandleExFromAddress(
            NativeMethods.GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | NativeMethods.GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            moduleHandle,
            out var stillMapped)
            && stillMapped == moduleHandle)
        {
            return;
        }

        lock (Ole2Lock)
        {
            PatchedModules.Remove(moduleHandle);
            SuppressedUninitializeModules.Remove(moduleHandle);
        }
    }

    internal static Ole2ProviderKind QueryAvailableNativeProvider()
    {
        var osShimHandle = NanoServerStaShim.OsStaShimModuleHandle;
        if (osShimHandle == IntPtr.Zero)
        {
            osShimHandle = NativeMethods.GetModuleHandleW("nanoserver_sta_shim.dll");
        }

        if (osShimHandle != IntPtr.Zero && HasAllOle2Exports(osShimHandle, verifyOleRunBehavior: true))
        {
            return Ole2ProviderKind.NanoOsShim;
        }

        var ole32Handle = NativeMethods.GetModuleHandleW(Ole32DllName);
        if (ole32Handle == IntPtr.Zero)
        {
            ole32Handle = NativeMethods.LoadLibraryEx(Ole32DllName, IntPtr.Zero, 0);
        }

        if (ole32Handle != IntPtr.Zero && HasAllOle2Exports(ole32Handle, verifyOleRunBehavior: true))
        {
            return Ole2ProviderKind.NativeOs;
        }

        return Ole2ProviderKind.None;
    }

    internal static IntPtr ResolveFallbackExport(string functionName)
        => FallbackExports.TryGetValue(functionName, out var ptr) ? ptr : IntPtr.Zero;

    internal static int CoInitializeImpl(IntPtr pvReserved)
    {
        if (pvReserved != IntPtr.Zero)
        {
            return NativeMethods.E_INVALIDARG;
        }

        return NativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_APARTMENTTHREADED);
    }

    internal static int OleInitializeImpl(IntPtr pvReserved)
    {
        if (pvReserved != IntPtr.Zero)
        {
            return NativeMethods.E_INVALIDARG;
        }

        var hr = NativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_APARTMENTTHREADED);
        if (hr < 0)
        {
            return hr;
        }

        _threadOleInitDepth++;
        return _threadOleInitDepth == 1 ? NativeMethods.S_OK : NativeMethods.S_FALSE;
    }

    internal static void OleUninitializeImpl()
    {
        if (_threadOleInitDepth > 0)
        {
            _threadOleInitDepth--;
            NativeMethods.CoUninitialize();
        }
    }

    internal static int OleRunImpl(IntPtr pUnknown)
    {
        if (pUnknown == IntPtr.Zero)
        {
            return NativeMethods.E_INVALIDARG;
        }

        var iid = IidIRunnableObject;
        var qiHr = NativeMethods.QueryInterface(pUnknown, ref iid, out var runnablePtr);
        if (qiHr == NativeMethods.E_NOINTERFACE)
        {
            // Per Win32 OleRun specification: objects not implementing IRunnableObject are already running.
            return NativeMethods.S_OK;
        }

        if (qiHr < 0 || runnablePtr == IntPtr.Zero)
        {
            return qiHr < 0 ? qiHr : NativeMethods.E_NOINTERFACE;
        }

        try
        {
            var vtable = Marshal.ReadIntPtr(runnablePtr);
            var runMethodPtr = Marshal.ReadIntPtr(vtable, 4 * IntPtr.Size);
            var runFn = Marshal.GetDelegateForFunctionPointer<RunnableObjectRunDelegate>(runMethodPtr);
            return runFn(runnablePtr, IntPtr.Zero);
        }
        finally
        {
            _ = Marshal.Release(runnablePtr);
        }
    }

    internal static int OleSaveToStreamImpl(IntPtr pPersistStream, IntPtr pStream)
    {
        if (pPersistStream == IntPtr.Zero || pStream == IntPtr.Zero)
        {
            return NativeMethods.E_INVALIDARG;
        }

        var psVtable = Marshal.ReadIntPtr(pPersistStream);
        var getClassIdPtr = Marshal.ReadIntPtr(psVtable, 3 * IntPtr.Size);
        var getClassId = Marshal.GetDelegateForFunctionPointer<PersistGetClassIdDelegate>(getClassIdPtr);
        var hr = getClassId(pPersistStream, out var clsid);
        if (hr < 0)
        {
            return hr;
        }

        var guidBytes = clsid.ToByteArray();
        var guidBuffer = Marshal.AllocHGlobal(16);
        try
        {
            Marshal.Copy(guidBytes, 0, guidBuffer, 16);
            var stmVtable = Marshal.ReadIntPtr(pStream);
            var writePtr = Marshal.ReadIntPtr(stmVtable, 4 * IntPtr.Size);
            var writeFn = Marshal.GetDelegateForFunctionPointer<SequentialStreamWriteDelegate>(writePtr);
            hr = writeFn(pStream, guidBuffer, 16, out var written);
            if (hr < 0)
            {
                return hr;
            }

            if (written != 16)
            {
                return NativeMethods.E_UNEXPECTED;
            }
        }
        finally
        {
            Marshal.FreeHGlobal(guidBuffer);
        }

        var savePtr = Marshal.ReadIntPtr(psVtable, 6 * IntPtr.Size);
        var saveFn = Marshal.GetDelegateForFunctionPointer<PersistStreamSaveDelegate>(savePtr);
        return saveFn(pPersistStream, pStream, fClearDirty: 1);
    }

    internal static int OleLoadFromStreamImpl(IntPtr pStream, ref Guid iidInterface, out IntPtr ppvObj)
    {
        ppvObj = IntPtr.Zero;
        if (pStream == IntPtr.Zero)
        {
            return NativeMethods.E_INVALIDARG;
        }

        if (!TryReadClsidFromStream(pStream, out var clsid, out var readHr))
        {
            return readHr;
        }

        if (clsid == Guid.Empty)
        {
            return NativeMethods.S_OK;
        }

        var persistIid = IidIPersistStream;
        var createHr = NativeMethods.CoCreateInstance(
            ref clsid,
            IntPtr.Zero,
            NativeMethods.CLSCTX_ALL,
            ref persistIid,
            out var persistStreamPtr);
        if (createHr < 0 || persistStreamPtr == IntPtr.Zero)
        {
            return createHr < 0 ? createHr : NativeMethods.REGDB_E_CLASSNOTREG;
        }

        try
        {
            var psVtable = Marshal.ReadIntPtr(persistStreamPtr);
            var loadPtr = Marshal.ReadIntPtr(psVtable, 5 * IntPtr.Size);
            var loadFn = Marshal.GetDelegateForFunctionPointer<PersistStreamLoadDelegate>(loadPtr);
            var loadHr = loadFn(persistStreamPtr, pStream);
            if (loadHr < 0)
            {
                return loadHr;
            }

            return NativeMethods.QueryInterface(persistStreamPtr, ref iidInterface, out ppvObj);
        }
        finally
        {
            _ = Marshal.Release(persistStreamPtr);
        }
    }

    private static bool TryReadClsidFromStream(IntPtr pStream, out Guid clsid, out int hr)
    {
        clsid = Guid.Empty;
        var guidBuffer = Marshal.AllocHGlobal(16);
        try
        {
            var stmVtable = Marshal.ReadIntPtr(pStream);
            var readPtr = Marshal.ReadIntPtr(stmVtable, 3 * IntPtr.Size);
            var readFn = Marshal.GetDelegateForFunctionPointer<SequentialStreamReadDelegate>(readPtr);
            hr = readFn(pStream, guidBuffer, 16, out var bytesRead);
            if (hr < 0)
            {
                return false;
            }

            if (bytesRead != 16)
            {
                hr = NativeMethods.STG_E_READFAULT;
                return false;
            }

            var guidBytes = new byte[16];
            Marshal.Copy(guidBuffer, guidBytes, 0, 16);
            clsid = new Guid(guidBytes);
            return true;
        }
        finally
        {
            Marshal.FreeHGlobal(guidBuffer);
        }
    }

    private static void RefreshStateCore()
    {
        EnsureProcessExitHookedUnderLock();

        if (!NanoServerOptions.ReadEnableLegacyOle2ShimRaw() && !ForceInProcessForTesting)
        {
            RestoreAllModulesUnderLock();
            Volatile.Write(ref _activeProvider, (int)Ole2ProviderKind.None);
            return;
        }

        if (!ForceInProcessForTesting)
        {
            var nativeProvider = QueryAvailableNativeProvider();
            if (nativeProvider != Ole2ProviderKind.None)
            {
                // Capability query succeeded: prefer Nano OS shim or native Windows ole32.dll.
                RestoreAllModulesUnderLock();
                Volatile.Write(ref _activeProvider, (int)nativeProvider);
                return;
            }
        }

        Volatile.Write(ref _activeProvider, (int)Ole2ProviderKind.InProcessFallback);
    }

    private static bool HasAllOle2Exports(IntPtr moduleHandle, bool verifyOleRunBehavior)
    {
        var oleRunPtr = IntPtr.Zero;
        foreach (var exportName in FallbackExports.Keys)
        {
            var proc = NativeMethods.GetProcAddress(moduleHandle, exportName);
            if (proc == IntPtr.Zero)
            {
                return false;
            }

            if (string.Equals(exportName, "OleRun", StringComparison.Ordinal))
            {
                oleRunPtr = proc;
            }
        }

        if (!verifyOleRunBehavior || oleRunPtr == IntPtr.Zero)
        {
            return true;
        }

        // Probe OleRun(NULL): functional Win32/Nano shim implementations return E_INVALIDARG (0x80070057),
        // whereas unimplemented api-set stubs return E_NOTIMPL or ERROR_MOD_NOT_FOUND / ERROR_PROC_NOT_FOUND HRESULTs.
        var probe = Marshal.GetDelegateForFunctionPointer<OleRunDelegate>(oleRunPtr);
        var hr = probe(IntPtr.Zero);
        return hr == NativeMethods.E_INVALIDARG;
    }

    private static void EnsureProcessExitHookedUnderLock()
    {
        if (!_processExitHooked)
        {
            AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
            _processExitHooked = true;
        }
    }

    private static void ValidateTargetModuleNotProtected(IntPtr moduleHandle)
    {
        var modulePath = NativeMethods.TryGetModuleFilePath(moduleHandle);
        if (modulePath is null || string.IsNullOrWhiteSpace(modulePath))
        {
            return;
        }

        var fileName = Path.GetFileName(modulePath);
        if (ProtectedModuleFileNames.Contains(fileName) || IsUnderWindowsSystemDirectory(modulePath))
        {
            throw new InvalidOperationException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Refusing to patch CoUninitialize/OleUninitialize imports in protected system or CLR module '{0}'.",
                    modulePath));
        }
    }

    private static bool IsUnderWindowsSystemDirectory(string fullModulePath)
    {
        var normalizedModuleDir = Path.GetDirectoryName(Path.GetFullPath(fullModulePath));
        if (normalizedModuleDir is null || string.IsNullOrEmpty(normalizedModuleDir))
        {
            return false;
        }

        return IsSameOrChildDirectory(normalizedModuleDir, Environment.GetFolderPath(Environment.SpecialFolder.System))
            || IsSameOrChildDirectory(normalizedModuleDir, Environment.GetFolderPath(Environment.SpecialFolder.SystemX86));
    }

    private static bool IsSameOrChildDirectory(string candidateDir, string protectedRoot)
    {
        if (string.IsNullOrWhiteSpace(protectedRoot))
        {
            return false;
        }

        var normalizedCandidate = candidateDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedRoot = Path.GetFullPath(protectedRoot).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        if (string.Equals(normalizedCandidate, normalizedRoot, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return normalizedCandidate.StartsWith(
            normalizedRoot + Path.DirectorySeparatorChar,
            StringComparison.OrdinalIgnoreCase);
    }

    private static IntPtr EnsureUninitializeNoOpStubUnderLock()
    {
        if (_uninitializeNoOpStub != IntPtr.Zero)
        {
            return _uninitializeNoOpStub;
        }

        var page = NativeMethods.VirtualAlloc(
            IntPtr.Zero,
            new UIntPtr(StubPageSize),
            NativeMethods.MEM_COMMIT | NativeMethods.MEM_RESERVE,
            NativeMethods.PAGE_READWRITE);
        if (page == IntPtr.Zero)
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        // 0xC3 = near `ret` (0 bytes popped; valid for 0-arg stdcall on x86 and Microsoft x64 ABI),
        // padded with 0xCC (`int3`) traps.
        Marshal.WriteByte(page, 0, 0xC3);
        for (var i = 1; i < 16; i++)
        {
            Marshal.WriteByte(page, i, 0xCC);
        }

        _ = NativeMethods.FlushInstructionCache(NativeMethods.GetCurrentProcess(), page, new UIntPtr(16));

        // Transition PAGE_READWRITE (0x04) -> PAGE_EXECUTE_READ (0x20) to uphold strict W^X.
        if (!NativeMethods.VirtualProtect(page, new UIntPtr(StubPageSize), NativeMethods.PAGE_EXECUTE_READ, out _))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        Volatile.Write(ref _uninitializeNoOpStub, page);
        return page;
    }

    private static int PatchModuleCore(IntPtr moduleBase)
    {
        if (PatchedModules.TryGetValue(moduleBase, out var existing))
        {
            return existing.Count;
        }

        var patchedSlots = PatchModuleImportsCore(moduleBase, IsOle32ImportModule, FallbackExports);
        if (patchedSlots.Count > 0)
        {
            PatchedModules[moduleBase] = patchedSlots;
        }

        return patchedSlots.Count;
    }

    private static List<PatchedImportSlot> PatchModuleImportsCore(
        IntPtr moduleBase,
        Func<string?, bool> importModulePredicate,
        Dictionary<string, IntPtr> exportReplacements)
    {
        var patchedSlots = new List<PatchedImportSlot>();
        if (Marshal.ReadInt16(moduleBase, 0) != 0x5A4D)
        {
            return patchedSlots;
        }

        var eLfanew = Marshal.ReadInt32(moduleBase, 0x3C);
        var ntHeaders = AddRva(moduleBase, eLfanew);
        if (Marshal.ReadInt32(ntHeaders, 0) != 0x00004550)
        {
            return patchedSlots;
        }

        var optionalHeader = AddRva(ntHeaders, 24);
        var magic = unchecked((ushort)Marshal.ReadInt16(optionalHeader, 0));
        var dataDirOffset = magic switch
        {
            Pe32Magic => 96,
            Pe32PlusMagic => 112,
            _ => -1,
        };
        if (dataDirOffset < 0)
        {
            return patchedSlots;
        }

        var importRva = Marshal.ReadInt32(optionalHeader, dataDirOffset + (ImportDirectoryIndex * 8));
        if (importRva > 0)
        {
            PatchStandardImports(moduleBase, importRva, importModulePredicate, exportReplacements, patchedSlots);
        }

        var delayImportRva = Marshal.ReadInt32(optionalHeader, dataDirOffset + (DelayImportDirectoryIndex * 8));
        if (delayImportRva > 0)
        {
            PatchDelayImports(moduleBase, delayImportRva, importModulePredicate, exportReplacements, patchedSlots);
        }

        return patchedSlots;
    }

    private static void PatchStandardImports(
        IntPtr moduleBase,
        int importRva,
        Func<string?, bool> importModulePredicate,
        Dictionary<string, IntPtr> exportReplacements,
        List<PatchedImportSlot> patchedSlots)
    {
        for (var descOffset = 0; descOffset < 0x10000; descOffset += 20)
        {
            var descAddr = AddRva(moduleBase, importRva + descOffset);
            var originalFirstThunk = Marshal.ReadInt32(descAddr, 0);
            var nameRva = Marshal.ReadInt32(descAddr, 12);
            var firstThunk = Marshal.ReadInt32(descAddr, 16);
            if (originalFirstThunk == 0 && nameRva == 0 && firstThunk == 0)
            {
                break;
            }

            if (nameRva == 0 || firstThunk == 0)
            {
                continue;
            }

            var dllName = Marshal.PtrToStringAnsi(AddRva(moduleBase, nameRva));
#pragma warning disable CC0031 // Caller provides non-null predicate.
            if (!importModulePredicate(dllName))
#pragma warning restore CC0031
            {
                continue;
            }

            var intRva = originalFirstThunk != 0 ? originalFirstThunk : firstThunk;
            PatchThunkChain(moduleBase, intRva, firstThunk, exportReplacements, patchedSlots);
        }
    }

    private static void PatchDelayImports(
        IntPtr moduleBase,
        int delayImportRva,
        Func<string?, bool> importModulePredicate,
        Dictionary<string, IntPtr> exportReplacements,
        List<PatchedImportSlot> patchedSlots)
    {
        for (var descOffset = 0; descOffset < 0x10000; descOffset += 32)
        {
            var descAddr = AddRva(moduleBase, delayImportRva + descOffset);
            var rvaDllName = Marshal.ReadInt32(descAddr, 4);
            var rvaIat = Marshal.ReadInt32(descAddr, 12);
            var rvaInt = Marshal.ReadInt32(descAddr, 16);
            if (rvaDllName == 0)
            {
                break;
            }

            var dllName = Marshal.PtrToStringAnsi(AddRva(moduleBase, rvaDllName));
#pragma warning disable CC0031 // Caller provides non-null predicate.
            if (importModulePredicate(dllName) && rvaInt != 0 && rvaIat != 0)
#pragma warning restore CC0031
            {
                PatchThunkChain(moduleBase, rvaInt, rvaIat, exportReplacements, patchedSlots);
            }
        }
    }

    private static void PatchThunkChain(
        IntPtr moduleBase,
        int intRva,
        int iatRva,
        Dictionary<string, IntPtr> exportReplacements,
        List<PatchedImportSlot> patchedSlots)
    {
        for (var offset = 0; offset < 0x10000; offset += IntPtr.Size)
        {
            var thunkData = Marshal.ReadIntPtr(moduleBase, intRva + offset);
            if (thunkData == IntPtr.Zero)
            {
                break;
            }

            if (NanoServerStaShim.IsOrdinalImport(thunkData))
            {
                continue;
            }

            var nameRva = (int)(thunkData.ToInt64() & 0x7FFFFFFF);
            var funcName = Marshal.PtrToStringAnsi(AddRva(moduleBase, nameRva + 2));
            if (funcName is null || !exportReplacements.TryGetValue(funcName, out var replacementPtr))
            {
                continue;
            }

            var iatSlot = AddRva(moduleBase, iatRva + offset);
            var originalPtr = Marshal.ReadIntPtr(iatSlot);
            if (originalPtr != replacementPtr)
            {
                // Uses PAGE_READWRITE only (strict W^X).
                NanoServerStaShim.WritePointerWithProtection(iatSlot, replacementPtr);
                patchedSlots.Add(new PatchedImportSlot(iatSlot, originalPtr));
            }
        }
    }

    private static bool IsOle32ImportModule(string? dllName)
    {
        if (dllName is not { Length: > 0 })
        {
            return false;
        }

        return string.Equals(dllName, Ole32DllName, StringComparison.OrdinalIgnoreCase)
            || dllName.StartsWith("ext-ms-win-com-ole32-", StringComparison.OrdinalIgnoreCase)
            || dllName.StartsWith("api-ms-win-ole32-", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsComUninitializeImportModule(string? dllName)
    {
        if (dllName is not { Length: > 0 })
        {
            return false;
        }

        return IsOle32ImportModule(dllName)
            || string.Equals(dllName, "combase.dll", StringComparison.OrdinalIgnoreCase)
            || dllName.StartsWith("api-ms-win-core-com-l1-", StringComparison.OrdinalIgnoreCase);
    }

    private static IntPtr AddRva(IntPtr baseAddress, int rva)
        => new(baseAddress.ToInt64() + rva);

    private static void RestoreAllModulesUnderLock()
    {
        foreach (var pair in PatchedModules)
        {
            RestoreSlots(pair.Value);
        }

        PatchedModules.Clear();
    }

    private static void RestoreSlots(List<PatchedImportSlot> slots)
    {
        for (var i = 0; i < slots.Count; i++)
        {
            if (NativeMethods.GetModuleHandleExFromAddress(
                NativeMethods.GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | NativeMethods.GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                slots[i].IatSlot,
                out _))
            {
                NanoServerStaShim.WritePointerWithProtection(slots[i].IatSlot, slots[i].OriginalTarget);
            }
        }
    }

    private static void OnProcessExit(object? sender, EventArgs e)
    {
        _ = sender;
        _ = e;
        lock (Ole2Lock)
        {
            foreach (var pair in SuppressedUninitializeModules)
            {
                RestoreSlots(pair.Value);
            }

            SuppressedUninitializeModules.Clear();
            RestoreAllModulesUnderLock();
            Volatile.Write(ref _activeProvider, (int)Ole2ProviderKind.None);
        }
    }

    [StructLayout(LayoutKind.Auto)]
    private readonly struct PatchedImportSlot
    {
        internal PatchedImportSlot(IntPtr iatSlot, IntPtr originalTarget)
        {
            IatSlot = iatSlot;
            OriginalTarget = originalTarget;
        }

        internal IntPtr IatSlot { get; }

        internal IntPtr OriginalTarget { get; }
    }
}
#pragma warning restore SCS0018, SEC0116
#pragma warning restore S101
#pragma warning restore SA1121
#pragma warning restore MA0158
