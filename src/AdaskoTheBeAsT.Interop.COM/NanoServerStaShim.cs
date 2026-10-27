using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable IDISP001 // Dispose created
#pragma warning disable IDISP003 // Dispose previous before re-assigning
#pragma warning disable IDISP006 // Implement IDisposable
#pragma warning disable IDISP007 // Don't dispose injected
#pragma warning disable MA0158 // Use System.Threading.Lock
#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable S101
#pragma warning disable SCS0018, SEC0116 // System directory paths are OS-controlled.
[SuppressUnmanagedCodeSecurity]
internal static class NanoServerStaShim
{
    private const uint PageExecuteReadWrite = 0x40;
    private const ushort Pe32Magic = 0x10B;
    private const ushort Pe32PlusMagic = 0x20B;
    private const int ImportDirectoryIndex = 1;
    private const int DelayImportDirectoryIndex = 13;
    private const int ErrorModNotFound = 126;
    private const int CachedPresent = 1;
    private const int CachedNotPresent = 2;

    private static readonly string[] ShimmedApiSetPrefixes =
    [
        "ext-ms-win-rtcore-ntuser-window-ext-l1-",
        "ext-ms-win-rtcore-ntuser-synch-ext-l1-",
        "ext-ms-win-ntuser-message-l1-",
        "ext-ms-win-ntuser-synch-l1-",
        "ext-ms-win-ntuser-window-l1-",
        "ext-ms-win-ntuser-windowclass-l1-",
        "ext-ms-win-ntuser-windowstation-l1-",
        "ext-ms-win-ntuser-private-l1-",
        "ext-ms-win-ntuser-misc-l1-",
        "ext-ms-win-com-sta-l1-",
    ];

    private static readonly object HookLock = new();
    private static readonly List<ModuleHookState> HookedModules = [];

    private static readonly ApiSetQueryApiSetPresenceDelegate HookedPresenceDelegate = OnApiSetQueryApiSetPresence;
    private static readonly ApiSetQueryApiSetPresenceExDelegate HookedPresenceExDelegate = OnApiSetQueryApiSetPresenceEx;
    private static readonly ResolveDelayLoadedApiDelegate HookedResolveDelayLoadDelegate = OnResolveDelayLoadedApi;
    private static readonly GetModuleHandleWDelegate HookedGetModuleHandleWDelegate = OnGetModuleHandleW;

    private static readonly IntPtr HookedPresencePtr = Marshal.GetFunctionPointerForDelegate(HookedPresenceDelegate);
    private static readonly IntPtr HookedPresenceExPtr = Marshal.GetFunctionPointerForDelegate(HookedPresenceExDelegate);
    private static readonly IntPtr HookedResolveDelayLoadPtr = Marshal.GetFunctionPointerForDelegate(HookedResolveDelayLoadDelegate);
    private static readonly IntPtr HookedGetModuleHandleWPtr = Marshal.GetFunctionPointerForDelegate(HookedGetModuleHandleWDelegate);

    private static ApiSetQueryApiSetPresenceDelegate? _originalPresence;
    private static ApiSetQueryApiSetPresenceExDelegate? _originalPresenceEx;
    private static ResolveDelayLoadedApiDelegate? _originalResolveDelayLoad;
    private static GetModuleHandleWDelegate? _originalGetModuleHandleW;
    private static OsStaShimActionDelegate? _osStaShimRefresh;
    private static OsStaShimActionDelegate? _osStaShimRestoreOriginal;
    private static bool _hasOsStaShimModule;
    private static IntPtr _combaseHandle;
    private static bool _isNanoServer;
    private static bool _processExitRegistered;
    private static Thread? _mainStaThread;
    private static ManualResetEventSlim? _mainStaReadyEvent;
    private static ManualResetEventSlim? _mainStaStopEvent;
    private static uint _mainStaThreadId;

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate void OsStaShimActionDelegate();

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int ApiSetQueryApiSetPresenceDelegate(IntPtr namespaceName, IntPtr present);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int ApiSetQueryApiSetPresenceExDelegate(IntPtr namespaceName, IntPtr isInSchema, IntPtr present);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr ResolveDelayLoadedApiDelegate(
        IntPtr parentModuleBase,
        IntPtr delayloadDescriptor,
        IntPtr failureDllHook,
        IntPtr failureSystemHook,
        IntPtr thunkAddress,
        uint flags);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate IntPtr GetModuleHandleWDelegate(IntPtr lpModuleName);

    internal static bool IsInstalled { get; private set; }

    internal static void Initialize()
    {
        if (Environment.OSVersion.Platform != PlatformID.Win32NT)
        {
            return;
        }

        lock (HookLock)
        {
            if (!_processExitRegistered)
            {
                AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
                _processExitRegistered = true;
            }

            var systemDir = Environment.SystemDirectory;
            var win32uPath = Path.Combine(systemDir, "win32u.dll");
            _isNanoServer = !File.Exists(win32uPath);
            if (_isNanoServer)
            {
                PreloadForwarders(systemDir);
            }

            RefreshStateCore();
        }
    }

    internal static void RefreshState()
    {
        if (Environment.OSVersion.Platform != PlatformID.Win32NT)
        {
            return;
        }

        lock (HookLock)
        {
            RefreshStateCore();
        }

        WaitForMainStaHostReady();
    }

    internal static void WaitForMainStaHostReady()
    {
        ManualResetEventSlim? readyEvent;
        Thread? hostThread;
        lock (HookLock)
        {
            readyEvent = _mainStaReadyEvent;
            hostThread = _mainStaThread;
        }

        if (readyEvent is not null && hostThread is { IsAlive: true } && Thread.CurrentThread != hostThread)
        {
            _ = readyEvent.Wait(5000);
        }
    }

    internal static bool IsShimmedApiSet(string? name)
    {
        if (name is not { Length: > 0 })
        {
            return false;
        }

        for (var i = 0; i < ShimmedApiSetPrefixes.Length; i++)
        {
            if (name.StartsWith(ShimmedApiSetPrefixes[i], StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static void PreloadForwarders(string systemDir)
    {
        var forwardersDir = Path.Combine(systemDir, "forwarders");
        if (Directory.Exists(forwardersDir))
        {
            var currentPath = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            if (currentPath.IndexOf(forwardersDir, StringComparison.OrdinalIgnoreCase) < 0)
            {
                var updatedPath = currentPath.Length > 0
                    ? currentPath + Path.PathSeparator + forwardersDir
                    : forwardersDir;
                Environment.SetEnvironmentVariable("PATH", updatedPath);
            }
        }

        _combaseHandle = EnsureModuleLoaded("combase.dll", forwardersDir);
        _ = EnsureModuleLoaded("user32.dll", forwardersDir);
        _ = EnsureModuleLoaded("ole32.dll", forwardersDir);
        _ = EnsureModuleLoaded("advapi32.dll", forwardersDir);
        _ = EnsureModuleLoaded("oleaut32.dll", forwardersDir);
    }

    private static IntPtr EnsureModuleLoaded(string dllName, string forwardersDir)
    {
        var handle = NativeMethods.LoadLibraryEx(dllName, IntPtr.Zero, 0);
        if (handle != IntPtr.Zero)
        {
            return handle;
        }

        var forwarderPath = Path.Combine(forwardersDir, dllName);
        if (File.Exists(forwarderPath))
        {
            handle = NativeMethods.LoadLibraryEx(forwarderPath, IntPtr.Zero, 0);
        }

        return handle;
    }

    private static void RefreshStateCore()
    {
        var forceShim = NanoServerOptions.ForceInProcessStaShim;
        if (!_isNanoServer && !forceShim)
        {
            if (IsInstalled)
            {
                IsInstalled = false;
                for (var i = 0; i < HookedModules.Count; i++)
                {
                    HookedModules[i].ApplyState(useInProcessShim: false);
                    HookedModules[i].SetKnownPresenceCaches(CachedPresent);
                }
            }

            return;
        }

        EnsureModulesDiscovered();
        var enableLibShim = NanoServerOptions.EnableStaShim;
        var osShimDisabled = NanoServerOptions.IsEnvironmentFlagTrue(NanoServerOptions.DisableOsStaShimEnvironmentVariable);
        var osStaAvailable = !osShimDisabled && (_hasOsStaShimModule || QueryOriginalWindowApiSetPresence());
        var useInProcessShim = enableLibShim && (!osStaAvailable || forceShim);
        var isStaActive = useInProcessShim || osStaAvailable;

        if (!isStaActive)
        {
            ApplyInactiveStaState();
            return;
        }

        if (!useInProcessShim)
        {
            ApplyOsStaShimState();
            return;
        }

        ApplyInProcessStaShimState();
    }

    private static void ApplyInactiveStaState()
    {
        IsInstalled = false;
        _osStaShimRestoreOriginal?.Invoke();
        for (var i = 0; i < HookedModules.Count; i++)
        {
            HookedModules[i].ApplyState(useInProcessShim: false);
            HookedModules[i].SetKnownPresenceCaches(CachedNotPresent);
        }
    }

    private static void ApplyOsStaShimState()
    {
        IsInstalled = false;
        for (var i = 0; i < HookedModules.Count; i++)
        {
            HookedModules[i].RestoreAllOriginalThunks();
        }

        if (_hasOsStaShimModule)
        {
            _osStaShimRefresh?.Invoke();
            return;
        }

        for (var i = 0; i < HookedModules.Count; i++)
        {
            HookedModules[i].ApplyState(useInProcessShim: false);
            HookedModules[i].SetKnownPresenceCaches(CachedPresent);
        }
    }

    private static void ApplyInProcessStaShimState()
    {
        _osStaShimRestoreOriginal?.Invoke();
        IsInstalled = true;
        for (var i = 0; i < HookedModules.Count; i++)
        {
            HookedModules[i].ApplyState(useInProcessShim: true);
            HookedModules[i].SetKnownPresenceCaches(CachedPresent);
        }

        if (_isNanoServer)
        {
            EnsureMainStaHostThreadStarted();
        }
    }

    private static void EnsureMainStaHostThreadStarted()
    {
        if (_mainStaThread is { IsAlive: true })
        {
            return;
        }

        var readyEvent = new ManualResetEventSlim(initialState: false);
        var stopEvent = new ManualResetEventSlim(initialState: false);
        var thread = new Thread(() => MainStaHostThreadProc(readyEvent, stopEvent))
        {
            IsBackground = true,
            Name = "NanoServerMainStaHost",
        };

        _mainStaReadyEvent = readyEvent;
        _mainStaStopEvent = stopEvent;
        _mainStaThread = thread;
        thread.Start();
    }

    private static void MainStaHostThreadProc(ManualResetEventSlim readyEvent, ManualResetEventSlim stopEvent)
    {
        _mainStaThreadId = NativeMethods.GetCurrentThreadId();
        NativeMethods.CoUninitialize();
        var hr = NativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_APARTMENTTHREADED);
        readyEvent.Set();
        if (hr < 0)
        {
            return;
        }

        try
        {
            NanoServerWindowManager.RunHostMessageLoop(stopEvent);
        }
        finally
        {
            NativeMethods.CoUninitialize();
        }
    }

    private static void StopMainStaHostThread()
    {
        var thread = _mainStaThread;
        var stopEvent = _mainStaStopEvent;
        var readyEvent = _mainStaReadyEvent;
        var threadId = _mainStaThreadId;
        _mainStaThread = null;
        _mainStaStopEvent = null;
        _mainStaReadyEvent = null;
        _mainStaThreadId = 0;

        stopEvent?.Set();
        if (threadId != 0)
        {
            NanoServerWindowManager.WakeThreadQueue(threadId);
        }

        if (thread is { IsAlive: true } && Thread.CurrentThread != thread)
        {
            _ = thread.Join(2000);
        }

        stopEvent?.Dispose();
        readyEvent?.Dispose();
    }

    private static void OnProcessExit(object? sender, EventArgs e)
    {
        _ = sender;
        _ = e;
        lock (HookLock)
        {
            StopMainStaHostThread();
            IsInstalled = false;
            for (var i = 0; i < HookedModules.Count; i++)
            {
                HookedModules[i].RestoreAllOriginalThunks();
            }
        }
    }

    private static void EnsureModulesDiscovered()
    {
        if (HookedModules.Count > 0)
        {
            return;
        }

        var systemDir = Environment.SystemDirectory;
        var forwardersDir = Path.Combine(systemDir, "forwarders");
        var osShimHandle = EnsureModuleLoaded("nanoserver_sta_shim.dll", forwardersDir);
        if (osShimHandle != IntPtr.Zero)
        {
            var refreshPtr = NativeMethods.GetProcAddress(osShimHandle, "NanoServerStaShim_Refresh");
            var restorePtr = NativeMethods.GetProcAddress(osShimHandle, "NanoServerStaShim_RestoreOriginal");
            if (refreshPtr != IntPtr.Zero && restorePtr != IntPtr.Zero)
            {
                _osStaShimRefresh = Marshal.GetDelegateForFunctionPointer<OsStaShimActionDelegate>(refreshPtr);
                _osStaShimRestoreOriginal = Marshal.GetDelegateForFunctionPointer<OsStaShimActionDelegate>(restorePtr);
                _hasOsStaShimModule = true;
                _osStaShimRestoreOriginal?.Invoke();
            }
        }

        string[] targetDlls = ["combase.dll", "user32.dll", "ole32.dll"];
        for (var i = 0; i < targetDlls.Length; i++)
        {
            var handle = EnsureModuleLoaded(targetDlls[i], forwardersDir);
            if (handle == IntPtr.Zero)
            {
                continue;
            }

            if (i == 0)
            {
                _combaseHandle = handle;
            }

            var state = ModuleHookState.TryCreate(handle);
            if (state is not null)
            {
                CaptureOriginalImports(state);
                HookedModules.Add(state);
            }
        }
    }

    private static void CaptureOriginalImports(ModuleHookState state)
    {
        if (_originalPresence is null && state.PresenceIatSlot != IntPtr.Zero)
        {
            var orig = Marshal.ReadIntPtr(state.PresenceIatSlot);
            if (orig != IntPtr.Zero && orig != HookedPresencePtr)
            {
                _originalPresence = Marshal.GetDelegateForFunctionPointer<ApiSetQueryApiSetPresenceDelegate>(orig);
            }
        }

        if (_originalPresenceEx is null && state.PresenceExIatSlot != IntPtr.Zero)
        {
            var orig = Marshal.ReadIntPtr(state.PresenceExIatSlot);
            if (orig != IntPtr.Zero && orig != HookedPresenceExPtr)
            {
                _originalPresenceEx = Marshal.GetDelegateForFunctionPointer<ApiSetQueryApiSetPresenceExDelegate>(orig);
            }
        }

        if (_originalResolveDelayLoad is null && state.ResolveDelayLoadIatSlot != IntPtr.Zero)
        {
            var orig = Marshal.ReadIntPtr(state.ResolveDelayLoadIatSlot);
            if (orig != IntPtr.Zero && orig != HookedResolveDelayLoadPtr)
            {
                _originalResolveDelayLoad = Marshal.GetDelegateForFunctionPointer<ResolveDelayLoadedApiDelegate>(orig);
            }
        }

        if (_originalGetModuleHandleW is null && state.GetModuleHandleWIatSlot != IntPtr.Zero)
        {
            var orig = Marshal.ReadIntPtr(state.GetModuleHandleWIatSlot);
            if (orig != IntPtr.Zero && orig != HookedGetModuleHandleWPtr)
            {
                _originalGetModuleHandleW = Marshal.GetDelegateForFunctionPointer<GetModuleHandleWDelegate>(orig);
            }
        }
    }

    private static bool QueryOriginalWindowApiSetPresence()
    {
        if (_originalPresence is null)
        {
            return false;
        }

        const string contract = "ext-ms-win-rtcore-ntuser-window-ext-l1-1-0";
        var strPtr = Marshal.StringToHGlobalUni(contract);
        var unicodeStr = Marshal.AllocHGlobal(IntPtr.Size * 2);
        var presentPtr = Marshal.AllocHGlobal(4);
        try
        {
            var byteLen = checked((short)(contract.Length * 2));
            Marshal.WriteInt16(unicodeStr, 0, byteLen);
            Marshal.WriteInt16(unicodeStr, 2, checked((short)(byteLen + 2)));
            Marshal.WriteIntPtr(unicodeStr, IntPtr.Size, strPtr);
            Marshal.WriteByte(presentPtr, 0);

            var status = _originalPresence(unicodeStr, presentPtr);
            return status >= 0 && Marshal.ReadByte(presentPtr) != 0;
        }
        finally
        {
            Marshal.FreeHGlobal(presentPtr);
            Marshal.FreeHGlobal(unicodeStr);
            Marshal.FreeHGlobal(strPtr);
        }
    }

    private static bool IsAnyStaProviderActive()
    {
        if (IsInstalled)
        {
            return true;
        }

        if (NanoServerOptions.IsEnvironmentFlagTrue(NanoServerOptions.DisableOsStaShimEnvironmentVariable))
        {
            return false;
        }

        return _hasOsStaShimModule || QueryOriginalWindowApiSetPresence();
    }

    private static int OnApiSetQueryApiSetPresence(IntPtr namespaceName, IntPtr present)
    {
        var name = ReadUnicodeString(namespaceName);
        if (IsShimmedApiSet(name))
        {
            if (present != IntPtr.Zero)
            {
                Marshal.WriteByte(present, IsAnyStaProviderActive() ? (byte)1 : (byte)0);
            }

            return 0;
        }

        if (_originalPresence is not null)
        {
            return _originalPresence(namespaceName, present);
        }

        if (present != IntPtr.Zero)
        {
            Marshal.WriteByte(present, 0);
        }

        return 0;
    }

    private static int OnApiSetQueryApiSetPresenceEx(IntPtr namespaceName, IntPtr isInSchema, IntPtr present)
    {
        var name = ReadUnicodeString(namespaceName);
        if (IsShimmedApiSet(name))
        {
            var active = IsAnyStaProviderActive() ? (byte)1 : (byte)0;
            if (isInSchema != IntPtr.Zero)
            {
                Marshal.WriteByte(isInSchema, active);
            }

            if (present != IntPtr.Zero)
            {
                Marshal.WriteByte(present, active);
            }

            return 0;
        }

        if (_originalPresenceEx is not null)
        {
            return _originalPresenceEx(namespaceName, isInSchema, present);
        }

        return 0;
    }

    private static IntPtr OnResolveDelayLoadedApi(
        IntPtr parentModuleBase,
        IntPtr delayloadDescriptor,
        IntPtr failureDllHook,
        IntPtr failureSystemHook,
        IntPtr thunkAddress,
        uint flags)
    {
        if (parentModuleBase != IntPtr.Zero && delayloadDescriptor != IntPtr.Zero && thunkAddress != IntPtr.Zero)
        {
            var rvaDllName = Marshal.ReadInt32(delayloadDescriptor, 4);
            var dllName = Marshal.PtrToStringAnsi(AddRva(parentModuleBase, rvaDllName));
            if (IsShimmedApiSet(dllName) && IsInstalled)
            {
                var resolved = TryResolveShimmedThunk(parentModuleBase, delayloadDescriptor, thunkAddress);
                if (resolved != IntPtr.Zero)
                {
                    return resolved;
                }
            }
        }

        return _originalResolveDelayLoad is not null
            ? _originalResolveDelayLoad(parentModuleBase, delayloadDescriptor, failureDllHook, failureSystemHook, thunkAddress, flags)
            : IntPtr.Zero;
    }

    private static IntPtr TryResolveShimmedThunk(IntPtr parentModuleBase, IntPtr delayloadDescriptor, IntPtr thunkAddress)
    {
        var rvaIat = Marshal.ReadInt32(delayloadDescriptor, 12);
        var rvaInt = Marshal.ReadInt32(delayloadDescriptor, 16);
        var iatBase = AddRva(parentModuleBase, rvaIat);
        var intBase = AddRva(parentModuleBase, rvaInt);
        var slotOffset = unchecked((int)(thunkAddress.ToInt64() - iatBase.ToInt64()));
        if (slotOffset < 0 || (slotOffset % IntPtr.Size) != 0)
        {
            return IntPtr.Zero;
        }

        var intEntry = Marshal.ReadIntPtr(intBase, slotOffset);
        if (intEntry == IntPtr.Zero || IsOrdinalImport(intEntry))
        {
            return IntPtr.Zero;
        }

        var nameRva = (int)(intEntry.ToInt64() & 0x7FFFFFFF);
        var funcName = Marshal.PtrToStringAnsi(AddRva(parentModuleBase, nameRva + 2));
        if (funcName is not { Length: > 0 })
        {
            return IntPtr.Zero;
        }

        var fnPtr = NanoServerWindowManager.ResolveFunction(funcName);
        if (fnPtr == IntPtr.Zero)
        {
            return IntPtr.Zero;
        }

        WritePointerWithProtection(thunkAddress, fnPtr);
        return fnPtr;
    }

    private static IntPtr OnGetModuleHandleW(IntPtr lpModuleName)
    {
        if (lpModuleName != IntPtr.Zero)
        {
            var name = Marshal.PtrToStringUni(lpModuleName);
            if (IsShimmedApiSet(name))
            {
                if (IsInstalled)
                {
                    return _combaseHandle != IntPtr.Zero ? _combaseHandle : new IntPtr(0x10000000);
                }

                if (!IsAnyStaProviderActive())
                {
                    NativeMethods.SetLastError(ErrorModNotFound);
                    return IntPtr.Zero;
                }
            }
        }

        return _originalGetModuleHandleW is not null
            ? _originalGetModuleHandleW(lpModuleName)
            : IntPtr.Zero;
    }

    private static string? ReadUnicodeString(IntPtr unicodeStringPtr)
    {
        if (unicodeStringPtr == IntPtr.Zero)
        {
            return null;
        }

        var byteLength = unchecked((ushort)Marshal.ReadInt16(unicodeStringPtr, 0));
        var buffer = Marshal.ReadIntPtr(unicodeStringPtr, IntPtr.Size);
        if (buffer == IntPtr.Zero || byteLength == 0)
        {
            return string.Empty;
        }

        return Marshal.PtrToStringUni(buffer, byteLength / 2);
    }

    private static IntPtr AddRva(IntPtr baseAddress, int rva)
        => new(baseAddress.ToInt64() + rva);

    private static bool IsOrdinalImport(IntPtr thunkValue)
        => IntPtr.Size == 8
            ? (thunkValue.ToInt64() & unchecked((long)0x8000000000000000UL)) != 0
            : (thunkValue.ToInt32() & unchecked((int)0x80000000u)) != 0;

    private static void WritePointerWithProtection(IntPtr address, IntPtr value)
    {
        if (NativeMethods.VirtualProtect(address, new UIntPtr((uint)IntPtr.Size), PageExecuteReadWrite, out var oldProtect))
        {
            Marshal.WriteIntPtr(address, value);
            _ = NativeMethods.VirtualProtect(address, new UIntPtr((uint)IntPtr.Size), oldProtect, out _);
        }
    }

    [StructLayout(LayoutKind.Auto)]
    private readonly struct DelayLoadSlot
    {
        internal DelayLoadSlot(IntPtr iatSlot, IntPtr initialThunk, IntPtr shimTarget)
        {
            IatSlot = iatSlot;
            InitialThunk = initialThunk;
            ShimTarget = shimTarget;
        }

        internal IntPtr IatSlot { get; }

        internal IntPtr InitialThunk { get; }

        internal IntPtr ShimTarget { get; }
    }

    private sealed class ModuleHookState
    {
        private const uint ImageScnMemExecute = 0x20000000;
        private readonly IntPtr _moduleBase;
        private readonly List<DelayLoadSlot> _shimmedSlots = [];
        private readonly List<IntPtr> _shimmedHmodAddresses = [];
        private readonly List<IntPtr> _knownPresenceCacheAddresses = [];
        private IntPtr _originalPresenceThunk;
        private IntPtr _originalPresenceExThunk;
        private IntPtr _originalResolveDelayLoadThunk;
        private IntPtr _originalGetModuleHandleWThunk;

        private ModuleHookState(IntPtr moduleBase)
        {
            _moduleBase = moduleBase;
        }

        internal IntPtr PresenceIatSlot { get; private set; }

        internal IntPtr PresenceExIatSlot { get; private set; }

        internal IntPtr ResolveDelayLoadIatSlot { get; private set; }

        internal IntPtr GetModuleHandleWIatSlot { get; private set; }

        internal static ModuleHookState? TryCreate(IntPtr moduleBase)
        {
            if (moduleBase == IntPtr.Zero || Marshal.ReadInt16(moduleBase, 0) != 0x5A4D)
            {
                return null;
            }

            var eLfanew = Marshal.ReadInt32(moduleBase, 0x3C);
            var ntHeaders = AddRva(moduleBase, eLfanew);
            if (Marshal.ReadInt32(ntHeaders, 0) != 0x00004550)
            {
                return null;
            }

            var numberOfSections = unchecked((ushort)Marshal.ReadInt16(ntHeaders, 6));
            var sizeOfOptionalHeader = unchecked((ushort)Marshal.ReadInt16(ntHeaders, 20));
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
                return null;
            }

            var state = new ModuleHookState(moduleBase);
            var importRva = Marshal.ReadInt32(optionalHeader, dataDirOffset + (ImportDirectoryIndex * 8));
            if (importRva > 0)
            {
                state.ParseStandardImports(importRva);
            }

            var delayImportRva = Marshal.ReadInt32(optionalHeader, dataDirOffset + (DelayImportDirectoryIndex * 8));
            if (delayImportRva > 0)
            {
                state.ParseDelayImports(delayImportRva);
                state.DiscoverPresenceCaches(AddRva(optionalHeader, sizeOfOptionalHeader), numberOfSections);
            }

            return state;
        }

        internal void ApplyState(bool useInProcessShim)
        {
            InstallStandardHooks();
            for (var i = 0; i < _shimmedSlots.Count; i++)
            {
                var slot = _shimmedSlots[i];
                var target = useInProcessShim && slot.ShimTarget != IntPtr.Zero
                    ? slot.ShimTarget
                    : slot.InitialThunk;
                WritePointerWithProtection(slot.IatSlot, target);
            }

            for (var i = 0; i < _shimmedHmodAddresses.Count; i++)
            {
                WritePointerWithProtection(_shimmedHmodAddresses[i], IntPtr.Zero);
            }
        }

        internal void RestoreAllOriginalThunks()
        {
            if (PresenceIatSlot != IntPtr.Zero && _originalPresenceThunk != IntPtr.Zero)
            {
                WritePointerWithProtection(PresenceIatSlot, _originalPresenceThunk);
            }

            if (PresenceExIatSlot != IntPtr.Zero && _originalPresenceExThunk != IntPtr.Zero)
            {
                WritePointerWithProtection(PresenceExIatSlot, _originalPresenceExThunk);
            }

            if (ResolveDelayLoadIatSlot != IntPtr.Zero && _originalResolveDelayLoadThunk != IntPtr.Zero)
            {
                WritePointerWithProtection(ResolveDelayLoadIatSlot, _originalResolveDelayLoadThunk);
            }

            if (GetModuleHandleWIatSlot != IntPtr.Zero && _originalGetModuleHandleWThunk != IntPtr.Zero)
            {
                WritePointerWithProtection(GetModuleHandleWIatSlot, _originalGetModuleHandleWThunk);
            }

            for (var i = 0; i < _shimmedSlots.Count; i++)
            {
                var slot = _shimmedSlots[i];
                WritePointerWithProtection(slot.IatSlot, slot.InitialThunk);
            }

            for (var i = 0; i < _shimmedHmodAddresses.Count; i++)
            {
                WritePointerWithProtection(_shimmedHmodAddresses[i], IntPtr.Zero);
            }
        }

        internal void SetKnownPresenceCaches(int cacheState)
        {
            for (var i = 0; i < _knownPresenceCacheAddresses.Count; i++)
            {
                WriteInt32WithProtection(_knownPresenceCacheAddresses[i], cacheState);
            }
        }

        private static void WriteInt32WithProtection(IntPtr address, int value)
        {
            if (NativeMethods.VirtualProtect(address, new UIntPtr(4), PageExecuteReadWrite, out var oldProtect))
            {
                Marshal.WriteInt32(address, value);
                _ = NativeMethods.VirtualProtect(address, new UIntPtr(4), oldProtect, out _);
            }
        }

        private void InstallStandardHooks()
        {
            if (PresenceIatSlot != IntPtr.Zero)
            {
                WritePointerWithProtection(PresenceIatSlot, HookedPresencePtr);
            }

            if (PresenceExIatSlot != IntPtr.Zero)
            {
                WritePointerWithProtection(PresenceExIatSlot, HookedPresenceExPtr);
            }

            _ = ResolveDelayLoadIatSlot;
            _ = GetModuleHandleWIatSlot;
        }

        private void DiscoverPresenceCaches(IntPtr firstSectionHeader, int numberOfSections)
        {
            for (var i = 0; i < numberOfSections; i++)
            {
                var section = AddRva(firstSectionHeader, i * 40);
                var characteristics = unchecked((uint)Marshal.ReadInt32(section, 36));
                if ((characteristics & ImageScnMemExecute) == 0)
                {
                    continue;
                }

                var virtualSize = Marshal.ReadInt32(section, 8);
                var virtualAddress = Marshal.ReadInt32(section, 12);
                var rawSize = Marshal.ReadInt32(section, 16);
                var span = Math.Min(virtualSize, rawSize);
                if (span < 64)
                {
                    continue;
                }

                if (IntPtr.Size == 4)
                {
                    ScanExecutableSectionX86(virtualAddress, span);
                }
                else
                {
                    ScanExecutableSectionX64(virtualAddress, span);
                }
            }
        }

        private void ScanExecutableSectionX86(int startRva, int span)
        {
            var endRva = startRva + span - 12;
            for (var rva = startRva + 44; rva <= endRva; rva++)
            {
                var addr = AddRva(_moduleBase, rva);
                if (Marshal.ReadInt32(addr) != 0x41C1940F
                    || unchecked((ushort)Marshal.ReadInt16(addr, 4)) != 0x0D89
                    || Marshal.ReadByte(addr, -44) != 0xA1
                    || Marshal.ReadByte(addr, -21) != 0x68)
                {
                    continue;
                }

                var cacheVa1 = Marshal.ReadInt32(addr, -43);
                var cacheVa2 = Marshal.ReadInt32(addr, 6);
                if (cacheVa1 != cacheVa2)
                {
                    continue;
                }

                var ustrAddr = new IntPtr(Marshal.ReadInt32(addr, -20));
                if (IsShimmedApiSet(ReadUnicodeString(ustrAddr)))
                {
                    var cacheAddr = new IntPtr(cacheVa2);
                    if (!_knownPresenceCacheAddresses.Contains(cacheAddr))
                    {
                        _knownPresenceCacheAddresses.Add(cacheAddr);
                    }
                }
            }
        }

        private void ScanExecutableSectionX64(int startRva, int span)
        {
            var endRva = startRva + span - 16;
            for (var rva = startRva + 48; rva <= endRva; rva++)
            {
                var addr = AddRva(_moduleBase, rva);
                if (Marshal.ReadInt32(addr) != unchecked((int)0xD21BD9F6)
                    || Marshal.ReadInt32(addr, 4) != unchecked((int)0x8902C283)
                    || Marshal.ReadByte(addr, 8) != 0x15
                    || Marshal.ReadByte(addr, -48) != 0x8B
                    || Marshal.ReadByte(addr, -47) != 0x0D
                    || Marshal.ReadByte(addr, -22) != 0x48
                    || Marshal.ReadByte(addr, -21) != 0x8D
                    || Marshal.ReadByte(addr, -20) != 0x0D)
                {
                    continue;
                }

                var disp1 = Marshal.ReadInt32(addr, -46);
                var disp2 = Marshal.ReadInt32(addr, 9);
                var cacheAddr1 = AddRva(addr, -42 + disp1);
                var cacheAddr2 = AddRva(addr, 13 + disp2);
                if (cacheAddr1 != cacheAddr2)
                {
                    continue;
                }

                var ustrDisp = Marshal.ReadInt32(addr, -19);
                var ustrAddr = AddRva(addr, -15 + ustrDisp);
                if (IsShimmedApiSet(ReadUnicodeString(ustrAddr)) && !_knownPresenceCacheAddresses.Contains(cacheAddr2))
                {
                    _knownPresenceCacheAddresses.Add(cacheAddr2);
                }
            }
        }

        private void ParseStandardImports(int importRva)
        {
            for (var descOffset = 0; descOffset < 0x10000; descOffset += 20)
            {
                var descAddr = AddRva(_moduleBase, importRva + descOffset);
                var originalFirstThunk = Marshal.ReadInt32(descAddr, 0);
                var nameRva = Marshal.ReadInt32(descAddr, 12);
                var firstThunk = Marshal.ReadInt32(descAddr, 16);
                if (originalFirstThunk == 0 && nameRva == 0 && firstThunk == 0)
                {
                    break;
                }

                var thunkRva = originalFirstThunk != 0 ? originalFirstThunk : firstThunk;
                ScanStandardThunkChain(thunkRva, firstThunk);
            }
        }

        private void ScanStandardThunkChain(int intRva, int iatRva)
        {
            for (var offset = 0; offset < 0x10000; offset += IntPtr.Size)
            {
                var thunkData = Marshal.ReadIntPtr(_moduleBase, intRva + offset);
                if (thunkData == IntPtr.Zero)
                {
                    break;
                }

                if (IsOrdinalImport(thunkData))
                {
                    continue;
                }

                var nameRva = (int)(thunkData.ToInt64() & 0x7FFFFFFF);
                var funcName = Marshal.PtrToStringAnsi(AddRva(_moduleBase, nameRva + 2));
                var iatSlot = AddRva(_moduleBase, iatRva + offset);
                switch (funcName)
                {
                    case "ApiSetQueryApiSetPresence":
                        PresenceIatSlot = iatSlot;
                        _originalPresenceThunk = Marshal.ReadIntPtr(iatSlot);
                        break;
                    case "ApiSetQueryApiSetPresenceEx":
                        PresenceExIatSlot = iatSlot;
                        _originalPresenceExThunk = Marshal.ReadIntPtr(iatSlot);
                        break;
                    case "ResolveDelayLoadedAPI":
                        ResolveDelayLoadIatSlot = iatSlot;
                        _originalResolveDelayLoadThunk = Marshal.ReadIntPtr(iatSlot);
                        break;
                    case "GetModuleHandleW":
                        GetModuleHandleWIatSlot = iatSlot;
                        _originalGetModuleHandleWThunk = Marshal.ReadIntPtr(iatSlot);
                        break;
                    default:
                        break;
                }
            }
        }

        private void ParseDelayImports(int delayImportRva)
        {
            for (var descOffset = 0; descOffset < 0x10000; descOffset += 32)
            {
                var descAddr = AddRva(_moduleBase, delayImportRva + descOffset);
                var rvaDllName = Marshal.ReadInt32(descAddr, 4);
                var rvaHmod = Marshal.ReadInt32(descAddr, 8);
                var rvaIat = Marshal.ReadInt32(descAddr, 12);
                var rvaInt = Marshal.ReadInt32(descAddr, 16);
                if (rvaDllName == 0)
                {
                    break;
                }

                var dllName = Marshal.PtrToStringAnsi(AddRva(_moduleBase, rvaDllName));
                if (IsShimmedApiSet(dllName))
                {
                    if (rvaHmod > 0)
                    {
                        _shimmedHmodAddresses.Add(AddRva(_moduleBase, rvaHmod));
                    }

                    ScanDelayThunkChain(rvaInt, rvaIat);
                }
            }
        }

        private void ScanDelayThunkChain(int rvaInt, int rvaIat)
        {
            for (var offset = 0; offset < 0x10000; offset += IntPtr.Size)
            {
                var intEntry = Marshal.ReadIntPtr(_moduleBase, rvaInt + offset);
                if (intEntry == IntPtr.Zero)
                {
                    break;
                }

                if (IsOrdinalImport(intEntry))
                {
                    continue;
                }

                var nameRva = (int)(intEntry.ToInt64() & 0x7FFFFFFF);
                var funcName = Marshal.PtrToStringAnsi(AddRva(_moduleBase, nameRva + 2));
                if (funcName is not { Length: > 0 })
                {
                    continue;
                }

                var iatSlot = AddRva(_moduleBase, rvaIat + offset);
                var initialThunk = Marshal.ReadIntPtr(iatSlot);
                var shimTarget = NanoServerWindowManager.ResolveFunction(funcName);
                _shimmedSlots.Add(new DelayLoadSlot(iatSlot, initialThunk, shimTarget));
            }
        }
    }
}
#pragma warning restore SCS0018, SEC0116
#pragma warning restore S101
#pragma warning restore SA1121
#pragma warning restore MA0158
#pragma warning restore IDISP007
#pragma warning restore IDISP006
#pragma warning restore IDISP003
#pragma warning restore IDISP001
