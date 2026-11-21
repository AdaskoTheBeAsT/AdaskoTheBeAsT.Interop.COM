using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable MA0158 // Use System.Threading.Lock
#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable S101
#pragma warning disable SCS0018, SEC0116 // System directory check uses OS-controlled path.
[SuppressUnmanagedCodeSecurity]
internal static class NanoServerDispenserManager
{
#pragma warning disable MA0176 // Optimize guid creation
    // Verified against Windows SDK um\ComSvcs.h:
    // CLSID_DispenserManager = {ecabb0c0-7f19-11d2-978e-0000f8757e2a}
    // IID_IDispenserManager  = {5cb31e10-2b5f-11cf-be10-00aa00a2fa25}
    // IID_IHolder            = {bf6a1850-2b45-11cf-be10-00aa00a2fa25}
    internal static readonly Guid ClsidDispenserManager = new("ecabb0c0-7f19-11d2-978e-0000f8757e2a");
    internal static readonly Guid IidIDispenserManager = new("5cb31e10-2b5f-11cf-be10-00aa00a2fa25");
    internal static readonly Guid IidIHolder = new("bf6a1850-2b45-11cf-be10-00aa00a2fa25");
#pragma warning restore MA0176

    private const uint PageSize = 4096;
    private const int HolderPoolCapacity = 64;

    private const int OffsetIidIUnknown = 0x000;
    private const int OffsetIidIClassFactory = 0x010;
    private const int OffsetIidIDispenserManager = 0x020;
    private const int OffsetIidIHolder = 0x030;
    private const int OffsetClassFactoryObj = 0x040;
    private const int OffsetDispenserManagerObj = 0x048;
    private const int OffsetClassFactoryVtbl = 0x050;
    private const int OffsetDispenserManagerVtbl = 0x078;
    private const int OffsetHolderVtbl = 0x0A0;
    private const int OffsetCodeStart = 0x100;

#pragma warning disable MA0176 // Optimize guid creation
    private static readonly Guid IidIUnknown = new("00000000-0000-0000-C000-000000000046");
    private static readonly Guid IidIClassFactory = new("00000001-0000-0000-C000-000000000046");
#pragma warning restore MA0176

#if NET10_0_OR_GREATER
    private static readonly Lock DispenserLock = new();
#endif
#if NET472_OR_GREATER
    private static readonly object DispenserLock = new();
#endif

    // Intentionally static-lifetime native pages (never freed before ExitProcess) so background
    // native OLE DB / Jet threads detaching during DLL_THREAD_DETACH after CLR shutdown never
    // re-enter managed code or touch unmapped pages.
    private static IntPtr _codePage;
    private static IntPtr _dataPage;
    private static IntPtr _classFactoryObjectPtr;
    private static IntPtr _mtaUsageCookie;
    private static int _activeProvider;
    private static int _isFallbackRegistered;

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int OsShimEnsureDispenserDelegate();

    internal enum DispenserProviderKind
    {
        None = 0,
        NanoOsShim = 1,
        NativeOs = 2,
        InProcessWxNative = 3,
    }

    internal static DispenserProviderKind ActiveProvider
        => (DispenserProviderKind)Volatile.Read(ref _activeProvider);

    internal static bool IsFallbackRegistered
        => Volatile.Read(ref _isFallbackRegistered) != 0;

    internal static IntPtr CodePageAddress => Volatile.Read(ref _codePage);

    internal static IntPtr DataPageAddress => Volatile.Read(ref _dataPage);

    internal static void RefreshState()
    {
        if (Environment.OSVersion.Platform != PlatformID.Win32NT)
        {
            return;
        }

        lock (DispenserLock)
        {
            if (!NanoServerOptions.ReadEnableDispenserManagerFallbackRaw())
            {
                if (!IsFallbackRegistered)
                {
                    Volatile.Write(ref _activeProvider, (int)DispenserProviderKind.None);
                }

                return;
            }

            if (IsFallbackRegistered)
            {
                Volatile.Write(ref _activeProvider, (int)DispenserProviderKind.InProcessWxNative);
                return;
            }

            var nativeProvider = QueryNativeProviderCapability();
            if (nativeProvider != DispenserProviderKind.None)
            {
                Volatile.Write(ref _activeProvider, (int)nativeProvider);
                return;
            }

            var factoryPtr = EnsureNativeStubAllocatedUnderLock();
            if (factoryPtr == IntPtr.Zero)
            {
                return;
            }

            if (_mtaUsageCookie == IntPtr.Zero)
            {
                _ = NativeMethods.CoIncrementMTAUsage(out _mtaUsageCookie);
            }

            var clsid = ClsidDispenserManager;
            var hr = NativeMethods.CoRegisterClassObject(
                ref clsid,
                factoryPtr,
                NativeMethods.CLSCTX_INPROC_SERVER,
                NativeMethods.REGCLS_MULTIPLEUSE,
                out _);
            if (hr >= 0)
            {
                Volatile.Write(ref _isFallbackRegistered, 1);
                Volatile.Write(ref _activeProvider, (int)DispenserProviderKind.InProcessWxNative);
            }
        }
    }

    internal static IntPtr GetOrCreateNativeClassFactoryForTesting()
    {
        lock (DispenserLock)
        {
            return EnsureNativeStubAllocatedUnderLock();
        }
    }

    internal static DispenserProviderKind QueryNativeProviderCapability()
    {
        // 1. Prefer nanoserver_sta_shim.dll native provider when present and capable.
        var osShimHandle = NanoServerStaShim.OsStaShimModuleHandle;
        if (osShimHandle == IntPtr.Zero)
        {
            osShimHandle = NativeMethods.GetModuleHandleW("nanoserver_sta_shim.dll");
        }

        if (osShimHandle != IntPtr.Zero && HasOsShimDispenserProvider(osShimHandle))
        {
            return DispenserProviderKind.NanoOsShim;
        }

        // 2. Check whether full Windows comsvcs.dll or an already-registered provider satisfies CLSID_DispenserManager.
        var comsvcsPath = Path.Combine(Environment.SystemDirectory, "comsvcs.dll");
        if (File.Exists(comsvcsPath))
        {
            return DispenserProviderKind.NativeOs;
        }

        var probeClsid = ClsidDispenserManager;
        var probeIid = IidIClassFactory;
        var hr = NativeMethods.CoGetClassObject(
            ref probeClsid,
            NativeMethods.CLSCTX_INPROC_SERVER,
            IntPtr.Zero,
            ref probeIid,
            out var existingFactory);
        if (hr >= 0 && existingFactory != IntPtr.Zero)
        {
            _ = Marshal.Release(existingFactory);
            return osShimHandle != IntPtr.Zero ? DispenserProviderKind.NanoOsShim : DispenserProviderKind.NativeOs;
        }

        return DispenserProviderKind.None;
    }

    private static bool HasOsShimDispenserProvider(IntPtr osShimHandle)
    {
        var ensureExport = NativeMethods.GetProcAddress(osShimHandle, "NanoServerStaShim_EnsureDispenserManager");
        if (ensureExport != IntPtr.Zero)
        {
            var ensureFn = Marshal.GetDelegateForFunctionPointer<OsShimEnsureDispenserDelegate>(ensureExport);
            if (ensureFn() >= 0)
            {
                return true;
            }
        }

        var dllGetClassObject = NativeMethods.GetProcAddress(osShimHandle, "DllGetClassObject");
        if (dllGetClassObject == IntPtr.Zero)
        {
            return false;
        }

        var fn = Marshal.GetDelegateForFunctionPointer<NativeMethods.DllGetClassObjectDelegate>(dllGetClassObject);
        var clsid = ClsidDispenserManager;
        var iid = IidIClassFactory;
        if (fn(ref clsid, ref iid, out var shimFactory) < 0 || shimFactory == IntPtr.Zero)
        {
            return false;
        }

        _ = Marshal.Release(shimFactory);
        return true;
    }

    private static IntPtr EnsureNativeStubAllocatedUnderLock()
    {
        if (_classFactoryObjectPtr != IntPtr.Zero)
        {
            return _classFactoryObjectPtr;
        }

        // Allocate separate data (PAGE_READWRITE forever) and code (PAGE_READWRITE -> PAGE_EXECUTE_READ) pages.
        var dataPage = NativeMethods.VirtualAlloc(
            IntPtr.Zero,
            new UIntPtr(PageSize),
            NativeMethods.MEM_COMMIT | NativeMethods.MEM_RESERVE,
            NativeMethods.PAGE_READWRITE);
        if (dataPage == IntPtr.Zero)
        {
            return IntPtr.Zero;
        }

        var codePage = NativeMethods.VirtualAlloc(
            IntPtr.Zero,
            new UIntPtr(PageSize),
            NativeMethods.MEM_COMMIT | NativeMethods.MEM_RESERVE,
            NativeMethods.PAGE_READWRITE);
        if (codePage == IntPtr.Zero)
        {
            return IntPtr.Zero;
        }

        PopulateNativeStubPages(codePage, dataPage);

        _ = NativeMethods.FlushInstructionCache(NativeMethods.GetCurrentProcess(), codePage, new UIntPtr(PageSize));

        // Transition code page from PAGE_READWRITE (0x04) to PAGE_EXECUTE_READ (0x20) to uphold W^X.
        if (!NativeMethods.VirtualProtect(codePage, new UIntPtr(PageSize), NativeMethods.PAGE_EXECUTE_READ, out _))
        {
            return IntPtr.Zero;
        }

        _dataPage = dataPage;
        _codePage = codePage;
        _classFactoryObjectPtr = OffsetAddress(codePage, OffsetClassFactoryObj);
        return _classFactoryObjectPtr;
    }

    private static void PopulateNativeStubPages(IntPtr codePage, IntPtr dataPage)
    {
        WriteGuid(OffsetAddress(codePage, OffsetIidIUnknown), IidIUnknown);
        WriteGuid(OffsetAddress(codePage, OffsetIidIClassFactory), IidIClassFactory);
        WriteGuid(OffsetAddress(codePage, OffsetIidIDispenserManager), IidIDispenserManager);
        WriteGuid(OffsetAddress(codePage, OffsetIidIHolder), IidIHolder);

        var cfVtblAddr = OffsetAddress(codePage, OffsetClassFactoryVtbl);
        var dmVtblAddr = OffsetAddress(codePage, OffsetDispenserManagerVtbl);
        var holderVtblAddr = OffsetAddress(codePage, OffsetHolderVtbl);
        var dmObjAddr = OffsetAddress(codePage, OffsetDispenserManagerObj);

        Marshal.WriteIntPtr(OffsetAddress(codePage, OffsetClassFactoryObj), cfVtblAddr);
        Marshal.WriteIntPtr(dmObjAddr, dmVtblAddr);

        var emitter = new CodeEmitter(codePage, OffsetCodeStart);
        var methods = IntPtr.Size == 8
            ? EmitAllMethodsX64(
                emitter,
                OffsetAddress(codePage, OffsetIidIUnknown),
                OffsetAddress(codePage, OffsetIidIClassFactory),
                OffsetAddress(codePage, OffsetIidIDispenserManager),
                OffsetAddress(codePage, OffsetIidIHolder),
                dmObjAddr,
                dataPage,
                holderVtblAddr)
            : EmitAllMethodsX86(
                emitter,
                OffsetAddress(codePage, OffsetIidIUnknown),
                OffsetAddress(codePage, OffsetIidIClassFactory),
                OffsetAddress(codePage, OffsetIidIDispenserManager),
                OffsetAddress(codePage, OffsetIidIHolder),
                dmObjAddr,
                dataPage,
                holderVtblAddr);

        PopulateAllVtables(cfVtblAddr, dmVtblAddr, holderVtblAddr, methods);
    }

    private static void PopulateAllVtables(
        IntPtr cfVtblAddr,
        IntPtr dmVtblAddr,
        IntPtr holderVtblAddr,
        StubMethodAddresses methods)
    {
        WriteVtableSlot(cfVtblAddr, 0, methods.ClassFactoryQi);
        WriteVtableSlot(cfVtblAddr, 1, methods.StaticAddRef);
        WriteVtableSlot(cfVtblAddr, 2, methods.StaticRelease);
        WriteVtableSlot(cfVtblAddr, 3, methods.ClassFactoryCreateInstance);
        WriteVtableSlot(cfVtblAddr, 4, methods.LockServer);

        WriteVtableSlot(dmVtblAddr, 0, methods.DispenserManagerQi);
        WriteVtableSlot(dmVtblAddr, 1, methods.StaticAddRef);
        WriteVtableSlot(dmVtblAddr, 2, methods.StaticRelease);
        WriteVtableSlot(dmVtblAddr, 3, methods.RegisterDispenser);
        WriteVtableSlot(dmVtblAddr, 4, methods.GetContext);

        WriteVtableSlot(holderVtblAddr, 0, methods.HolderQi);
        WriteVtableSlot(holderVtblAddr, 1, methods.HolderAddRef);
        WriteVtableSlot(holderVtblAddr, 2, methods.HolderRelease);
        WriteVtableSlot(holderVtblAddr, 3, methods.HolderAllocResource);
        WriteVtableSlot(holderVtblAddr, 4, methods.HolderFreeResource);
        WriteVtableSlot(holderVtblAddr, 5, methods.HolderTrackResource);
        WriteVtableSlot(holderVtblAddr, 6, methods.HolderTrackResource);
        WriteVtableSlot(holderVtblAddr, 7, methods.HolderUntrackResource);
        WriteVtableSlot(holderVtblAddr, 8, methods.HolderUntrackResourceS);
        WriteVtableSlot(holderVtblAddr, 9, methods.HolderClose);
        WriteVtableSlot(holderVtblAddr, 10, methods.HolderFreeResource);
    }

#pragma warning disable S107
    private static StubMethodAddresses EmitAllMethodsX64(
        CodeEmitter e,
        IntPtr iidUnk,
        IntPtr iidCf,
        IntPtr iidDm,
        IntPtr iidHolder,
        IntPtr dmObj,
        IntPtr dataPage,
        IntPtr holderVtbl)
#pragma warning restore S107
    {
        var staticAddRef = EmitStaticConstReturnX64(e, 2);
        var staticRelease = EmitStaticConstReturnX64(e, 1);
        var zeroRet = EmitZeroReturnX64(e);
        var cfQi = EmitQueryInterfaceX64(e, iidUnk, iidCf, addRefHolderSlot: false);
        var dmQiOffset = e.Align16();
        var dmQi = EmitQueryInterfaceX64(e, iidUnk, iidDm, addRefHolderSlot: false);
        var cfCreate = EmitCreateInstanceX64(e, dmObj, dmQiOffset);
        var getContext = EmitGetContextX64(e);
        var regDisp = EmitRegisterDispenserX64(e, dataPage, holderVtbl);
        var holderQi = EmitQueryInterfaceX64(e, iidUnk, iidHolder, addRefHolderSlot: true);
        var holderAddRef = EmitHolderAddRefX64(e);
        var holderRelease = EmitHolderReleaseX64(e);
        var holderClose = EmitHolderCloseX64(e);
        var holderAlloc = EmitHolderAllocResourceX64(e);
        var holderFree = EmitHolderDestroyForwardX64(e, driverVtableDisp: 0x38, conditionalOnR8: false);
        var holderUntrack = EmitHolderDestroyForwardX64(e, driverVtableDisp: 0x38, conditionalOnR8: true);
        var holderUntrackS = EmitHolderDestroyForwardX64(e, driverVtableDisp: 0x40, conditionalOnR8: true);

        return new StubMethodAddresses
        {
            StaticAddRef = staticAddRef,
            StaticRelease = staticRelease,
            LockServer = zeroRet,
            ClassFactoryQi = cfQi,
            ClassFactoryCreateInstance = cfCreate,
            DispenserManagerQi = dmQi,
            RegisterDispenser = regDisp,
            GetContext = getContext,
            HolderQi = holderQi,
            HolderAddRef = holderAddRef,
            HolderRelease = holderRelease,
            HolderClose = holderClose,
            HolderAllocResource = holderAlloc,
            HolderFreeResource = holderFree,
            HolderTrackResource = zeroRet,
            HolderUntrackResource = holderUntrack,
            HolderUntrackResourceS = holderUntrackS,
        };
    }

    private static IntPtr EmitStaticConstReturnX64(CodeEmitter e, int value)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0xB8);
        e.EmitInt32(value);
        e.Emit(0xC3);
        return start;
    }

    private static IntPtr EmitZeroReturnX64(CodeEmitter e)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0x31, 0xC0, 0xC3);
        return start;
    }

    private static IntPtr EmitQueryInterfaceX64(CodeEmitter e, IntPtr iidUnkAddr, IntPtr iidTargetAddr, bool addRefHolderSlot)
    {
        e.Align16();
        var start = e.CurrentAddress;

        // Return E_POINTER if the output pointer in r8 is null.
        e.Emit(0x4D, 0x85, 0xC0);
        var jnzPpvOk = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_POINTER);
        e.Emit(0xC3);

        e.MarkRel8(jnzPpvOk);
        e.Emit(0x49, 0xC7, 0x00, 0x00, 0x00, 0x00, 0x00);

        // Return E_INVALIDARG if the interface ID pointer in rdx is null.
        e.Emit(0x48, 0x85, 0xD2);
        var jnzRiidOk = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_INVALIDARG);
        e.Emit(0xC3);

        e.MarkRel8(jnzRiidOk);
        e.Emit(0x48, 0x8B, 0x02);
        e.Emit(0x4C, 0x8B, 0x4A, 0x08);

        var jeMatchedFromUnk = EmitGuidMatchBranchX64(e, iidUnkAddr);
        var jeMatchedFromTarget = EmitGuidMatchBranchX64(e, iidTargetAddr);

        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_NOINTERFACE);
        e.Emit(0xC3);

        e.MarkRel8(jeMatchedFromUnk);
        e.MarkRel8(jeMatchedFromTarget);
        e.Emit(0x49, 0x89, 0x08);
        if (addRefHolderSlot)
        {
            e.Emit(0xF0, 0xFF, 0x41, 0x08);
        }

        e.Emit(0x31, 0xC0, 0xC3);
        return start;
    }

    private static int EmitGuidMatchBranchX64(CodeEmitter e, IntPtr targetGuidAddr)
    {
        e.Emit(0x49, 0xBA);
        e.EmitInt64(targetGuidAddr.ToInt64());
        e.Emit(0x49, 0x3B, 0x02);
        var jneMiss = e.EmitJccRel8(0x75);
        e.Emit(0x4D, 0x3B, 0x4A, 0x08);
        var jeHit = e.EmitJccRel8(0x74);
        e.MarkRel8(jneMiss);
        return jeHit;
    }

    private static IntPtr EmitCreateInstanceX64(CodeEmitter e, IntPtr dmObjAddr, int dmQiOffset)
    {
        e.Align16();
        var start = e.CurrentAddress;

        // Return E_POINTER if the output pointer in r9 is null.
        e.Emit(0x4D, 0x85, 0xC9);
        var jnzPpvOk = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_POINTER);
        e.Emit(0xC3);

        e.MarkRel8(jnzPpvOk);
        e.Emit(0x49, 0xC7, 0x01, 0x00, 0x00, 0x00, 0x00);

        // Return CLASS_E_NOAGGREGATION if the caller supplies an outer object in rdx.
        e.Emit(0x48, 0x85, 0xD2);
        var jzNoAgg = e.EmitJccRel8(0x74);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.CLASS_E_NOAGGREGATION);
        e.Emit(0xC3);

        e.MarkRel8(jzNoAgg);

        // mov rcx, dmObjAddr; mov rdx, r8; mov r8, r9; jmp dmQi
        e.Emit(0x48, 0xB9);
        e.EmitInt64(dmObjAddr.ToInt64());
        e.Emit(0x4C, 0x89, 0xC2);
        e.Emit(0x4D, 0x89, 0xC8);
        e.EmitRel32JumpOrCall(0xE9, dmQiOffset);
        return start;
    }

    private static IntPtr EmitGetContextX64(CodeEmitter e)
    {
        e.Align16();
        var start = e.CurrentAddress;

        e.Emit(0x48, 0x85, 0xD2);
        var jzSkip1 = e.EmitJccRel8(0x74);
        e.Emit(0x48, 0xC7, 0x02, 0x00, 0x00, 0x00, 0x00);
        e.MarkRel8(jzSkip1);

        e.Emit(0x4D, 0x85, 0xC0);
        var jzSkip2 = e.EmitJccRel8(0x74);
        e.Emit(0x49, 0xC7, 0x00, 0x00, 0x00, 0x00, 0x00);
        e.MarkRel8(jzSkip2);

        e.Emit(0x31, 0xC0, 0xC3);
        return start;
    }

    private static IntPtr EmitRegisterDispenserX64(CodeEmitter e, IntPtr dataPageAddr, IntPtr holderVtblAddr)
    {
        e.Align16();
        var start = e.CurrentAddress;

        // Return E_POINTER if the output pointer in r9 is null.
        e.Emit(0x4D, 0x85, 0xC9);
        var jnzOk = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_POINTER);
        e.Emit(0xC3);

        e.MarkRel8(jnzOk);
        e.Emit(0x49, 0xC7, 0x01, 0x00, 0x00, 0x00, 0x00);

        EmitClaimHolderSlotX64(e, dataPageAddr);
        EmitInitHolderAndAddRefDriverX64(e, holderVtblAddr);
        return start;
    }

    private static void EmitClaimHolderSlotX64(CodeEmitter e, IntPtr dataPageAddr)
    {
        e.Emit(0x49, 0xBA);
        e.EmitInt64(dataPageAddr.ToInt64());
        e.Emit(0x41, 0xBB);
        e.EmitInt32(HolderPoolCapacity);

        var loopStart = e.Position;
        e.Emit(0x31, 0xC0);
        e.Emit(0xB9);
        e.EmitInt32(1);
        e.Emit(0xF0, 0x41, 0x0F, 0xB1, 0x4A, 0x0C);
        var jeFound = e.EmitJccRel8(0x74);
        e.Emit(0x49, 0x83, 0xC2, 0x20);
        e.Emit(0x41, 0xFF, 0xCB);
        e.EmitJccRel8To(0x75, loopStart);

        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_OUTOFMEMORY);
        e.Emit(0xC3);
        e.MarkRel8(jeFound);
    }

    private static void EmitInitHolderAndAddRefDriverX64(CodeEmitter e, IntPtr holderVtblAddr)
    {
        e.Emit(0x48, 0xB8);
        e.EmitInt64(holderVtblAddr.ToInt64());
        e.Emit(0x49, 0x89, 0x02);
        e.Emit(0x41, 0xC7, 0x42, 0x08, 0x01, 0x00, 0x00, 0x00);
        e.Emit(0x49, 0x89, 0x52, 0x10);
        e.Emit(0x4D, 0x89, 0x11);

        e.Emit(0x48, 0x85, 0xD2);
        var jzDone = e.EmitJccRel8(0x74);
        e.Emit(0x48, 0x83, 0xEC, 0x28);
        e.Emit(0x48, 0x89, 0xD1);
        e.Emit(0x48, 0x8B, 0x01);
        e.Emit(0xFF, 0x50, 0x08);
        e.Emit(0x48, 0x83, 0xC4, 0x28);

        e.MarkRel8(jzDone);
        e.Emit(0x31, 0xC0, 0xC3);
    }

    private static IntPtr EmitHolderAddRefX64(CodeEmitter e)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0xB8, 0x01, 0x00, 0x00, 0x00);
        e.Emit(0xF0, 0x0F, 0xC1, 0x41, 0x08);
        e.Emit(0xFF, 0xC0, 0xC3);
        return start;
    }

    private static IntPtr EmitHolderCloseX64(CodeEmitter e)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0x31, 0xD2);
        e.Emit(0x48, 0x87, 0x51, 0x10);
        e.Emit(0x48, 0x85, 0xD2);
        var jzDone = e.EmitJccRel8(0x74);
        e.Emit(0x48, 0x83, 0xEC, 0x28);
        e.Emit(0x48, 0x89, 0xD1);
        e.Emit(0x48, 0x8B, 0x01);
        e.Emit(0xFF, 0x50, 0x10);
        e.Emit(0x48, 0x83, 0xC4, 0x28);
        e.MarkRel8(jzDone);
        e.Emit(0x31, 0xC0, 0xC3);
        return start;
    }

    private static IntPtr EmitHolderReleaseX64(CodeEmitter e)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0xB8, 0xFF, 0xFF, 0xFF, 0xFF);
        e.Emit(0xF0, 0x0F, 0xC1, 0x41, 0x08);
        e.Emit(0xFF, 0xC8);
        var jnzAlive = e.EmitJccRel8(0x75);

        e.Emit(0x53);
        e.Emit(0x48, 0x83, 0xEC, 0x20);
        e.Emit(0x48, 0x89, 0xCB);
        e.Emit(0x31, 0xD2);
        e.Emit(0x48, 0x87, 0x53, 0x10);
        e.Emit(0x48, 0x85, 0xD2);
        var jzClear = e.EmitJccRel8(0x74);
        e.Emit(0x48, 0x89, 0xD1);
        e.Emit(0x48, 0x8B, 0x01);
        e.Emit(0xFF, 0x50, 0x10);

        e.MarkRel8(jzClear);
        e.Emit(0xC7, 0x43, 0x0C, 0x00, 0x00, 0x00, 0x00);
        e.Emit(0x48, 0x83, 0xC4, 0x20);
        e.Emit(0x5B);
        e.Emit(0x31, 0xC0);

        e.MarkRel8(jnzAlive);
        e.Emit(0xC3);
        return start;
    }

    private static IntPtr EmitHolderAllocResourceX64(CodeEmitter e)
    {
        e.Align16();
        var start = e.CurrentAddress;

        e.Emit(0x4D, 0x85, 0xC0);
        var jnzOk = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_POINTER);
        e.Emit(0xC3);

        e.MarkRel8(jnzOk);
        e.Emit(0x49, 0xC7, 0x00, 0x00, 0x00, 0x00, 0x00);
        e.Emit(0x48, 0x8B, 0x49, 0x10);
        e.Emit(0x48, 0x85, 0xC9);
        var jnzDriver = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_UNEXPECTED);
        e.Emit(0xC3);

        e.MarkRel8(jnzDriver);
        e.Emit(0x48, 0x83, 0xEC, 0x38);
        e.Emit(0x48, 0xC7, 0x44, 0x24, 0x20, 0x00, 0x00, 0x00, 0x00);
        e.Emit(0x4C, 0x8D, 0x4C, 0x24, 0x20);
        e.Emit(0x48, 0x8B, 0x01);
        e.Emit(0xFF, 0x50, 0x18);
        e.Emit(0x48, 0x83, 0xC4, 0x38);
        e.Emit(0xC3);
        return start;
    }

    private static IntPtr EmitHolderDestroyForwardX64(CodeEmitter e, byte driverVtableDisp, bool conditionalOnR8)
    {
        e.Align16();
        var start = e.CurrentAddress;
        var jzFromFlag = -1;
        if (conditionalOnR8)
        {
            e.Emit(0x45, 0x85, 0xC0);
            jzFromFlag = e.EmitJccRel8(0x74);
        }

        e.Emit(0x48, 0x8B, 0x49, 0x10);
        e.Emit(0x48, 0x85, 0xC9);
        var jzNoDriver = e.EmitJccRel8(0x74);

        e.Emit(0x48, 0x83, 0xEC, 0x28);
        e.Emit(0x48, 0x8B, 0x01);
        e.Emit(0xFF, 0x50, driverVtableDisp);
        e.Emit(0x48, 0x83, 0xC4, 0x28);
        e.Emit(0xC3);

        if (jzFromFlag >= 0)
        {
            e.MarkRel8(jzFromFlag);
        }

        e.MarkRel8(jzNoDriver);
        e.Emit(0x31, 0xC0, 0xC3);
        return start;
    }

#pragma warning disable S107
    private static StubMethodAddresses EmitAllMethodsX86(
        CodeEmitter e,
        IntPtr iidUnk,
        IntPtr iidCf,
        IntPtr iidDm,
        IntPtr iidHolder,
        IntPtr dmObj,
        IntPtr dataPage,
        IntPtr holderVtbl)
#pragma warning restore S107
    {
        var staticAddRef = EmitStaticConstReturnX86(e, 2, 4);
        var staticRelease = EmitStaticConstReturnX86(e, 1, 4);
        var zeroRet8 = EmitZeroReturnX86(e, 8);
        var cfQi = EmitQueryInterfaceX86(e, unchecked((uint)iidUnk.ToInt64()), unchecked((uint)iidCf.ToInt64()), addRefHolderSlot: false);
        var dmQiOffset = e.Align16();
        var dmQi = EmitQueryInterfaceX86(e, unchecked((uint)iidUnk.ToInt64()), unchecked((uint)iidDm.ToInt64()), addRefHolderSlot: false);
        var cfCreate = EmitCreateInstanceX86(e, unchecked((uint)dmObj.ToInt64()), dmQiOffset);
        var getContext = EmitGetContextX86(e);
        var regDisp = EmitRegisterDispenserX86(e, unchecked((uint)dataPage.ToInt64()), unchecked((uint)holderVtbl.ToInt64()));
        var holderQi = EmitQueryInterfaceX86(e, unchecked((uint)iidUnk.ToInt64()), unchecked((uint)iidHolder.ToInt64()), addRefHolderSlot: true);
        var holderAddRef = EmitHolderAddRefX86(e);
        var holderRelease = EmitHolderReleaseX86(e);
        var holderClose = EmitHolderCloseX86(e);
        var holderAlloc = EmitHolderAllocResourceX86(e);
        var holderFree = EmitHolderDestroyForwardX86(e, driverVtableDisp: 0x1C, conditionalOnArg2: false);
        var holderUntrack = EmitHolderDestroyForwardX86(e, driverVtableDisp: 0x1C, conditionalOnArg2: true);
        var holderUntrackS = EmitHolderDestroyForwardX86(e, driverVtableDisp: 0x20, conditionalOnArg2: true);

        return new StubMethodAddresses
        {
            StaticAddRef = staticAddRef,
            StaticRelease = staticRelease,
            LockServer = zeroRet8,
            ClassFactoryQi = cfQi,
            ClassFactoryCreateInstance = cfCreate,
            DispenserManagerQi = dmQi,
            RegisterDispenser = regDisp,
            GetContext = getContext,
            HolderQi = holderQi,
            HolderAddRef = holderAddRef,
            HolderRelease = holderRelease,
            HolderClose = holderClose,
            HolderAllocResource = holderAlloc,
            HolderFreeResource = holderFree,
            HolderTrackResource = zeroRet8,
            HolderUntrackResource = holderUntrack,
            HolderUntrackResourceS = holderUntrackS,
        };
    }

    private static IntPtr EmitStaticConstReturnX86(CodeEmitter e, int value, ushort popBytes)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0xB8);
        e.EmitInt32(value);
        e.Emit(0xC2);
        e.EmitUInt16(popBytes);
        return start;
    }

    private static IntPtr EmitZeroReturnX86(CodeEmitter e, ushort popBytes)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0x31, 0xC0, 0xC2);
        e.EmitUInt16(popBytes);
        return start;
    }

    private static IntPtr EmitQueryInterfaceX86(CodeEmitter e, uint iidUnkAddr, uint iidTargetAddr, bool addRefHolderSlot)
    {
        e.Align16();
        var start = e.CurrentAddress;

        e.Emit(0x55, 0x89, 0xE5);
        e.Emit(0x8B, 0x55, 0x10, 0x85, 0xD2);
        var jnzPpvOk = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_POINTER);
        e.Emit(0x5D, 0xC2, 0x0C, 0x00);

        e.MarkRel8(jnzPpvOk);
        e.Emit(0xC7, 0x02, 0x00, 0x00, 0x00, 0x00);

        e.Emit(0x8B, 0x4D, 0x0C, 0x85, 0xC9);
        var jnzRiidOk = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_INVALIDARG);
        e.Emit(0x5D, 0xC2, 0x0C, 0x00);

        e.MarkRel8(jnzRiidOk);
        e.Emit(0x56);

        var jeMatchUnk = EmitGuidMatchBranchX86(e, iidUnkAddr);
        var jeMatchTarget = EmitGuidMatchBranchX86(e, iidTargetAddr);

        e.Emit(0x5E, 0xB8);
        e.EmitInt32(NativeMethods.E_NOINTERFACE);
        e.Emit(0x5D, 0xC2, 0x0C, 0x00);

        e.MarkRel8(jeMatchUnk);
        e.MarkRel8(jeMatchTarget);
        e.Emit(0x5E, 0x8B, 0x45, 0x08, 0x89, 0x02);
        if (addRefHolderSlot)
        {
            e.Emit(0xF0, 0xFF, 0x40, 0x08);
        }

        e.Emit(0x31, 0xC0, 0x5D, 0xC2, 0x0C, 0x00);
        return start;
    }

    private static int EmitGuidMatchBranchX86(CodeEmitter e, uint targetGuidAddr)
    {
        e.Emit(0xBE);
        e.EmitUInt32(targetGuidAddr);
        e.Emit(0x8B, 0x01, 0x3B, 0x06);
        var jne1 = e.EmitJccRel8(0x75);
        e.Emit(0x8B, 0x41, 0x04, 0x3B, 0x46, 0x04);
        var jne2 = e.EmitJccRel8(0x75);
        e.Emit(0x8B, 0x41, 0x08, 0x3B, 0x46, 0x08);
        var jne3 = e.EmitJccRel8(0x75);
        e.Emit(0x8B, 0x41, 0x0C, 0x3B, 0x46, 0x0C);
        var jeHit = e.EmitJccRel8(0x74);
        e.MarkRel8(jne1);
        e.MarkRel8(jne2);
        e.MarkRel8(jne3);
        return jeHit;
    }

    private static IntPtr EmitCreateInstanceX86(CodeEmitter e, uint dmObjAddr, int dmQiOffset)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0x55, 0x89, 0xE5);
        e.Emit(0x8B, 0x55, 0x14, 0x85, 0xD2);
        var jnzPpvOk = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_POINTER);
        e.Emit(0x5D, 0xC2, 0x10, 0x00);

        e.MarkRel8(jnzPpvOk);
        e.Emit(0xC7, 0x02, 0x00, 0x00, 0x00, 0x00);
        e.Emit(0x8B, 0x45, 0x0C, 0x85, 0xC0);
        var jzNoAgg = e.EmitJccRel8(0x74);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.CLASS_E_NOAGGREGATION);
        e.Emit(0x5D, 0xC2, 0x10, 0x00);

        e.MarkRel8(jzNoAgg);
        e.Emit(0x52);
        e.Emit(0xFF, 0x75, 0x10);
        e.Emit(0x68);
        e.EmitUInt32(dmObjAddr);
        e.EmitRel32JumpOrCall(0xE8, dmQiOffset);
        e.Emit(0x5D, 0xC2, 0x10, 0x00);
        return start;
    }

    private static IntPtr EmitGetContextX86(CodeEmitter e)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0x8B, 0x4C, 0x24, 0x08, 0x85, 0xC9);
        var jzSkip1 = e.EmitJccRel8(0x74);
        e.Emit(0xC7, 0x01, 0x00, 0x00, 0x00, 0x00);
        e.MarkRel8(jzSkip1);

        e.Emit(0x8B, 0x4C, 0x24, 0x0C, 0x85, 0xC9);
        var jzSkip2 = e.EmitJccRel8(0x74);
        e.Emit(0xC7, 0x01, 0x00, 0x00, 0x00, 0x00);
        e.MarkRel8(jzSkip2);

        e.Emit(0x31, 0xC0, 0xC2, 0x0C, 0x00);
        return start;
    }

    private static IntPtr EmitRegisterDispenserX86(CodeEmitter e, uint dataPageAddr, uint holderVtblAddr)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0x55, 0x89, 0xE5);
        e.Emit(0x8B, 0x55, 0x14, 0x85, 0xD2);
        var jnzOk = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_POINTER);
        e.Emit(0x5D, 0xC2, 0x10, 0x00);

        e.MarkRel8(jnzOk);
        e.Emit(0xC7, 0x02, 0x00, 0x00, 0x00, 0x00);
        e.Emit(0x53, 0x56);

        EmitClaimHolderSlotX86(e, dataPageAddr);
        EmitInitHolderAndAddRefDriverX86(e, holderVtblAddr);
        return start;
    }

    private static void EmitClaimHolderSlotX86(CodeEmitter e, uint dataPageAddr)
    {
        e.Emit(0xBE);
        e.EmitUInt32(dataPageAddr);
        e.Emit(0xBB);
        e.EmitInt32(HolderPoolCapacity);

        var loopStart = e.Position;
        e.Emit(0x31, 0xC0);
        e.Emit(0xB9);
        e.EmitInt32(1);
        e.Emit(0xF0, 0x0F, 0xB1, 0x4E, 0x0C);
        var jeFound = e.EmitJccRel8(0x74);
        e.Emit(0x83, 0xC6, 0x20);
        e.Emit(0x4B);
        e.EmitJccRel8To(0x75, loopStart);

        e.Emit(0x5E, 0x5B, 0xB8);
        e.EmitInt32(NativeMethods.E_OUTOFMEMORY);
        e.Emit(0x5D, 0xC2, 0x10, 0x00);
        e.MarkRel8(jeFound);
    }

    private static void EmitInitHolderAndAddRefDriverX86(CodeEmitter e, uint holderVtblAddr)
    {
        e.Emit(0xC7, 0x06);
        e.EmitUInt32(holderVtblAddr);
        e.Emit(0xC7, 0x46, 0x08, 0x01, 0x00, 0x00, 0x00);
        e.Emit(0x8B, 0x4D, 0x0C);
        e.Emit(0x89, 0x4E, 0x10);
        e.Emit(0x8B, 0x55, 0x14);
        e.Emit(0x89, 0x32);
        e.Emit(0x85, 0xC9);
        var jzDone = e.EmitJccRel8(0x74);
        e.Emit(0x8B, 0x01, 0x51, 0xFF, 0x50, 0x04);

        e.MarkRel8(jzDone);
        e.Emit(0x5E, 0x5B, 0x31, 0xC0, 0x5D, 0xC2, 0x10, 0x00);
    }

    private static IntPtr EmitHolderAddRefX86(CodeEmitter e)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0x8B, 0x4C, 0x24, 0x04);
        e.Emit(0xB8, 0x01, 0x00, 0x00, 0x00);
        e.Emit(0xF0, 0x0F, 0xC1, 0x41, 0x08);
        e.Emit(0x40, 0xC2, 0x04, 0x00);
        return start;
    }

    private static IntPtr EmitHolderCloseX86(CodeEmitter e)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0x55, 0x89, 0xE5);
        e.Emit(0x8B, 0x4D, 0x08, 0x31, 0xD2, 0x87, 0x51, 0x10, 0x85, 0xD2);
        var jzDone = e.EmitJccRel8(0x74);
        e.Emit(0x8B, 0x02, 0x52, 0xFF, 0x50, 0x08);
        e.MarkRel8(jzDone);
        e.Emit(0x31, 0xC0, 0x5D, 0xC2, 0x04, 0x00);
        return start;
    }

    private static IntPtr EmitHolderReleaseX86(CodeEmitter e)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0x55, 0x89, 0xE5);
        e.Emit(0x8B, 0x4D, 0x08);
        e.Emit(0xB8, 0xFF, 0xFF, 0xFF, 0xFF);
        e.Emit(0xF0, 0x0F, 0xC1, 0x41, 0x08);
        e.Emit(0x48);
        var jnzAlive = e.EmitJccRel8(0x75);

        e.Emit(0x53, 0x89, 0xCB, 0x31, 0xD2, 0x87, 0x53, 0x10, 0x85, 0xD2);
        var jzClear = e.EmitJccRel8(0x74);
        e.Emit(0x8B, 0x02, 0x52, 0xFF, 0x50, 0x08);

        e.MarkRel8(jzClear);
        e.Emit(0xC7, 0x43, 0x0C, 0x00, 0x00, 0x00, 0x00);
        e.Emit(0x5B, 0x31, 0xC0);

        e.MarkRel8(jnzAlive);
        e.Emit(0x5D, 0xC2, 0x04, 0x00);
        return start;
    }

    private static IntPtr EmitHolderAllocResourceX86(CodeEmitter e)
    {
        e.Align16();
        var start = e.CurrentAddress;
        e.Emit(0x55, 0x89, 0xE5, 0x83, 0xEC, 0x04);
        e.Emit(0x8B, 0x55, 0x10, 0x85, 0xD2);
        var jnzOk = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_POINTER);
        e.Emit(0x89, 0xEC, 0x5D, 0xC2, 0x0C, 0x00);

        e.MarkRel8(jnzOk);
        e.Emit(0xC7, 0x02, 0x00, 0x00, 0x00, 0x00);
        e.Emit(0x8B, 0x4D, 0x08, 0x8B, 0x49, 0x10, 0x85, 0xC9);
        var jnzDriver = e.EmitJccRel8(0x75);
        e.Emit(0xB8);
        e.EmitInt32(NativeMethods.E_UNEXPECTED);
        e.Emit(0x89, 0xEC, 0x5D, 0xC2, 0x0C, 0x00);

        e.MarkRel8(jnzDriver);
        e.Emit(0xC7, 0x45, 0xFC, 0x00, 0x00, 0x00, 0x00);
        e.Emit(0x8D, 0x45, 0xFC, 0x50, 0x52, 0xFF, 0x75, 0x0C, 0x51);
        e.Emit(0x8B, 0x01, 0xFF, 0x50, 0x0C);
        e.Emit(0x89, 0xEC, 0x5D, 0xC2, 0x0C, 0x00);
        return start;
    }

    private static IntPtr EmitHolderDestroyForwardX86(CodeEmitter e, byte driverVtableDisp, bool conditionalOnArg2)
    {
        e.Align16();
        var start = e.CurrentAddress;
        var popBytes = conditionalOnArg2 ? (ushort)12 : (ushort)8;
        e.Emit(0x55, 0x89, 0xE5);

        var jzFlag = -1;
        if (conditionalOnArg2)
        {
            e.Emit(0x8B, 0x45, 0x10, 0x85, 0xC0);
            jzFlag = e.EmitJccRel8(0x74);
        }

        e.Emit(0x8B, 0x4D, 0x08, 0x8B, 0x49, 0x10, 0x85, 0xC9);
        var jzNoDriver = e.EmitJccRel8(0x74);

        e.Emit(0xFF, 0x75, 0x0C, 0x51, 0x8B, 0x01, 0xFF, 0x50, driverVtableDisp);
        e.Emit(0x5D, 0xC2);
        e.EmitUInt16(popBytes);

        if (jzFlag >= 0)
        {
            e.MarkRel8(jzFlag);
        }

        e.MarkRel8(jzNoDriver);
        e.Emit(0x31, 0xC0, 0x5D, 0xC2);
        e.EmitUInt16(popBytes);
        return start;
    }

    private static IntPtr OffsetAddress(IntPtr baseAddress, int offset)
        => new(baseAddress.ToInt64() + offset);

    private static void WriteGuid(IntPtr destination, Guid guid)
    {
        var bytes = guid.ToByteArray();
        Marshal.Copy(bytes, 0, destination, bytes.Length);
    }

    private static void WriteVtableSlot(IntPtr vtableBase, int slotIndex, IntPtr methodAddress)
        => Marshal.WriteIntPtr(vtableBase, slotIndex * IntPtr.Size, methodAddress);

    private sealed class StubMethodAddresses
    {
        internal IntPtr StaticAddRef { get; set; }

        internal IntPtr StaticRelease { get; set; }

        internal IntPtr LockServer { get; set; }

        internal IntPtr ClassFactoryQi { get; set; }

        internal IntPtr ClassFactoryCreateInstance { get; set; }

        internal IntPtr DispenserManagerQi { get; set; }

        internal IntPtr RegisterDispenser { get; set; }

        internal IntPtr GetContext { get; set; }

        internal IntPtr HolderQi { get; set; }

        internal IntPtr HolderAddRef { get; set; }

        internal IntPtr HolderRelease { get; set; }

        internal IntPtr HolderClose { get; set; }

        internal IntPtr HolderAllocResource { get; set; }

        internal IntPtr HolderFreeResource { get; set; }

        internal IntPtr HolderTrackResource { get; set; }

        internal IntPtr HolderUntrackResource { get; set; }

        internal IntPtr HolderUntrackResourceS { get; set; }
    }

    private sealed class CodeEmitter
    {
        private readonly IntPtr _basePage;

        internal CodeEmitter(IntPtr basePage, int startOffset)
        {
            _basePage = basePage;
            Position = startOffset;
        }

        internal int Position { get; private set; }

        internal IntPtr CurrentAddress => OffsetAddress(_basePage, Position);

        internal int Align16()
        {
            while ((Position & 0x0F) != 0)
            {
                Emit(0xCC);
            }

            return Position;
        }

        internal void Emit(byte b)
        {
            Marshal.WriteByte(_basePage, Position, b);
            Position++;
        }

        internal void Emit(params byte[] bytes)
        {
            for (var i = 0; i < bytes.Length; i++)
            {
                Emit(bytes[i]);
            }
        }

        internal void EmitUInt16(ushort value)
        {
            Marshal.WriteInt16(_basePage, Position, unchecked((short)value));
            Position += 2;
        }

        internal void EmitInt32(int value)
        {
            Marshal.WriteInt32(_basePage, Position, value);
            Position += 4;
        }

        internal void EmitUInt32(uint value)
            => EmitInt32(unchecked((int)value));

        internal void EmitInt64(long value)
        {
            Marshal.WriteInt64(_basePage, Position, value);
            Position += 8;
        }

        internal int EmitJccRel8(byte opcode)
        {
            Emit(opcode);
            var dispPos = Position;
            Emit(0x00);
            return dispPos;
        }

        internal void MarkRel8(int dispPos)
        {
            var rel = checked((sbyte)(Position - (dispPos + 1)));
            Marshal.WriteByte(_basePage, dispPos, unchecked((byte)rel));
        }

        internal void EmitJccRel8To(byte opcode, int targetPos)
        {
            Emit(opcode);
            var rel = checked((sbyte)(targetPos - (Position + 1)));
            Emit(unchecked((byte)rel));
        }

        internal void EmitRel32JumpOrCall(byte opcode, int targetPos)
        {
            Emit(opcode);
            var rel = checked(targetPos - (Position + 4));
            EmitInt32(rel);
        }
    }
}
#pragma warning restore SCS0018, SEC0116
#pragma warning restore S101
#pragma warning restore SA1121
#pragma warning restore MA0158
