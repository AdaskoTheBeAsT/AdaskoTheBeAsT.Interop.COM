using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1121 // Use built-in type alias
/// <summary>
/// Diagnostic utilities for inspecting and verifying which native PE module (<c>HMODULE</c>) backs
/// a COM object's interface vtable, guarding multi-worker hosts against accidental fallback to globally
/// registered COM servers.
/// </summary>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public static class ComModuleDiagnostics
{
#pragma warning disable MA0176 // Optimize guid creation
    private static readonly Guid IidIUnknown = new("00000000-0000-0000-C000-000000000046");
#pragma warning restore MA0176

    private static readonly HashSet<string> SharedRuntimeAndProxyModuleNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "combase.dll",
        "ole32.dll",
        "oleaut32.dll",
        "rpcrt4.dll",
        "msvbvm60.dll",
        "msvbvm50.dll",
        "vbrun300.dll",
        "clr.dll",
        "coreclr.dll",
        "mscorwks.dll",
        "ntdll.dll",
        "kernelbase.dll",
        "kernel32.dll",
    };

    /// <summary>
    /// Resolves the native <c>HMODULE</c> owning the vtable function pointer at <paramref name="methodSlotIndex"/>
    /// for <paramref name="comObject"/> (on <c>IUnknown</c> by default, or on <paramref name="interfaceId"/> when specified).
    /// </summary>
    /// <param name="comObject">The COM object or RCW to inspect.</param>
    /// <param name="interfaceId">Optional specific native COM interface IID to <c>QueryInterface</c> before reading the vtable.</param>
    /// <param name="methodSlotIndex">Zero-based vtable slot index to inspect (slot <c>0</c> is <c>QueryInterface</c>; custom methods typically start at slot <c>3</c>).</param>
    /// <returns>The owning <c>HMODULE</c>, or <see cref="IntPtr.Zero"/> if the vtable slot does not map to a loaded PE module.</returns>
    public static IntPtr GetInstanceModuleHandle(
        object comObject,
        Guid? interfaceId = null,
        int methodSlotIndex = 0)
    {
        ThrowHelper.ThrowIfNull(comObject, nameof(comObject));
        if (methodSlotIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(methodSlotIndex), "Method slot index cannot be negative.");
        }

        ReadVtableSlot(comObject, interfaceId, methodSlotIndex, out _, out _, out var moduleHandle);
        return moduleHandle;
    }

    /// <summary>
    /// Inspects the native module backing <paramref name="comObject"/> without throwing on mismatch or inconclusive runtime/proxy vtables.
    /// </summary>
    /// <param name="comObject">The COM object or RCW to inspect.</param>
    /// <param name="expectedModuleHandle">Expected isolated <c>HMODULE</c> (may be <see cref="IntPtr.Zero"/> when <paramref name="expectedDllPath"/> is provided).</param>
    /// <param name="expectedDllPath">Optional expected DLL path used to resolve or cross-check module identity.</param>
    /// <param name="interfaceId">Optional specific native COM interface IID to inspect instead of <c>IUnknown</c>.</param>
    /// <param name="methodSlotIndex">Zero-based vtable slot index to inspect.</param>
    /// <returns>A structured <see cref="ComModuleInspectionResult"/> describing whether the vtable matched, mismatched, or was inconclusive.</returns>
    public static ComModuleInspectionResult InspectInstanceModule(
        object comObject,
        IntPtr expectedModuleHandle,
        string? expectedDllPath = null,
        Guid? interfaceId = null,
        int methodSlotIndex = 0)
    {
        ThrowHelper.ThrowIfNull(comObject, nameof(comObject));
        if (methodSlotIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(methodSlotIndex), "Method slot index cannot be negative.");
        }

        var normalizedExpectedPath = expectedDllPath is not null && !string.IsNullOrWhiteSpace(expectedDllPath)
            ? Path.GetFullPath(expectedDllPath)
            : null;

        if (expectedModuleHandle == IntPtr.Zero && normalizedExpectedPath is null)
        {
            throw new ArgumentException(
                "Either expectedModuleHandle must be non-zero or expectedDllPath must be provided.",
                nameof(expectedModuleHandle));
        }

        var effectiveExpectedHandle = ResolveEffectiveExpectedHandle(expectedModuleHandle, normalizedExpectedPath);
        var effectiveIid = interfaceId ?? IidIUnknown;
        ReadVtableSlot(comObject, interfaceId, methodSlotIndex, out var vtablePtr, out var methodPtr, out var actualModule);
        var actualPath = NativeMethods.TryGetModuleFilePath(actualModule);
        var normalizedActualPath = actualPath is not null && !string.IsNullOrWhiteSpace(actualPath)
            ? Path.GetFullPath(actualPath)
            : null;

        return CreateInspectionResult(
            actualModule,
            normalizedActualPath,
            effectiveExpectedHandle,
            normalizedExpectedPath,
            vtablePtr,
            methodPtr,
            methodSlotIndex,
            effectiveIid);
    }

    /// <summary>
    /// Verifies that <paramref name="comObject"/> is backed by <paramref name="expectedModuleHandle"/> (and/or <paramref name="expectedDllPath"/>),
    /// throwing <see cref="InvalidOperationException"/> when a concrete module mismatch is detected.
    /// </summary>
    /// <param name="comObject">The COM object or RCW to verify.</param>
    /// <param name="expectedModuleHandle">Expected isolated <c>HMODULE</c>.</param>
    /// <param name="expectedDllPath">Optional expected DLL path.</param>
    /// <param name="interfaceId">Optional specific native COM interface IID to inspect instead of <c>IUnknown</c>.</param>
    /// <param name="methodSlotIndex">Zero-based vtable slot index to inspect.</param>
    /// <param name="throwOnInconclusive">
    /// When <see langword="true"/>, also throws <see cref="InvalidOperationException"/> if the vtable belongs to a shared
    /// runtime or marshaling proxy (<see cref="ComModuleVerificationStatus.Inconclusive"/>). Defaults to <see langword="false"/>.
    /// </param>
    /// <returns>The <see cref="ComModuleInspectionResult"/> when verification does not throw.</returns>
    public static ComModuleInspectionResult VerifyInstanceModule(
        object comObject,
        IntPtr expectedModuleHandle,
        string? expectedDllPath = null,
        Guid? interfaceId = null,
        int methodSlotIndex = 0,
        bool throwOnInconclusive = false)
    {
        var inspection = InspectInstanceModule(
            comObject,
            expectedModuleHandle,
            expectedDllPath,
            interfaceId,
            methodSlotIndex);

        if (inspection.Status == ComModuleVerificationStatus.Mismatch
            || (throwOnInconclusive && inspection.Status == ComModuleVerificationStatus.Inconclusive))
        {
            throw new InvalidOperationException(inspection.DiagnosticMessage);
        }

        return inspection;
    }

    /// <summary>
    /// Instantiates <paramref name="clsid"/> as an in-process COM server (<c>CLSCTX_INPROC_SERVER</c>) via a raw native
    /// interface pointer without allocating a managed Runtime Callable Wrapper (RCW), resolves the normalized file path
    /// of the native PE module owning vtable slot <paramref name="methodSlotIndex"/>, and immediately releases the native
    /// pointer before returning.
    /// </summary>
    /// <param name="clsid">The COM class identifier (<c>CLSID</c>) to probe in the active activation context or apartment.</param>
    /// <param name="interfaceId">Optional native COM interface IID to request from <c>CoCreateInstance</c> instead of <c>IUnknown</c>.</param>
    /// <param name="methodSlotIndex">Zero-based vtable slot index to inspect (defaults to slot <c>0</c>, <c>QueryInterface</c>).</param>
    /// <returns>The full path of the backing native PE module, or <see langword="null"/> if the vtable slot does not map to a loaded PE module.</returns>
    public static string? GetRegisteredInprocServerPath(
        Guid clsid,
        Guid? interfaceId = null,
        int methodSlotIndex = 0)
    {
        ValidateClsidAndSlot(clsid, methodSlotIndex);
        var effectiveIid = interfaceId ?? IidIUnknown;
        ReadInprocServerVtableSlot(clsid, effectiveIid, methodSlotIndex, out _, out _, out var actualModule);
        var actualPath = NativeMethods.TryGetModuleFilePath(actualModule);
        return actualPath is not null && !string.IsNullOrWhiteSpace(actualPath)
            ? Path.GetFullPath(actualPath)
            : null;
    }

    /// <summary>
    /// Instantiates <paramref name="clsid"/> as an in-process COM server (<c>CLSCTX_INPROC_SERVER</c>) without allocating
    /// a managed RCW, inspects the native PE module owning vtable slot <paramref name="methodSlotIndex"/> against
    /// <paramref name="expectedModuleHandle"/> and/or <paramref name="expectedDllPath"/>, and releases the native pointer immediately.
    /// </summary>
    /// <param name="clsid">The COM class identifier (<c>CLSID</c>) to inspect.</param>
    /// <param name="expectedModuleHandle">Expected isolated <c>HMODULE</c> (may be <see cref="IntPtr.Zero"/> when <paramref name="expectedDllPath"/> is provided).</param>
    /// <param name="expectedDllPath">Optional expected DLL path used to resolve or cross-check module identity.</param>
    /// <param name="interfaceId">Optional native COM interface IID to request instead of <c>IUnknown</c>.</param>
    /// <param name="methodSlotIndex">Zero-based vtable slot index to inspect.</param>
    /// <returns>A structured <see cref="ComModuleInspectionResult"/> describing whether the server matched, mismatched, or was inconclusive.</returns>
    public static ComModuleInspectionResult InspectRegisteredInprocServer(
        Guid clsid,
        IntPtr expectedModuleHandle,
        string? expectedDllPath = null,
        Guid? interfaceId = null,
        int methodSlotIndex = 0)
    {
        ValidateClsidAndSlot(clsid, methodSlotIndex);
        var normalizedExpectedPath = expectedDllPath is not null && !string.IsNullOrWhiteSpace(expectedDllPath)
            ? Path.GetFullPath(expectedDllPath)
            : null;

        if (expectedModuleHandle == IntPtr.Zero && normalizedExpectedPath is null)
        {
            throw new ArgumentException(
                "Either expectedModuleHandle must be non-zero or expectedDllPath must be provided.",
                nameof(expectedModuleHandle));
        }

        var effectiveIid = interfaceId ?? IidIUnknown;
        ReadInprocServerVtableSlot(clsid, effectiveIid, methodSlotIndex, out var vtablePtr, out var methodPtr, out var actualModule);

        // Resolve expected handle after CoCreateInstance as well so modules loaded lazily by Desktop WinSxS on first
        // CoCreateInstance resolve cleanly when expectedModuleHandle was IntPtr.Zero.
        var effectiveExpectedHandle = ResolveEffectiveExpectedHandle(expectedModuleHandle, normalizedExpectedPath);
        var actualPath = NativeMethods.TryGetModuleFilePath(actualModule);
        var normalizedActualPath = actualPath is not null && !string.IsNullOrWhiteSpace(actualPath)
            ? Path.GetFullPath(actualPath)
            : null;

        return CreateInspectionResult(
            actualModule,
            normalizedActualPath,
            effectiveExpectedHandle,
            normalizedExpectedPath,
            vtablePtr,
            methodPtr,
            methodSlotIndex,
            effectiveIid);
    }

    /// <summary>
    /// Instantiates <paramref name="clsid"/> as an in-process COM server (<c>CLSCTX_INPROC_SERVER</c>) without allocating
    /// a managed RCW, verifies that it is backed by <paramref name="expectedModuleHandle"/> and/or <paramref name="expectedDllPath"/>,
    /// releases the native pointer immediately, and throws <see cref="InvalidOperationException"/> on mismatch.
    /// </summary>
    /// <param name="clsid">The COM class identifier (<c>CLSID</c>) to verify.</param>
    /// <param name="expectedModuleHandle">Expected isolated <c>HMODULE</c> (may be <see cref="IntPtr.Zero"/> when <paramref name="expectedDllPath"/> is provided).</param>
    /// <param name="expectedDllPath">Optional expected DLL path.</param>
    /// <param name="interfaceId">Optional native COM interface IID to request instead of <c>IUnknown</c>.</param>
    /// <param name="methodSlotIndex">Zero-based vtable slot index to inspect.</param>
    /// <param name="throwOnInconclusive">
    /// When <see langword="true"/>, also throws <see cref="InvalidOperationException"/> if the vtable belongs to a shared
    /// runtime or marshaling proxy (<see cref="ComModuleVerificationStatus.Inconclusive"/>). Defaults to <see langword="false"/>.
    /// </param>
    /// <returns>The <see cref="ComModuleInspectionResult"/> when verification succeeds.</returns>
    public static ComModuleInspectionResult VerifyRegisteredInprocServer(
        Guid clsid,
        IntPtr expectedModuleHandle,
        string? expectedDllPath = null,
        Guid? interfaceId = null,
        int methodSlotIndex = 0,
        bool throwOnInconclusive = false)
    {
        var inspection = InspectRegisteredInprocServer(
            clsid,
            expectedModuleHandle,
            expectedDllPath,
            interfaceId,
            methodSlotIndex);

        if (inspection.Status == ComModuleVerificationStatus.Mismatch
            || (throwOnInconclusive && inspection.Status == ComModuleVerificationStatus.Inconclusive))
        {
            throw new InvalidOperationException(inspection.DiagnosticMessage);
        }

        return inspection;
    }

    private static void ValidateClsidAndSlot(Guid clsid, int methodSlotIndex)
    {
        if (clsid == Guid.Empty)
        {
            throw new ArgumentException("CLSID cannot be Guid.Empty.", nameof(clsid));
        }

        if (methodSlotIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(methodSlotIndex), "Method slot index cannot be negative.");
        }
    }

    private static void ReadInprocServerVtableSlot(
        Guid clsid,
        Guid targetIid,
        int methodSlotIndex,
        out IntPtr vtablePtr,
        out IntPtr methodPtr,
        out IntPtr moduleHandle)
    {
        NanoServerOptions.EnsureCurrentStaThreadInitialized();
        var localClsid = clsid;
        var localIid = targetIid;
        var hr = NativeMethods.CoCreateInstance(
            ref localClsid,
            IntPtr.Zero,
            NativeMethods.CLSCTX_INPROC_SERVER,
            ref localIid,
            out var rawPtr);
        if (hr < 0 || rawPtr == IntPtr.Zero)
        {
            throw ComErrorDiagnostics.CreateExceptionForHr(hr < 0 ? hr : NativeMethods.E_NOINTERFACE);
        }

        try
        {
            ReadRawInterfaceVtableSlot(rawPtr, methodSlotIndex, out vtablePtr, out methodPtr, out moduleHandle);
        }
        finally
        {
            _ = Marshal.Release(rawPtr);
        }
    }

    private static IntPtr ResolveEffectiveExpectedHandle(IntPtr expectedModuleHandle, string? normalizedExpectedPath)
    {
        if (expectedModuleHandle == IntPtr.Zero
            && normalizedExpectedPath is not null
            && NativeMethods.GetModuleHandleExByName(
                NativeMethods.GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                normalizedExpectedPath,
                out var resolvedExpected))
        {
            return resolvedExpected;
        }

        return expectedModuleHandle;
    }

#pragma warning disable S107
    private static ComModuleInspectionResult CreateInspectionResult(
        IntPtr actualModule,
        string? normalizedActualPath,
        IntPtr effectiveExpectedHandle,
        string? normalizedExpectedPath,
        IntPtr vtablePtr,
        IntPtr methodPtr,
        int methodSlotIndex,
        Guid effectiveIid)
#pragma warning restore S107
    {
        if (IsExpectedModuleMatch(actualModule, normalizedActualPath, effectiveExpectedHandle, normalizedExpectedPath))
        {
            return new ComModuleInspectionResult(
                ComModuleVerificationStatus.Match,
                actualModule,
                normalizedActualPath,
                vtablePtr,
                methodPtr,
                methodSlotIndex,
                effectiveIid,
                string.Format(
                    CultureInfo.InvariantCulture,
                    "COM interface {0:B} slot {1} matched expected module 0x{2:X} ({3}).",
                    effectiveIid,
                    methodSlotIndex,
                    actualModule.ToInt64(),
                    normalizedActualPath ?? "<unknown>"));
        }

        if (IsInconclusiveModule(actualModule, normalizedActualPath))
        {
            return new ComModuleInspectionResult(
                ComModuleVerificationStatus.Inconclusive,
                actualModule,
                normalizedActualPath,
                vtablePtr,
                methodPtr,
                methodSlotIndex,
                effectiveIid,
                string.Format(
                    CultureInfo.InvariantCulture,
                    "COM interface {0:B} slot {1} resolved to 0x{2:X} ({3}), which is a shared runtime, proxy/stub, tear-off, or dynamic thunk rather than direct proof of vendor DLL isolation (expected 0x{4:X} {5}). Inspect a direct native interface or custom method slot.",
                    effectiveIid,
                    methodSlotIndex,
                    actualModule.ToInt64(),
                    normalizedActualPath ?? "<unmapped thunk>",
                    effectiveExpectedHandle.ToInt64(),
                    normalizedExpectedPath ?? string.Empty));
        }

        return new ComModuleInspectionResult(
            ComModuleVerificationStatus.Mismatch,
            actualModule,
            normalizedActualPath,
            vtablePtr,
            methodPtr,
            methodSlotIndex,
            effectiveIid,
            string.Format(
                CultureInfo.InvariantCulture,
                "COM module isolation mismatch on interface {0:B} slot {1}: expected HMODULE 0x{2:X} ({3}), but instance vtable belongs to HMODULE 0x{4:X} ({5}).",
                effectiveIid,
                methodSlotIndex,
                effectiveExpectedHandle.ToInt64(),
                normalizedExpectedPath ?? "<unspecified>",
                actualModule.ToInt64(),
                normalizedActualPath ?? "<unknown>"));
    }

    private static void ReadVtableSlot(
        object comObject,
        Guid? interfaceId,
        int methodSlotIndex,
        out IntPtr vtablePtr,
        out IntPtr methodPtr,
        out IntPtr moduleHandle)
    {
        var unk = Marshal.GetIUnknownForObject(comObject);
        var targetInterfacePtr = IntPtr.Zero;
        try
        {
            if (interfaceId.HasValue && interfaceId.Value != IidIUnknown)
            {
                var iid = interfaceId.Value;
                var hr = NativeMethods.QueryInterface(unk, ref iid, out targetInterfacePtr);
                if (hr < 0 || targetInterfacePtr == IntPtr.Zero)
                {
                    Marshal.ThrowExceptionForHR(hr < 0 ? hr : NativeMethods.E_NOINTERFACE);
                }
            }

            var effectivePtr = targetInterfacePtr != IntPtr.Zero ? targetInterfacePtr : unk;
            ReadRawInterfaceVtableSlot(effectivePtr, methodSlotIndex, out vtablePtr, out methodPtr, out moduleHandle);
        }
        finally
        {
            if (targetInterfacePtr != IntPtr.Zero)
            {
                _ = Marshal.Release(targetInterfacePtr);
            }

            if (unk != IntPtr.Zero)
            {
                _ = Marshal.Release(unk);
            }
        }
    }

    private static void ReadRawInterfaceVtableSlot(
        IntPtr interfacePtr,
        int methodSlotIndex,
        out IntPtr vtablePtr,
        out IntPtr methodPtr,
        out IntPtr moduleHandle)
    {
        vtablePtr = Marshal.ReadIntPtr(interfacePtr);
        if (vtablePtr == IntPtr.Zero)
        {
            methodPtr = IntPtr.Zero;
            moduleHandle = IntPtr.Zero;
            return;
        }

        methodPtr = Marshal.ReadIntPtr(vtablePtr, checked(methodSlotIndex * IntPtr.Size));
        if (methodPtr == IntPtr.Zero
            || !NativeMethods.GetModuleHandleExFromAddress(
                NativeMethods.GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | NativeMethods.GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                methodPtr,
                out moduleHandle))
        {
            moduleHandle = IntPtr.Zero;
        }
    }

    private static bool IsExpectedModuleMatch(
        IntPtr actualModule,
        string? normalizedActualPath,
        IntPtr expectedModule,
        string? normalizedExpectedPath)
    {
        if (actualModule == IntPtr.Zero)
        {
            return false;
        }

        if (expectedModule != IntPtr.Zero && actualModule != expectedModule)
        {
            return false;
        }

        if (normalizedExpectedPath is not null
            && (normalizedActualPath is null
                || !string.Equals(normalizedActualPath, normalizedExpectedPath, StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        return expectedModule != IntPtr.Zero || normalizedExpectedPath is not null;
    }

    private static bool IsInconclusiveModule(IntPtr actualModule, string? normalizedActualPath)
    {
        if (actualModule == IntPtr.Zero || normalizedActualPath is null)
        {
            return true;
        }

        var fileName = Path.GetFileName(normalizedActualPath);
        return SharedRuntimeAndProxyModuleNames.Contains(fileName);
    }
}
#pragma warning restore SA1121
