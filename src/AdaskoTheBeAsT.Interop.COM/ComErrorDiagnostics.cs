using System;
using System.Globalization;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1121 // Use built-in type alias
/// <summary>
/// Diagnostic helpers for reading or peeking thread-local COM <c>IErrorInfo</c> records,
/// enriching <see cref="COMException"/> instances, and unpacking Visual Basic 6 / ActiveX
/// <c>FACILITY_CONTROL</c>, <c>DISP_E_*</c>, and <c>FACILITY_ITF</c> <c>HRESULT</c> codes.
/// </summary>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public static class ComErrorDiagnostics
{
    private const uint FacilityMask = 0xFFFF0000u;
    private const uint FacilityControlPrefix = 0x800A0000u;
    private const uint FacilityItfPrefix = 0x80040000u;

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int ErrorInfoGetGuidDelegate(IntPtr thisPtr, out Guid pGuid);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int ErrorInfoGetBstrDelegate(IntPtr thisPtr, out IntPtr pBstr);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int ErrorInfoGetHelpContextDelegate(IntPtr thisPtr, out uint pdwHelpContext);

    /// <summary>
    /// Reads and clears the calling thread's current COM <c>IErrorInfo</c> object via <c>oleaut32!GetErrorInfo</c>.
    /// </summary>
    /// <returns>
    /// A populated <see cref="ComErrorInfo"/> when an error object was set on the calling thread;
    /// otherwise <see langword="null"/> when <c>GetErrorInfo</c> returns <c>S_FALSE</c> or all fields are empty.
    /// </returns>
    public static ComErrorInfo? ConsumeCurrentErrorInfo()
    {
        var hr = NativeMethods.GetErrorInfo(0, out var errorInfoPtr);
        if (hr != NativeMethods.S_OK || errorInfoPtr == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            return ReadErrorInfoFromPointer(errorInfoPtr);
        }
        finally
        {
            _ = Marshal.Release(errorInfoPtr);
        }
    }

    /// <summary>
    /// Reads the calling thread's current COM <c>IErrorInfo</c> object and immediately restores it via
    /// <c>oleaut32!SetErrorInfo</c> so downstream callers still observe the thread error state.
    /// </summary>
    /// <returns>
    /// A populated <see cref="ComErrorInfo"/> when an error object is present on the calling thread;
    /// otherwise <see langword="null"/>.
    /// </returns>
    public static ComErrorInfo? PeekCurrentErrorInfo()
    {
        var hr = NativeMethods.GetErrorInfo(0, out var errorInfoPtr);
        if (hr != NativeMethods.S_OK || errorInfoPtr == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            var snapshot = ReadErrorInfoFromPointer(errorInfoPtr);
            _ = NativeMethods.SetErrorInfo(0, errorInfoPtr);
            return snapshot;
        }
        finally
        {
            _ = Marshal.Release(errorInfoPtr);
        }
    }

    /// <summary>
    /// Creates an enriched <see cref="COMException"/> for a negative <paramref name="hr"/>, attaching any
    /// available thread-local <c>IErrorInfo</c> <c>Source</c>, <c>Description</c>, and <c>HelpFile</c> metadata.
    /// </summary>
    /// <param name="hr">The negative COM <c>HRESULT</c> failure code.</param>
    /// <param name="consumeErrorInfo">
    /// When <see langword="true"/> (default), consumes the thread's <c>IErrorInfo</c> slot; when <see langword="false"/>,
    /// peeks and restores the <c>IErrorInfo</c> slot on the calling thread.
    /// </param>
    /// <returns>A <see cref="COMException"/> carrying <paramref name="hr"/> and any captured <c>IErrorInfo</c> details.</returns>
    public static Exception CreateExceptionForHr(int hr, bool consumeErrorInfo = true)
    {
        if (hr >= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(hr), "HRESULT must be a negative failure code.");
        }

        var info = consumeErrorInfo ? ConsumeCurrentErrorInfo() : PeekCurrentErrorInfo();
        if (info is not null)
        {
            var summary = info.ToSummaryString();
            var message = string.IsNullOrWhiteSpace(summary)
                ? string.Format(CultureInfo.InvariantCulture, "COM error 0x{0:X8}.", unchecked((uint)hr))
                : string.Format(CultureInfo.InvariantCulture, "{0} (HRESULT: 0x{1:X8})", summary, unchecked((uint)hr));

            var enriched = new COMException(message, hr);
            if (info.Source is not null && !string.IsNullOrWhiteSpace(info.Source))
            {
                enriched.Source = info.Source;
            }

            if (info.HelpFile is not null && !string.IsNullOrWhiteSpace(info.HelpFile))
            {
                enriched.HelpLink = info.HelpContext != 0
                    ? string.Format(CultureInfo.InvariantCulture, "{0}#{1}", info.HelpFile, info.HelpContext)
                    : info.HelpFile;
            }

            return enriched;
        }

        var baseException = Marshal.GetExceptionForHR(hr, new IntPtr(-1));
        if (baseException is COMException directComException)
        {
            return directComException;
        }

        var fallbackMessage = baseException?.Message
            ?? string.Format(CultureInfo.InvariantCulture, "COM error 0x{0:X8}.", unchecked((uint)hr));
        return new COMException(fallbackMessage, hr);
    }

    /// <summary>
    /// Attempts to unpack a Visual Basic 6 / ActiveX runtime or application error number from <paramref name="hr"/>,
    /// recognizing <c>FACILITY_CONTROL</c> (<c>0x800Axxxx</c>) and standard OLE Automation <c>DISP_E_*</c> equivalents.
    /// </summary>
    /// <param name="hr">The COM <c>HRESULT</c> to inspect.</param>
    /// <param name="errorNumber">Receives the unpacked Visual Basic error number when this method returns <see langword="true"/>.</param>
    /// <returns><see langword="true"/> if <paramref name="hr"/> maps to a Visual Basic error number; otherwise <see langword="false"/>.</returns>
    public static bool TryGetVisualBasicErrorNumber(int hr, out int errorNumber)
    {
        var raw = unchecked((uint)hr);
        if ((raw & FacilityMask) == FacilityControlPrefix)
        {
            var code = (int)(raw & 0xFFFFu);
            if (code > 0)
            {
                errorNumber = code;
                return true;
            }
        }

        switch (hr)
        {
            case NativeMethods.DISP_E_TYPEMISMATCH:
                errorNumber = 13;
                return true;
            case NativeMethods.DISP_E_OVERFLOW:
                errorNumber = 6;
                return true;
            case NativeMethods.DISP_E_BADINDEX:
                errorNumber = 9;
                return true;
            case NativeMethods.DISP_E_DIVBYZERO:
                errorNumber = 11;
                return true;
            case NativeMethods.E_OUTOFMEMORY:
                errorNumber = 7;
                return true;
            default:
                errorNumber = 0;
                return false;
        }
    }

    /// <summary>
    /// Attempts to unpack an interface-specific <c>FACILITY_ITF</c> (<c>0x8004xxxx</c>) error offset relative to
    /// <paramref name="baseCode"/> (defaulting to Visual Basic's <c>vbObjectError</c> = <c>0x80040000</c>).
    /// </summary>
    /// <param name="hr">The COM <c>HRESULT</c> to inspect.</param>
    /// <param name="errorOffset">Receives the non-negative offset above <paramref name="baseCode"/> when successful.</param>
    /// <param name="baseCode">Base <c>FACILITY_ITF</c> error code (for example <c>0x80040000</c> or <c>0x80040200</c>).</param>
    /// <returns><see langword="true"/> if <paramref name="hr"/> is in <c>FACILITY_ITF</c> at or above <paramref name="baseCode"/>; otherwise <see langword="false"/>.</returns>
    public static bool TryGetInterfaceErrorOffset(
        int hr,
        out int errorOffset,
        int baseCode = unchecked((int)0x80040000))
    {
        var rawBase = unchecked((uint)baseCode);
        if ((rawBase & FacilityMask) != FacilityItfPrefix)
        {
            throw new ArgumentOutOfRangeException(
                nameof(baseCode),
                "Base code must belong to FACILITY_ITF (0x8004xxxx).");
        }

        var rawHr = unchecked((uint)hr);
        if ((rawHr & FacilityMask) != FacilityItfPrefix || rawHr < rawBase)
        {
            errorOffset = 0;
            return false;
        }

        errorOffset = (int)(rawHr - rawBase);
        return true;
    }

    private static ComErrorInfo? ReadErrorInfoFromPointer(IntPtr errorInfoPtr)
    {
        var vtable = Marshal.ReadIntPtr(errorInfoPtr);
        if (vtable == IntPtr.Zero)
        {
            return null;
        }

        var interfaceId = ReadGuidSlot(errorInfoPtr, vtable, slotIndex: 3);
        var source = ReadBstrSlot(errorInfoPtr, vtable, slotIndex: 4);
        var description = ReadBstrSlot(errorInfoPtr, vtable, slotIndex: 5);
        var helpFile = ReadBstrSlot(errorInfoPtr, vtable, slotIndex: 6);
        var helpContext = ReadUInt32Slot(errorInfoPtr, vtable, slotIndex: 7);

        var isEmpty = interfaceId == Guid.Empty
            && string.IsNullOrWhiteSpace(source)
            && string.IsNullOrWhiteSpace(description)
            && string.IsNullOrWhiteSpace(helpFile)
            && helpContext == 0;
        return isEmpty ? null : new ComErrorInfo(interfaceId, source, description, helpFile, helpContext);
    }

    private static Guid ReadGuidSlot(IntPtr errorInfoPtr, IntPtr vtable, int slotIndex)
    {
        var fnPtr = Marshal.ReadIntPtr(vtable, slotIndex * IntPtr.Size);
        if (fnPtr == IntPtr.Zero)
        {
            return Guid.Empty;
        }

        var fn = Marshal.GetDelegateForFunctionPointer<ErrorInfoGetGuidDelegate>(fnPtr);
        return fn(errorInfoPtr, out var guid) >= 0 ? guid : Guid.Empty;
    }

    private static string? ReadBstrSlot(IntPtr errorInfoPtr, IntPtr vtable, int slotIndex)
    {
        var fnPtr = Marshal.ReadIntPtr(vtable, slotIndex * IntPtr.Size);
        if (fnPtr == IntPtr.Zero)
        {
            return null;
        }

        var fn = Marshal.GetDelegateForFunctionPointer<ErrorInfoGetBstrDelegate>(fnPtr);
        var bstr = IntPtr.Zero;
        try
        {
            var hr = fn(errorInfoPtr, out bstr);
            return hr >= 0 && bstr != IntPtr.Zero
                ? Marshal.PtrToStringBSTR(bstr)
                : null;
        }
        finally
        {
            if (bstr != IntPtr.Zero)
            {
                Marshal.FreeBSTR(bstr);
            }
        }
    }

    private static uint ReadUInt32Slot(IntPtr errorInfoPtr, IntPtr vtable, int slotIndex)
    {
        var fnPtr = Marshal.ReadIntPtr(vtable, slotIndex * IntPtr.Size);
        if (fnPtr == IntPtr.Zero)
        {
            return 0;
        }

        var fn = Marshal.GetDelegateForFunctionPointer<ErrorInfoGetHelpContextDelegate>(fnPtr);
        return fn(errorInfoPtr, out var value) >= 0 ? value : 0;
    }
}
#pragma warning restore SA1121
