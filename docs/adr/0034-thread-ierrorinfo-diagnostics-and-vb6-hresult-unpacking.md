# ADR-0034: Thread `IErrorInfo` Diagnostics and VB6 / `FACILITY_ITF` `HRESULT` Unpacking

- Status: Accepted
- Date: 2026-11-17
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0

## Context

Visual Basic 6 (`msvbvm60.dll` `Err.Raise`), ATL (`AtlReportError`), and Delphi (`SafeCallError` / `CreateErrorInfo`) COM servers report rich domain errors by populating the calling thread's logical error slot via `oleaut32!SetErrorInfo(0, pErrorInfo)` before returning a failure `HRESULT`—typically `FACILITY_CONTROL` (`0x800A0000 | vbErrorNumber`, such as `0x800A000D` Type Mismatch or custom `vbObjectError + n`) or `FACILITY_ITF` (`0x80040200 + offset`).

When callers invoke raw vtable methods (`[PreserveSig]`), `DllGetClassObject` / `IClassFactory::CreateInstance`, or inspect caught `COMException` instances across target frameworks:

- `oleaut32!GetErrorInfo` is destructive by Win32 contract (clearing the thread error slot on read), so diagnostic logging before `Marshal.GetExceptionForHR` accidentally strips the description unless the `IErrorInfo*` is restored.
- Standard `Marshal.GetExceptionForHR` populates `Exception.Message`, `Source`, and `HelpLink` inconsistently across `.NET Framework` and modern `.NET` when `Description` is present or absent, and discards the `IErrorInfo::GetGUID` interface identifier and `GetHelpContext` numeric ID unless unpacked explicitly.
- Extracting VB6 error numbers from `0x800Axxxx` or interface error offsets from `0x80040200..0x8004FFFF` requires boilerplate bitmask manipulation in every consumer.

## Decision

Introduce immutable `ComErrorInfo` and static `ComErrorDiagnostics`:

1. **Destructive and non-destructive thread `IErrorInfo` readers:**
   - `ConsumeCurrentErrorInfo()` calls `oleaut32!GetErrorInfo(0, out var errorInfoPtr)` and reads `IErrorInfo` vtable slots `3`–`7` (`GetGUID`, `GetSource`, `GetDescription`, `GetHelpFile`, `GetHelpContext`), converting and freeing every `BSTR` via `Marshal.PtrToStringBSTR` + `Marshal.FreeBSTR` and releasing `errorInfoPtr` in `finally`.
   - `PeekCurrentErrorInfo()` reads the same `ComErrorInfo` snapshot and immediately restores `errorInfoPtr` to the calling thread via `oleaut32!SetErrorInfo(0, errorInfoPtr)` before releasing its local reference so subsequent `GetExceptionForHR` or native callers still observe the error slot.
2. **Consistent `COMException` enrichment and `HRESULT` helpers:**
   - `CreateExceptionForHr(int hr, bool consumeErrorInfo = true)` returns a `COMException` (always preserving `ErrorCode == hr`), attaches any captured `ComErrorInfo` snapshot under `exception.Data["ComErrorInfo"]`, populates `Source` and `HelpLink` when available, and uses `ComErrorInfo.Description` as the exception message when non-empty.
   - `TryGetVisualBasicErrorNumber(int hr, out int errorNumber)` unpacks negative `FACILITY_CONTROL` (`facility == 10`) `HRESULT`s (`hr & 0xFFFF`), and `TryGetInterfaceErrorOffset(int hr, out int offset)` unpacks negative `FACILITY_ITF` (`facility == 4`) user codes at or above `0x0200` (`(hr & 0xFFFF) - 0x0200`).

## Consequences

- Positive: Hosts calling `[PreserveSig]` interfaces, `ComClassFactory`, or VB6/Delphi servers can inspect or consume full `IErrorInfo` metadata (`Description`, `Source`, `HelpFile`, `HelpContext`, `InterfaceId`) without leaking `BSTR`s or clobbering downstream exception translation.
- Positive: `ComClassFactory` uses `ComErrorDiagnostics.CreateExceptionForHr` automatically when `DllGetClassObject` or `IClassFactory::CreateInstance` fails.
- Negative: `IErrorInfo` is thread-local in Win32 COM (`GetErrorInfo` / `SetErrorInfo`); callers must invoke `ComErrorDiagnostics` on the same OS thread that executed the failing COM call before making unrelated COM calls that clear the thread error slot.

## Alternatives Considered

- Wrapping `System.Runtime.InteropServices.ComTypes.IErrorInfo` via an RCW: rejected because creating a managed RCW just to read five `BSTR`/`GUID`/`DWORD` slots allocates unnecessary GC-tracked interop wrappers and complicates deterministic release on failure paths.
- Only offering destructive `GetErrorInfo`: rejected because diagnostic interceptors frequently need `PeekCurrentErrorInfo()` without altering standard CLR exception marshaling.

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComErrorInfo.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComErrorDiagnostics.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/ComErrorDiagnosticsAndBidirectionalRegistryTest.cs`
