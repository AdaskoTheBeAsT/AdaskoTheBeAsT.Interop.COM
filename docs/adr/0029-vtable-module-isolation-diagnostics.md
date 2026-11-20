# ADR-0029: Vtable-to-`HMODULE` Worker Isolation Diagnostics (`ComModuleDiagnostics`)

- Status: Accepted
- Date: 2026-11-15
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0

## Context

When running isolated COM worker directories (`worker_01\server.dll`, `worker_02\server.dll`) on developer workstations or Windows Servers where a vendor COM DLL is also globally registered in `HKLM\Software\Classes\CLSID`, a subtle manifest mismatch or unactivated thread causes `CoCreateInstance` to silently load the globally registered singleton DLL instead of the worker's isolated copy, corrupting mutable `.data`/`.bss` state across concurrent sessions.

However, naively reading vtable slot `0` (`IUnknown::QueryInterface`) from `Marshal.GetIUnknownForObject(comObject)` and asserting equality against `expectedModuleHandle` produces false positives when:

- The RCW wraps a cross-apartment proxy/stub owned by `combase.dll`, `ole32.dll`, `oleaut32.dll`, or `rpcrt4.dll`.
- `IUnknown` is implemented by a shared language runtime (`msvbvm60.dll` in Visual Basic 6 ActiveX servers) or the .NET CLR (`clr.dll`, `coreclr.dll`, `mscorwks.dll`) while a concrete custom interface (`IID`) vtable resides in the vendor DLL (or also delegates through the runtime).

## Decision

Introduce `ComModuleDiagnostics`, `ComModuleInspectionResult`, and `ComModuleVerificationStatus` (`Match`, `Mismatch`, `Inconclusive`):

1. **Support both default `IUnknown` inspection and explicit `Guid? interfaceId` + `int methodSlotIndex` inspection:**
   - Query `Marshal.GetIUnknownForObject(comObject)` (and `Marshal.QueryInterface` when `interfaceId` is specified), always releasing acquired native `IntPtr` references in `finally`.
   - Read the function pointer at `methodSlotIndex` and resolve its owning module via `GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, fnPtr, out var hMod)` and `GetModuleFileNameW`.
2. **Classify shared runtime/proxy modules honestly as `Inconclusive`:**
   - When the resolved `HMODULE` differs from `expectedModuleHandle` and belongs to a known shared COM marshaler or language runtime (`combase.dll`, `ole32.dll`, `oleaut32.dll`, `rpcrt4.dll`, `msvbvm60.dll`, `clr.dll`, `coreclr.dll`, `mscorwks.dll`), `InspectInstanceModule` returns `ComModuleVerificationStatus.Inconclusive` with a descriptive diagnostic message rather than falsely claiming either isolation proof (`Match`) or wrong-worker loading (`Mismatch`).
   - When the vtable belongs to a concrete unexpected server binary (e.g., `C:\Program Files\Vendor\server.dll` or `worker_02\server.dll` instead of `worker_01\server.dll`), `InspectInstanceModule` returns `ComModuleVerificationStatus.Mismatch` and `VerifyInstanceModule` throws `InvalidOperationException` reporting both hex `HMODULE` addresses, resolved module paths, `IID`, and slot index.

## Consequences

- Positive: Isolated worker hosts can assert in integration tests or debug assertions that live COM instances resolved to the expected worker DLL rather than a global `HKLM` copy.
- Positive: Cross-apartment proxies and VB6 `msvbvm60.dll` `IUnknown` thunks are explicitly reported as `Inconclusive`, and callers can pass a direct native `IID` and method slot to inspect non-thunked vtable entries.
- Negative: Pure `IDispatch` / VB6 tear-off interfaces whose entire vtable resides inside `msvbvm60.dll` remain `Inconclusive` by vtable address alone and must rely on `scope.LoadedModuleHandles` or process module enumeration alongside `ComModuleDiagnostics`.

## Alternatives Considered

- Treating any `HMODULE != expectedModuleHandle` unconditionally as `Mismatch`: rejected because cross-apartment STA proxies (`combase.dll`) and VB6 `IUnknown` runtime stubs (`msvbvm60.dll`) legitimately own slot `0` even when the worker DLL is properly isolated.
- Inspecting only `IUnknown` slot `0`: rejected because VB6 and ATL tear-off objects frequently forward `IUnknown` to a shared base/runtime while custom native interfaces point directly into the server DLL.

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComModuleDiagnostics.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComModuleInspectionResult.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComModuleVerificationStatus.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/ComModuleDiagnosticsTest.cs`
