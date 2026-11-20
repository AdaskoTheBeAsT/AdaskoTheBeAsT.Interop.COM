# ADR-0036: CoreCLR STA `OleMainThreadWndClass` Pre-Warming and Opt-In Per-Module Unbalanced `CoUninitialize` IAT Guard

- Status: Accepted
- Date: 2026-11-17
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0
- Extends [ADR-0031](0031-wx-enforcement-legacy-ole2-dispenser-manager-and-startup-prep.md)

## Context

Two native STA apartment lifecycle hazards arise when hosting legacy VB6, ATL, and C++ COM servers on `.NET 10` / `.NET 11` and `.NET Framework`:

1. **Lazy CoreCLR STA initialization:** On `.NET Framework`, starting a `Thread` configured with `SetApartmentState(ApartmentState.STA)` (or `[STAThread]`) calls Win32 `CoInitializeEx(NULL, COINIT_APARTMENTTHREADED)` eagerly before user code runs, creating the thread's `OXID` and hidden `OleMainThreadWndClass` message window immediately. On modern CoreCLR (`net10.0`, `net11.0`), `Thread.GetApartmentState()` reports `ApartmentState.STA`, yet native `CoInitializeEx` is deferred lazily until the CLR's first built-in COM interop operation. If the caller's first operation inside `Executor.Activate` / `Execute` / `Create` is a raw native P/Invoke (`LoadLibraryExW` + `DllGetClassObject` via `ComClassFactory`, `CoGetApartmentType`, or `msvbvm60.dll` thread-attach window hooks), the OS thread still returns `CO_E_NOTINITIALIZED` (`0x800401F0`) and lacks `OleMainThreadWndClass`.
2. **Unbalanced `CoUninitialize()` / `OleUninitialize()` inside defective vendor DLLs:** Certain legacy COM DLLs invoke an extra `CoUninitialize()` or `OleUninitialize()` during object teardown or internal helper cleanup without a matching `CoInitialize[Ex]` / `OleInitialize`. On a managed STA thread whose native apartment reference count is `1`, that extra call tears down the thread's STA apartment prematurely (dropping the thread into `CO_E_NOTINITIALIZED` or the process-wide Implicit MTA), disconnecting all live STA RCWs and message-window marshaling on that worker thread.

## Decision

1. **Pre-warm native STA initialization on CoreCLR via `NanoServerOptions.EnsureCurrentStaThreadInitialized()`:**
   - Expose public idempotent `NanoServerOptions.EnsureCurrentStaThreadInitialized()` and invoke it automatically at the entry of `Executor.Activate`, `Executor.Execute`, and `Executor.Create` whenever `Thread.CurrentThread.GetApartmentState() == ApartmentState.STA`.
   - Guard with a `[ThreadStatic]` boolean fast path; on first observation per STA thread, inspect `combase!CoGetApartmentType`. Only when it returns `CO_E_NOTINITIALIZED` (`0x800401F0`), force CLR/OS native STA initialization while balancing any temporary native `CoInitializeEx(0, COINIT_APARTMENTTHREADED)` reference via `CoUninitialize()` so the CLR remains the sole owner of the thread's base STA reference count.
2. **Provide opt-in per-module `ComModuleQuirks.SuppressUninitializeImports` / `RestoreUninitializeImports`:**
   - Allocate a single static-lifetime W^X native code page (`VirtualAlloc` `PAGE_READWRITE` -> write `0xC3` near `ret`, which is ABI-valid for zero-argument `void __stdcall` on x86 and x64 -> `FlushInstructionCache` -> `VirtualProtect` `PAGE_EXECUTE_READ`, never RWX).
   - Patch standard (`.idata`) and delay-load (`IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT`) IAT entries for `CoUninitialize` and `OleUninitialize` **only** inside the caller-specified `HMODULE`, temporarily transitioning IAT data pages to `PAGE_READWRITE` and restoring original protections immediately.
   - Fail closed (`InvalidOperationException`) if the caller passes a protected OS, COM, or CLR module (`combase.dll`, `ole32.dll`, `oleaut32.dll`, `kernel32.dll`, `kernelbase.dll`, `ntdll.dll`, `rpcrt4.dll`, `user32.dll`, `clr.dll`, `coreclr.dll`, `mscorwks.dll`, or any binary under `%SystemRoot%\System32` or `%SystemRoot%\SysWOW64`), and automatically purge saved slot tracking in `NanoServerOle2Shim.OnModuleUnloaded` when a leased worker DLL is unloaded.

## Consequences

- Positive: Fresh `ApartmentState.STA` threads on `net10.0` and `net11.0` behave identically to `.NET Framework` from the first instruction inside `Executor` scopes or after calling `NanoServerOptions.EnsureCurrentStaThreadInitialized()`, without artificially inflating the STA reference count.
- Positive: Hosts dealing with a known defective third-party DLL can quarantine its stray `CoUninitialize` / `OleUninitialize` imports per `HMODULE` while keeping CLR, OS, and well-behaved COM modules completely untouched and strictly W^X-compliant.
- Negative: Suppressing `CoUninitialize` / `OleUninitialize` on a module that legitimately initializes and uninitializes its own private background threads would leak COM initialization on those private native threads; `ComModuleQuirks` therefore remains strictly opt-in per `HMODULE`.

## Alternatives Considered

- Incrementing the STA thread's `CoInitializeEx` reference count permanently by `+N` ("padding" the refcount) to absorb stray `CoUninitialize` calls: rejected because unbounded worker reuse eventually exhausts any fixed pad `N`, and leaked `CoInitializeEx` counts prevent clean apartment teardown when the managed STA thread exits.
- Hooking `combase!CoUninitialize` process-wide: rejected because global API detours violate W^X / CFG policies on system DLLs and interfere with legitimate CLR and OS thread teardown.

## References

- `src/AdaskoTheBeAsT.Interop.COM/NanoServerOptions.cs`
- `src/AdaskoTheBeAsT.Interop.COM/Executor.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComModuleQuirks.cs`
- `src/AdaskoTheBeAsT.Interop.COM/NanoServerOle2Shim.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/StaPreWarmingLocaleQuirksDiagnosticsAndManifestRewriterTest.cs`
