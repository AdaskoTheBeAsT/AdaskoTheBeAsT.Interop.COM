# ADR-0037: Thread-Affine Win32 Automation `LCID` and Managed Culture Scope (`ComAutomationLocaleScope`)

- Status: Accepted
- Date: 2026-11-17
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0

## Context

OLE Automation conversion routines inside `oleaut32.dll` (`VarR8FromStr`, `VarBstrFromR8`, `VarCyFromStr`, `VarDateFromStr`, `VariantChangeType`, and default `IDispatch::Invoke` coercion when callers pass `LOCALE_USER_DEFAULT`) read the **Win32 kernel thread locale** (`kernel32!GetThreadLocale` / `SetThreadLocale`), not `CultureInfo.CurrentCulture`.

On modern .NET (`net10.0`, `net11.0`), setting `CultureInfo.CurrentCulture = CultureInfo.InvariantCulture` updates managed thread-static state only and never calls Win32 `SetThreadLocale`. Consequently, on Polish (`0x0415`), German (`0x0407`), or French (`0x040C`) Windows hosts where the decimal separator is `,`, native VB6/Delphi/ATL COM components parsing `"1234.56"` via `VarR8FromStr(..., GetThreadLocale(), ...)` fail with `DISP_E_TYPEMISMATCH` (`0x80020005`) or silently misparse scale by `100x`.

Conversely, on `.NET Framework` (`net472`, `net48`, `net481`), assigning `Thread.CurrentThread.CurrentCulture` inside `clr.dll` *does* internally invoke `SetThreadLocale(culture.LCID)`—meaning that if a caller calls `SetThreadLocale(0x0409)` (`en-US`, required by Excel/VBA/VB6 Automation engines that reject `LOCALE_INVARIANT` `0x007F`) and *then* assigns `CultureInfo.CurrentCulture = CultureInfo.InvariantCulture`, `.NET Framework` silently overwrites the Win32 `LCID` from `0x0409` back to `0x007F`!

## Decision

Introduce sealed, thread-affine `ComAutomationLocaleScope : IDisposable`:

1. **Synchronize managed culture and Win32 `LCID` in the exact runtime-safe order:**
   - Expose `Enter(int threadLocaleId = UsEnglishLocaleId, CultureInfo? managedCulture = null, CultureInfo? managedUiCulture = null)` (defaulting `threadLocaleId` to `0x0409` `en-US` and `managedCulture` to `CultureInfo.InvariantCulture`, while keeping `CurrentUICulture` unchanged unless `managedUiCulture` is passed) and `EnterInvariant()` (`0x007F` + `CultureInfo.InvariantCulture`).
   - Capture prior `GetThreadLocale()`, `CultureInfo.CurrentCulture`, and `CultureInfo.CurrentUICulture`, plus both `Environment.CurrentManagedThreadId` and Win32 `GetCurrentThreadId()`.
   - Apply `CultureInfo.CurrentCulture` and `CultureInfo.CurrentUICulture` **before** calling `kernel32!SetThreadLocale(threadLocaleId)` so `.NET Framework`'s internal `SetThreadLocale(culture.LCID)` side effect never overwrites the explicitly requested Win32 `LCID`. If `SetThreadLocale` returns `false`, roll back the managed culture properties immediately and throw `Win32Exception`.
2. **Enforce strict thread affinity on `Dispose()`:**
   - Validate that `Dispose()` runs on the exact creating managed and Win32 OS thread (`InvalidOperationException` on cross-thread disposal, preventing accidental `async`/`await` continuation locale corruption on thread-pool threads).
   - Restore managed cultures first and Win32 `SetThreadLocale(_previousThreadLocaleId)` last so restoration is likewise immune to `.NET Framework`'s `CurrentCulture` side effect.

## Consequences

- Positive: VB6, Delphi, and Office/OLE Automation calls behave identically across Polish, German, US-English, and container hosts on all five TFMs (`net11.0`, `net10.0`, `net481`, `net48`, `net472`).
- Positive: Calling `ComAutomationLocaleScope.Enter()` gives `0x0409` (`en-US`) to native `oleaut32` / VB6 while giving `CultureInfo.InvariantCulture` to managed formatting inside the `using` block across both CoreCLR and `.NET Framework`.
- Negative: `ComAutomationLocaleScope` is strictly synchronous and thread-affine; callers must dispose it before yielding across `await` boundaries.

## Alternatives Considered

- Calling `SetThreadLocale` before setting `CultureInfo.CurrentCulture`: rejected because `.NET Framework` (`clr.dll`) internally calls `SetThreadLocale(culture.LCID)` during `CurrentCulture` assignment, clobbering `0x0409` with `0x007F` when `managedCulture` is `InvariantCulture`.
- Implementing `ComAutomationLocaleScope` as a `struct`: rejected because value-type copies risk double-restoration and prevent deterministic single-disposal state tracking.

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComAutomationLocaleScope.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/StaPreWarmingLocaleQuirksDiagnosticsAndManifestRewriterTest.cs`
