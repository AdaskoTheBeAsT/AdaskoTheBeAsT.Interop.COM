# ADR-0005: Encapsulate Win32 P/Invoke in an Internal `NativeMethods`

- Status: Accepted
- Date: 2023-12-09
- Deciders: @AdaskoTheBeAsT
- Release: v1.0.0
- Related commit(s): `fb360c4`

## Context

The library calls into `kernel32.dll` (`CreateActCtx`, `ActivateActCtx`, `DeactivateActCtx`, `ReleaseActCtx`) and `user32.dll` (`PeekMessage`, `TranslateMessage`, `DispatchMessage`). Exposing these at all would lock the library into their signatures forever, and calling them from scattered call-sites would duplicate the `[DllImport]` attributes and make the surface harder to audit.

Additionally, the native `ACTCTX` structure is size-sensitive: a mismatch between the managed layout and the process architecture (`x86` vs `x64`) yields hard-to-diagnose failures at `CreateActCtx` time.

## Decision

All Win32 imports live in a single `internal static class NativeMethods`. The class is decorated with `[SuppressUnmanagedCodeSecurity]` to elide runtime stack walks on platforms where it still matters. The layout-sensitive `ACTCTX` struct is declared in its own `internal struct ActCtx` and its `cbSize` is validated against a hard-coded architecture-specific expected value (`0x20` on x86, `0x38` on x64) so size mismatches are detected immediately via `ActCtxWrongSizeException`.

## Consequences

- Positive: every native call is audited in one place; changing a calling convention or signature touches exactly one file.
- Positive: the `cbSize` guard converts a class of silent runtime failures into an explicit, typed exception.
- Negative: the class grows slightly as frameworks evolve (see ADR-0008 regarding `LibraryImport`).
- Negative: `StyleCop` flags Hungarian-style field names in `ActCtx`; those suppressions are intentional because the names must mirror Win32 documentation.

## Alternatives Considered

- Inlining `[DllImport]` calls at each call site: rejected because it spreads a single Win32 API surface across multiple managed files.
- Using `Marshal.GetDelegateForFunctionPointer` + `LoadLibrary`: rejected as unnecessarily dynamic; the DLLs we consume are guaranteed to be present on any supported Windows version.

## References

- `src/AdaskoTheBeAsT.Interop.COM/NativeMethods.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ActCtx.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ActCtxWrongSizeException.cs`
