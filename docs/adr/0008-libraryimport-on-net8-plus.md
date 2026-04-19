# ADR-0008: Use `LibraryImport` Source Generator on .NET 8+

- Status: Accepted
- Date: 2025-08-10
- Deciders: @AdaskoTheBeAsT
- Release: v2.0.0
- Related commit(s): `3f8670c`

## Context

On .NET 7+, `[LibraryImport]` replaces the runtime-generated P/Invoke stubs emitted by `[DllImport]` with compile-time source-generated ones. The benefits are measurable: no reflection/IL emit at startup, smaller working set on AOT-capable runtimes, and explicit control over marshalling. On .NET Framework and `netstandard2.0`, the source generator is unavailable and classic `[DllImport]` remains the only option.

## Decision

The internal `NativeMethods` declares two parallel sets of imports gated by `#if`:

- `#if !NET8_0_OR_GREATER` → classic `[DllImport]` with attribute-based marshalling (`CharSet`, `return: MarshalAs(...)`), used by `netstandard2.0` and every .NET Framework target.
- `#if NET8_0_OR_GREATER` → `[LibraryImport]` with explicit `EntryPoint` suffixes (`CreateActCtxW`, `PeekMessageW`, `DispatchMessageW`) and blittable-only marshalling.

The structural declarations (`MSG`, `POINT`, constants, and the `PumpPendingMessages` helper) are duplicated inside each branch. This is slightly less DRY than a shared partial class would be, but it keeps each branch self-contained and easy to audit.

## Consequences

- Positive: modern .NET consumers benefit from AOT compatibility and lower startup overhead at zero call-site cost.
- Positive: legacy consumers keep the well-understood `[DllImport]` semantics.
- Negative: two bodies of near-identical Win32 declarations must be kept in sync.
- Negative: `StyleCop`/`S101` complains about the nested class style; a suppression is applied at file scope.

## Alternatives Considered

- Only `[DllImport]`: rejected because it forfeits the source generator's advantages on modern .NET.
- Only `[LibraryImport]`: rejected because it is unavailable on netstandard2.0 and .NET Framework.
- Shared partial class with `[DllImport]`/`[LibraryImport]` on different methods: rejected because the overloads differ per TFM and duplication is already bounded to one file.

## References

- `src/AdaskoTheBeAsT.Interop.COM/NativeMethods.cs`
- [LibraryImport overview](https://learn.microsoft.com/en-us/dotnet/standard/native-interop/pinvoke-source-generation)
