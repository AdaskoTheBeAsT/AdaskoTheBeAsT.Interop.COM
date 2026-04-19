# ADR-0007: Multi-Target `netstandard2.0`, Modern .NET, and .NET Framework

- Status: Accepted
- Date: 2025-08-10 (modern .NET), 2026-04-19 (.NET Framework 4.6.2-4.8.1)
- Deciders: @AdaskoTheBeAsT
- Release: v2.0.0 / v3.0.0
- Related commit(s): `3f8670c`, `5fa56fa`, v3.0.0 changes

## Context

Registration-free COM is predominantly consumed by legacy or semi-legacy applications — older WinForms/WPF apps, classic ASP.NET hosts, enterprise LOB software still on the .NET Framework, as well as modern .NET apps that need to keep calling existing COM components. Targeting only modern .NET would shut out a large fraction of realistic users; targeting only `netstandard2.0` would forfeit the newer interop source generators.

## Decision

The library multi-targets the broadest reasonable matrix:

- `netstandard2.0` — covers generic .NET Framework 4.6.1+ and modern .NET consumers who already set up a netstandard-compatible pipeline.
- `net462`, `net47`, `net471`, `net472`, `net48`, `net481` — explicit .NET Framework targets so consumers avoid the `netstandard2.0` binding-redirect pain and get correct `Serializable` exception patterns.
- `net8.0`, `net9.0`, `net10.0` — modern .NET with source-generated P/Invoke (see ADR-0008) and `[SupportedOSPlatform]` annotations (see ADR-0014).

Compile-time branching uses `#if NET8_0_OR_GREATER` for modern-only features, and `#if !NET8_0_OR_GREATER` (replacing the older `#if NETSTANDARD2_0`) for the classic code path, so a new .NET Framework target automatically picks up the right branch without extra `#if` noise.

## Consequences

- Positive: both legacy and modern consumers install the same NuGet package and get the best variant for their runtime.
- Positive: each framework flavour ships its idiomatic P/Invoke (source-generated on net8+, classic `DllImport` elsewhere) and the correct `ISerializable` constructor on exception types.
- Negative: 10 target frameworks multiply build time and increase the CI matrix.
- Negative: contributors must be aware of the `#if` branches and avoid using APIs that exist only on some TFMs without guarding them.

## Alternatives Considered

- `netstandard2.0` only: rejected because it forfeits `LibraryImport` and platform annotations.
- `net8.0+` only: rejected because it locks out the legacy consumers who actually need reg-free COM the most.
- A separate NuGet package per framework family: rejected because it fragments discoverability and versioning.

## References

- `src/AdaskoTheBeAsT.Interop.COM/AdaskoTheBeAsT.Interop.COM.csproj`
- `src/AdaskoTheBeAsT.Interop.COM/NativeMethods.cs`
