# ADR-0025: Update Supported Target Frameworks to .NET 11, .NET 10, and Supported .NET Framework

- Status: Accepted
- Date: 2026-11-15
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0
- Supersedes [ADR-0021](0021-consolidate-supported-target-frameworks.md)

## Context

With the upgrade of `global.json` to .NET SDK `11.0.100` and C# 14 (`<LangVersion>14.0</LangVersion>`) for the `v5.0.0` major release, maintaining six target frameworks (`net8.0`, `net9.0`, `net10.0`, `net472`, `net48`, `net481`) while adding `net11.0` would expand the build and native x86/x64 test matrix across out-of-scope modern runtimes:

- `.NET 9` was a Standard Term Support (STS) release superseded by `.NET 10` LTS and `.NET 11`.
- Modern .NET consumers upgrading to major version `v5.0.0` target `.NET 10` LTS or `.NET 11`, while enterprise desktop and Windows Server COM workloads continue to rely heavily on in-support `.NET Framework` (`4.7.2`, `4.8`, and `4.8.1`).
- Every shipped target framework in this repository must be built with the full analyzer suite (`TreatWarningsAsErrors=true`) and executed against native x86 and x64 C++ COM fixtures without shipping untested fallback binaries or taking external runtime NuGet dependencies.

## Decision

Consolidate `<TargetFrameworks>` in `AdaskoTheBeAsT.Interop.COM` and `AdaskoTheBeAsT.Interop.COM.Test` for `v5.0.0` to five concrete targets:

- Modern .NET: `net11.0`, `net10.0`
- .NET Framework: `net481`, `net48`, `net472`

Remove `net8.0` and `net9.0` package assets. Retain zero external runtime NuGet package dependencies across all five targets (keeping `#if NET8_0_OR_GREATER` guards for `LibraryImport`, `[SupportedOSPlatform("windows")]`, and `Span<T>` / `ReadOnlySpan<T>` overloads alongside `IReadOnlyList<T>` overloads available on `.NET Framework`).

## Consequences

- Positive: Every shipped target framework (`net11.0`, `net10.0`, `net481`, `net48`, `net472`) is directly compiled and tested against native C++ COM servers on both `x86` and `x64`.
- Positive: Zero external runtime NuGet dependencies on any target framework, avoiding transitive `System.Memory` binding redirects on `.NET Framework 4.7.2`–`4.8.1`.
- Negative: Applications still targeting `net8.0` or `net9.0` must retarget to `net10.0` or `net11.0` to adopt `v5.0.0`, or stay on `v4.2.0`.

## Alternatives Considered

- Retaining `net8.0` and `net9.0` alongside `net10.0` and `net11.0`: rejected for the `v5.0.0` major bump to keep the multi-architecture native COM build/test matrix focused on `.NET 10` LTS, `.NET 11`, and `.NET Framework 4.7.2`–`4.8.1`.
- Adding `System.Memory` as a NuGet dependency on `.NET Framework` to expose `Span<T>` everywhere: rejected to preserve the zero-dependency deployment footprint and avoid assembly binding-redirect issues in legacy COM hosts.

## References

- `global.json`
- `Directory.Build.props`
- `src/AdaskoTheBeAsT.Interop.COM/AdaskoTheBeAsT.Interop.COM.csproj`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/AdaskoTheBeAsT.Interop.COM.Test.csproj`
- `MIGRATION.md`
