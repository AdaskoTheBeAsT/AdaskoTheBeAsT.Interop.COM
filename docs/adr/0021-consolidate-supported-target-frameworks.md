# ADR-0021: Consolidate Supported Target Frameworks to Modern .NET and Supported .NET Framework

- Status: Accepted
- Date: 2026-09-10
- Deciders: @AdaskoTheBeAsT
- Release: v4.0.0
- Supersedes [ADR-0007](0007-multi-target-frameworks.md)
- Related commit(s): `1a4b9c8`

## Context

In v3.0.0 (ADR-0007), the library targeted ten frameworks (`netstandard2.0`, `net462`, `net47`, `net471`, `net472`, `net48`, `net481`, `net8.0`, `net9.0`, and `net10.0`). Maintaining the pre-`net472` .NET Framework targets and `netstandard2.0` introduced growing friction:

- Modern test tooling (`xunit.v3` and `Microsoft.Testing.Platform`) requires at least `.NET Framework 4.7.2` on classic Windows runtimes.
- Shipping explicit `.NET Framework 4.7.2+` and `.NET 8+` assets makes a separate `netstandard2.0` asset redundant for supported Windows runtimes while adding extra multi-target build and packaging overhead.
- `.NET Framework 4.6.2`, `4.7`, and `4.7.1` have minimal remaining usage compared to `4.7.2`, `4.8`, and `4.8.1`, which are the baseline in-support .NET Framework releases on modern Windows installations.

## Decision

Consolidate `<TargetFrameworks>` in both `AdaskoTheBeAsT.Interop.COM` and `AdaskoTheBeAsT.Interop.COM.Test` to six supported targets:

- Modern .NET: `net8.0`, `net9.0`, `net10.0`
- .NET Framework: `net472`, `net48`, `net481`

Remove `netstandard2.0`, `net462`, `net47`, and `net471` package assets starting in v4.0.0. Keep the `#if NET8_0_OR_GREATER` / `#if !NET8_0_OR_GREATER` split for source-generated `LibraryImport` vs. classic `DllImport` and binary serialization support.

## Consequences

- Positive: every shipped target framework is directly exercised by the `xunit.v3` + `Microsoft.Testing.Platform` test suite, including native COM lifetime and message-pumping tests.
- Positive: faster multi-target builds and a smaller NuGet package without untested fallback assets.
- Negative: consumers still targeting `net462`–`net471` or relying on `netstandard2.0` must upgrade to at least `net472` / `net8.0` or stay on v3.0.0.

## Alternatives Considered

- Keeping `netstandard2.0` in `src/` while testing only `net472+`: rejected because shipping an untested binary asset in a native COM interop library undermines confidence.
- Dropping `.NET Framework` entirely: rejected because registration-free COM remains heavily used in `.NET Framework 4.7.2`–`4.8.1` desktop and enterprise applications.

## References

- `src/AdaskoTheBeAsT.Interop.COM/AdaskoTheBeAsT.Interop.COM.csproj`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/AdaskoTheBeAsT.Interop.COM.Test.csproj`
- `MIGRATION.md`
