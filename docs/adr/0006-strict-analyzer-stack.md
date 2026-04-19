# ADR-0006: Enforce a Strict Analyzer Stack with Warnings as Errors

- Status: Accepted
- Date: 2023-12-09 (initial), reaffirmed 2025-01-19 (`f78fcb2`, "fix code smells") and 2026-04-07 (`5fa56fa`)
- Deciders: @AdaskoTheBeAsT
- Release: v1.0.0+
- Related commit(s): `fb360c4`, `f78fcb2`, `5fa56fa`

## Context

This library sits at the intersection of COM interop, unsafe code, P/Invoke, and threading — four of the most bug-prone areas of the .NET platform. The cost of shipping a subtle defect (e.g. a leaked activation context, a wrong apartment assumption, a missing `finally`) is high because such bugs tend to manifest only under load or on customer machines. We want tooling, not memory, to catch as many of those as possible before code review.

## Decision

The `Directory.Build.props` file turns on an intentionally broad analyzer stack that covers style, security, concurrency, reflection, async, disposables, and Sonar-style code smells. The project is compiled with `TreatWarningsAsErrors=true`, `Nullable=enable`, `EnforceCodeStyleInBuild=true`, and `AnalysisLevel=latest`. SARIF output is emitted to `.sarif/` for code-scanning tools.

The analyzer set includes:

- `AdaskoTheBeAsT.AsyncFixer` / `Asyncify`
- `AdaskoTheBeAsT.CodeCracker.CSharp`
- `AdaskoTheBeAsT.ConcurrencyLab.ParallelChecker`
- `AdaskoTheBeAsT.IDisposableAnalyzers`
- `AdaskoTheBeAsT.Puma.Security.Rules.2022`
- `AdaskoTheBeAsT.ReflectionAnalyzer`
- `AdaskoTheBeAsT.SecurityCodeScan.VS2022`
- `Meziantou.Analyzer`
- `Microsoft.CodeAnalysis.NetAnalyzers`
- `Microsoft.VisualStudio.Threading.Analyzers`
- `Roslynator.Analyzers` / `Roslynator.CodeAnalysis.Analyzers` / `Roslynator.Formatting.Analyzers`
- `SonarAnalyzer.CSharp`
- `StyleCop.Analyzers`

Suppressions are allowed only at the narrowest scope possible, with a comment explaining why, and only for rules that conflict with the native-interop reality (e.g. `MA0055` around the intentional diagnostic finalizer on `ComObjectHandle<T>`).

## Consequences

- Positive: a broad class of bugs (missing `ConfigureAwait`, swallowed exceptions, leaked disposables, insecure patterns) are caught during `dotnet build`.
- Positive: periodic commits named "update libs" and "fix code smells" have a clear mandate and cadence.
- Negative: onboarding cost is non-trivial: contributors must accept the analyzer policy rather than tune it down.
- Negative: occasionally an analyzer rule is too strict for genuinely correct interop code; such cases are handled by targeted `#pragma warning` blocks.

## Alternatives Considered

- Shipping a single recommended ruleset only (e.g. `Microsoft.CodeAnalysis.NetAnalyzers`): rejected because it misses security and threading concerns.
- Treating warnings as warnings: rejected because on a small, interop-heavy library, any unchecked warning is very likely a real bug.

## References

- `Directory.Build.props`
- `AdaskoTheBeAsT.ruleset`
- `.sarif/`
