# ADR-0014: Annotate Public Surface with `[SupportedOSPlatform("windows")]`

- Status: Accepted
- Date: 2026-04-19
- Deciders: @AdaskoTheBeAsT
- Release: v3.0.0
- Related commit(s): v3.0.0 changes

## Context

The library depends on Win32 APIs (`kernel32.dll`, `user32.dll`) and on COM itself, so it cannot possibly run on non-Windows platforms. Until v3.0.0, however, the public surface carried no annotation that expressed this constraint. Consumers compiling against `netstandard2.0` or `net8.0` (without `TargetPlatformIdentifier`) received no CA1416 warning from their own code, and a single `#pragma warning disable CA1416` inside `Executor.Free` suppressed the one place where `Marshal.FinalReleaseComObject` tripped the analyzer.

## Decision

On every TFM that supports it (`NET8_0_OR_GREATER`), the following types are now decorated with `[SupportedOSPlatform("windows")]`:

- `Executor`
- `ComObjectHandle<T>`
- `IComExecutor`
- `ComExecutor`

On older TFMs (`netstandard2.0`, `net462`-`net481`) the attribute is unavailable at the BCL level and the annotation is skipped via `#if NET8_0_OR_GREATER`. The blanket `#pragma warning disable CA1416` around `Marshal.FinalReleaseComObject` is removed — the attribute on `Executor` makes the call-site correctly platform-scoped.

The test assembly (`AdaskoTheBeAsT.Interop.COM.Test`) receives a matching assembly-level `[assembly: SupportedOSPlatform("windows")]` declaration under the same `#if` so that `CA1416` does not flag every test method.

## Consequences

- Positive: analyzers now correctly reason about the library's platform requirements. Consumers targeting cross-platform projects get a visible warning when they accidentally reference the library.
- Positive: the inline `CA1416` suppression is gone, which is more honest than suppressing a true analyzer signal.
- Negative: consumers who previously consumed the library from non-Windows-annotated projects may see new CA1416 warnings when they upgrade. This is the right signal, but it is a behavioural change that warrants a major version bump (v3.0.0).
- Negative: every future type that touches the public surface must remember to add the attribute. This is a maintenance cost but a small one.

## Alternatives Considered

- Annotating only the handful of methods that call Win32: rejected because the entire library is Windows-only; a class-level attribute is more honest than a method-level patchwork.
- Keeping the `#pragma warning disable CA1416`: rejected because it hides a real analyzer signal instead of expressing the constraint.
- Adding `SupportedOSPlatform` to the `.csproj` via `TargetPlatformIdentifier=Windows`: rejected because doing so on a multi-TFM library constrains more than necessary, and because the attribute approach composes better with `netstandard2.0` consumers.

## References

- `src/AdaskoTheBeAsT.Interop.COM/Executor.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComObjectHandle.cs`
- `src/AdaskoTheBeAsT.Interop.COM/IComExecutor.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComExecutor.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/AssemblyInfo.cs`
- [CA1416 (Microsoft Learn)](https://learn.microsoft.com/en-us/dotnet/fundamentals/code-analysis/quality-rules/ca1416)
