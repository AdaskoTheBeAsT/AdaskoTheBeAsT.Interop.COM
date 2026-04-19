# ADR-0013: Introduce `IComExecutor` Abstraction and `ComExecutor` Implementation

- Status: Accepted
- Date: 2026-04-19
- Deciders: @AdaskoTheBeAsT
- Release: v3.0.0
- Related commit(s): v3.0.0 changes

## Context

Since v1.0.0 the public surface has been a `static class Executor` (ADR-0004). That is great for discoverability but prevents consumers from substituting a fake in unit tests: you cannot assign a static to a field, you cannot inject it through a constructor, and `Moq`/`NSubstitute` cannot intercept static calls. Users who wanted to unit-test code paths that happen to call COM had no choice but to wrap `Executor` themselves, which meant every consumer reinvented the same tiny abstraction.

## Decision

Two new types are added without changing the static API:

- `IComExecutor` — an interface that mirrors every public method of `Executor` (`Execute`, `Create`, `Free`, including the single-descriptor and collection overloads). It is annotated with `[SupportedOSPlatform("windows")]` on .NET 8+ (see ADR-0014). The name is deliberately COM-specific (not the bare `IExecutor`) to avoid clashing with generic "executor" abstractions that consumers may already have in their codebases.
- `ComExecutor` — a `sealed class` that implements `IComExecutor` by delegating each call to the corresponding static method on `Executor`. It carries no state, so consumers can register it as a singleton in their DI container.

Starting in v3.0.0 **`IComExecutor` + `ComExecutor` is the recommended API for all new code**. The static `Executor` class is retained unchanged for source-level compatibility with v2.x callers, but every example in the README, every XML doc summary, and every piece of internal guidance points new consumers at the interface. `ComExecutor` is a thin adapter — there is no duplicated behaviour, only a pass-through — so the behaviour of both APIs is guaranteed identical.

## Consequences

- Positive: consumers can depend on `IComExecutor`, register `ComExecutor` as its implementation, and swap in a fake in tests.
- Positive: no existing caller is forced to migrate; `Executor.Execute` and `Executor.Free` still work as before.
- Positive: by establishing `IComExecutor` as the preferred API, downstream codebases grow in a testable direction by default.
- Positive: the COM-specific name keeps the library's abstraction distinct from unrelated `IExecutor` types that might already live in consumer codebases.
- Negative: the public surface grows by two types that must be maintained alongside `Executor`. Any new method added to `Executor` must also be added to `IComExecutor` and `ComExecutor`.
- Negative: some callers may read the presence of an interface as a hint that multiple implementations exist; documentation is clear that `ComExecutor` is the only one we ship.
- Negative: two coexisting APIs with identical semantics can confuse newcomers. The README and XML docs consistently mark the static `Executor` as "legacy, v2.x compatible" and `IComExecutor`/`ComExecutor` as "recommended" to minimise that confusion.

## Alternatives Considered

- Converting `Executor` into an instance class: rejected because it silently breaks every v1/v2 consumer that calls `Executor.Execute(...)`.
- Ship only `ComExecutor` and remove the static API: rejected for the same reason.
- Ship only an interface and let consumers write their own implementation: rejected because 95% of consumers want the out-of-the-box wiring.

## References

- `src/AdaskoTheBeAsT.Interop.COM/IComExecutor.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComExecutor.cs`
- [ADR-0004](0004-static-executor-api.md)
