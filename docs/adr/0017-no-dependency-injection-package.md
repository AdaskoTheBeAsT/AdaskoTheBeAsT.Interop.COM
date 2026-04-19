# ADR-0017: Do Not Ship a Separate `DependencyInjection` Package (For Now)

- Status: Accepted
- Date: 2026-04-19
- Deciders: @AdaskoTheBeAsT
- Release: v3.0.0

## Context

Many .NET libraries ship a companion `*.Extensions.DependencyInjection` (or `*.DependencyInjection`) NuGet — EF Core, MediatR, Polly, Serilog, FluentValidation, and so on. The question came up whether `AdaskoTheBeAsT.Interop.COM` should follow the same convention and ship a sibling `AdaskoTheBeAsT.Interop.COM.DependencyInjection` with:

- an `IServiceCollection.AddComExecutor(...)` extension method, and
- an `IOptions<ComExecutorOptions>`-style configuration class.

Arguments in favour:

- Idiomatic .NET wiring; consumers know where to look.
- Keeps the core package free of `Microsoft.Extensions.*` transitive dependencies (important on `netstandard2.0` / .NET Framework targets, see ADR-0007).
- A natural future home for cross-cutting concerns — logger integration, metrics, a named-descriptor registry, telemetry hooks.

Arguments against:

- `ComExecutor` is **stateless**. `services.AddSingleton<IComExecutor, ComExecutor>();` is already the entire wiring, so an `AddComExecutor()` helper buys only cosmetic sugar.
- `IOptions<ComExecutorOptions>` implies *real* options; we currently have none. Inventing them purely to justify the package is YAGNI.
- Consumers who need logging, metrics, or a named-descriptor catalogue can achieve that in a few lines via the **decorator pattern** on `IComExecutor`, documented in the README.
- An extra package doubles the release surface: its own version, CI pipeline, changelog, and update cadence.

## Decision

We will **not** ship a separate `AdaskoTheBeAsT.Interop.COM.DependencyInjection` package in v3.0.0.

The canonical DI wiring remains the one-line registration:

```csharp
services.AddSingleton<IComExecutor, ComExecutor>();
```

Keyed registrations (`AddKeyedSingleton` on .NET 8+), logging, and metrics integration are documented in the README as **decorator-pattern recipes** on top of `IComExecutor`. Because `IComExecutor` is a small, stable interface and `ComExecutor` is a pure pass-through, consumer-side decorators are sufficient and carry no intrinsic library-specific knowledge.

## Consequences

- Positive: the main package stays dependency-free (no `Microsoft.Extensions.*` transitive surface), which matters on `net462`+`netstandard2.0` targets that many consumers still deploy to.
- Positive: there is no separate package to version, release, and document; the release surface stays minimal.
- Positive: consumers who want tailored DI behaviour are not constrained by our opinionated helper — they can decorate, wrap, or keyed-register as they see fit.
- Negative: every consumer who adds logging or metrics writes their own decorator. If enough of them do, the cost of maintaining a shared implementation becomes worth the split and this ADR should be revisited.
- Negative: the "decorator shape" is documentation-only; a minor typo in the README examples lands on the user. If we see repeated confusion we should convert the README snippets into sample projects under `samples/`.

## Triggers That Would Revisit This Decision

Any of the following should reopen the question:

1. Real configuration surface emerges — e.g. a named `ComPathDescriptor` registry, a default base directory, or cross-cutting timeouts that cannot reasonably live on individual call-sites.
2. An upstream .NET BCL change makes decorator wiring awkward (e.g. open-generic DI semantics we want to expose cleanly).
3. Multiple consumers file issues asking for logging, metrics, or keyed-by-descriptor registration and the README recipes prove insufficient.
4. We decide to standardise a telemetry story across all `AdaskoTheBeAsT.Interop.*` packages (Threading, COM, …), at which point a shared `*.DependencyInjection` package makes sense as a single integration point.

When any of the above triggers, the follow-up ADR should specify which scope level we're shipping — minimal / descriptor-catalog / full observability — and keep the scope choices I sketched during the v3.0.0 investigation.

## Alternatives Considered

- **Ship a minimal DI package** (`AddComExecutor()` only, empty options) — rejected because it is pure cosmetic sugar over a one-liner.
- **Ship a descriptor-catalogue DI package** — rejected for v3.0.0 because no consumer has asked for named descriptors yet; adding the surface speculatively locks us into an API that may not match real needs.
- **Ship a full observability DI package** (logging + metrics wrappers) — rejected because the decorator pattern on `IComExecutor` is idiomatic and flexible; baking a specific logging / `Meter` schema into the library would prematurely constrain consumer observability stacks.
- **Inline DI helpers into the main package with an optional `Microsoft.Extensions.DependencyInjection.Abstractions` dependency** — rejected because it pollutes the transitive graph for consumers who don't need DI.

## References

- [ADR-0013](0013-icomexecutor-abstraction.md) — `IComExecutor` abstraction
- [ADR-0007](0007-multi-target-frameworks.md) — why the main library stays dependency-light
- README "Dependency Injection" section — canonical wiring, keyed services, logging & metrics decorator recipes
