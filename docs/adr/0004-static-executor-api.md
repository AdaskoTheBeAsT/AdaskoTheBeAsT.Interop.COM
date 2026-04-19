# ADR-0004: Expose Functionality Through a Static `Executor` Class

- Status: Accepted — mockability concerns addressed by [ADR-0013](0013-iexecutor-abstraction.md)
- Date: 2023-12-09
- Deciders: @AdaskoTheBeAsT
- Release: v1.0.0
- Related commit(s): `fb360c4`

## Context

The library offers stateless operations: supply a DLL path, a manifest path, and a delegate — get a result back. There is no per-instance configuration, no long-lived state, and no credential/authentication concern. A discoverable, "one call" API is therefore desirable.

## Decision

The primary public surface is a `static class Executor` that exposes `Execute`, `Create`, and `Free`. The class is intentionally thread-affine but not process-global state; each call builds its own activation context stack on the current thread.

## Consequences

- Positive: the call-site is concise, idiomatic, and obvious (`Executor.Execute(...)`).
- Positive: no need to construct, dispose, or inject an object for trivial one-shot uses.
- Negative: static APIs are hard to substitute in unit tests; callers that wanted to fake COM activation had no seam. This limitation is addressed in ADR-0013 via the new `IComExecutor` abstraction without removing the static API.
- Negative: static types cannot implement interfaces; adding DI support later required a separate wrapper class (see ADR-0013).

## Alternatives Considered

- Instance class with an empty constructor: rejected for v1.0.0 because it adds ceremony without benefits for stateless operations.
- Extension methods on `ComPathDescriptor`: rejected because it would scatter the entry points across multiple receivers.

## References

- `src/AdaskoTheBeAsT.Interop.COM/Executor.cs`
- [ADR-0013](0013-iexecutor-abstraction.md)
