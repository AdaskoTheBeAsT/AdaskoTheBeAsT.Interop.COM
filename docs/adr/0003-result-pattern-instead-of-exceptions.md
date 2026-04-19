# ADR-0003: Expose a Result Pattern Instead of Throwing from the Public API

- Status: Accepted
- Date: 2023-12-09
- Deciders: @AdaskoTheBeAsT
- Release: v1.0.0
- Related commit(s): `fb360c4`

## Context

Activation-context operations can fail for a variety of reasons that are hard to predict up front: a malformed manifest, a CLSID mismatch, a process/architecture mismatch (see ADR-0005 / ActCtx size guard), COM factory failures, or user code raising exceptions from inside the callback. Throwing from every failure mode would force every caller to wrap `Execute`/`Create` calls in `try`/`catch` and would hide valuable diagnostic information behind stack unwinding.

## Decision

The public entry points (`Executor.Execute`, `Executor.Create`, `Executor.Free`) return a `Result` (or derived `ComObjectCreationResult<T>`) that carries both a `Success` flag and the captured `Exception`. Argument validation failures (null/empty paths, null delegates, empty descriptor collections) remain thrown as `ArgumentNullException` / `ArgumentException` because they indicate programmer error that cannot be acted upon at runtime.

## Consequences

- Positive: callers can branch on `result.Success` and inspect `result.Exception` without unwinding the stack.
- Positive: transient or expected failures are no longer indistinguishable from bugs in user code.
- Negative: breaks the expectations of developers who prefer exception-propagation style; they must explicitly re-throw when desired.
- Negative: the result type has mutable settable properties for ergonomic construction, which may surprise readers expecting a pure value object.

## Alternatives Considered

- Classic exception propagation: rejected for the reasons above.
- An `Outcome<TSuccess, TError>` ADT: rejected as unnecessarily heavy for a library with a tiny public surface; the `Result` + optional `Value` in `ComObjectCreationResult<T>` is enough.
- Returning `bool` + `out Exception`: rejected because it doesn't compose well with generics and forces out-params on every call site.

## References

- `src/AdaskoTheBeAsT.Interop.COM/Result.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComObjectCreationResult.cs`
