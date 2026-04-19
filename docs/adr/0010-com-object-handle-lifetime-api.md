# ADR-0010: Introduce `ComObjectHandle<T>` Lifetime API with Explicit Free

- Status: Accepted — extended by [ADR-0012](0012-com-object-handle-idisposable.md)
- Date: 2026-04-07
- Deciders: @AdaskoTheBeAsT
- Release: v2.1.0
- Related commit(s): `5fa56fa`

## Context

The original `Executor.Execute(..., Action)` model requires that the COM object lives only for the lifetime of the callback. That restriction is fine for one-shot calls but is unusable for scenarios where the same COM object must be reused across many invocations on the same STA thread (for example, a long-lived Office automation session, a rendering engine, or a scheduler that hands the object back to callers via `async`/`await`).

## Decision

Two new entry points, `Executor.Create<T>` and `Executor.Free<T>`, express the extended lifetime explicitly:

- `Create<T>` activates the requested contexts, runs a `Func<T>` factory inside them, and returns a `ComObjectHandle<T>` that holds the COM object together with the cookies and activation-context handles required to keep it alive.
- `Free<T>` releases the COM object (through `Marshal.FinalReleaseComObject`), deactivates the contexts in reverse order, releases the underlying activation contexts, and is idempotent.

A sibling result type `ComObjectCreationResult<T>` extends `Result` with a nullable `Value` so the caller gets the success flag, the exception (if any), and the handle in a single return. Handle internals (the `IntPtr` lists, `IsReleased`) are kept `internal`; the caller only ever touches `ComObject` and, starting in v3.0.0, the `Dispose`/`IsReleased` members (see ADR-0012).

## Consequences

- Positive: long-lived COM usage becomes expressible without leaking the activation-context detail into user code.
- Positive: the same `ComPathDescriptor` collection model from ADR-0009 applies uniformly to `Create`/`Free`.
- Negative: it introduces a real resource to manage. In v2.1.0 there was no `IDisposable` contract, so callers had to remember to call `Executor.Free`; missed calls leaked activation contexts. ADR-0012 closes that gap.
- Negative: the handle is thread-affine and must be released on the creating thread; the library cannot enforce this and only documents it.

## Alternatives Considered

- Re-entering `Execute` for each call: rejected because it tears down and rebuilds the activation context per operation, which is the wrong cost model for long-lived COM usage.
- A `using`-scoped type that releases on dispose only (no `Free`): deferred to ADR-0012; v2.1.0 introduced the primitive first.

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComObjectHandle.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComObjectCreationResult.cs`
- `src/AdaskoTheBeAsT.Interop.COM/Executor.cs`
