# ADR-0012: Make `ComObjectHandle<T>` Implement `IDisposable`

- Status: Accepted
- Date: 2026-04-19
- Deciders: @AdaskoTheBeAsT
- Release: v3.0.0
- Related commit(s): v3.0.0 changes

## Context

In v2.1.0 the lifetime API (ADR-0010) relied on the caller calling `Executor.Free(handle)` explicitly. That works when the code path is simple and easy to follow, but the modern .NET idiom for deterministic cleanup is `using`/`using var`, and C# developers reflexively reach for it on any type that looks resource-owning. Forgetting `Free` leaks activation contexts and COM references; such leaks are silent, process-scoped, and usually only detected under load.

## Decision

`ComObjectHandle<T>` now implements `IDisposable`:

- `Dispose()` delegates to `Executor.Free(this)` for the actual cleanup and calls `GC.SuppressFinalize(this)` afterwards.
- A finalizer is added purely as a diagnostic safety net. It does **not** attempt native cleanup because the finalizer thread is the wrong COM apartment and deactivating activation contexts from a different thread would corrupt the per-thread cookie stack; instead it calls `Debug.Fail(...)` so that leaks are noisy in Debug builds.
- `IsReleased` is promoted from `internal` to `public` so code can safely check state after dispose without relying on `InternalsVisibleTo`.
- `Executor.Free` remains the canonical release routine (Dispose simply calls it) so existing callers keep working unchanged.

The analyzers that forbid finalizers for non-native-handle types (`MA0055`, `IDISP023`, `CA1821`) are suppressed around the finalizer with an inline comment explaining the diagnostic-only intent.

## Consequences

- Positive: callers can write `using var handle = creation.Value!;` and be done.
- Positive: leaked handles now yell in Debug builds rather than silently accumulating.
- Positive: `Executor.Free` remains; no caller has to migrate eagerly.
- Negative: `ComObjectHandle<T>` is now a "real" `IDisposable` and will trip `IDISP0xx` analyzers in consumer code. Tests in this repo had to add narrow pragma suppressions around intentional Dispose exercises; downstream users may see the same if they write similar tests.
- Negative: the finalizer exists even though it does not clean up native state; consumers reading the source should not mistake it for real safety.

## Alternatives Considered

- Making the finalizer perform cleanup: rejected because `Marshal.FinalReleaseComObject` on the wrong apartment and `DeactivateActCtx` on a thread that does not own the cookie stack are both undefined behaviour.
- Removing `Executor.Free`: rejected because it is a silent breaking change for v2.1.0 callers. Keeping both entry points preserves source compatibility.
- Implementing `IAsyncDisposable`: rejected because the cleanup path is purely synchronous and per-thread; async would be misleading.

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComObjectHandle.cs`
- [ADR-0010](0010-com-object-handle-lifetime-api.md)
