# ADR-0018: Diagnostic `EventSource` for Leaked `ComObjectHandle<T>`

- Status: Accepted
- Date: 2026-04-19
- Deciders: @AdaskoTheBeAsT
- Release: v3.0.0
- Related ADRs: [ADR-0012](0012-com-object-handle-idisposable.md)

## Context

Since v3.0.0 `ComObjectHandle<T>` implements `IDisposable` (ADR-0012) and is given a diagnostic finalizer as a safety net: if the caller forgets to call `Dispose()` / `Executor.Free`, the finalizer fires and raises a `Debug.Fail(...)` asserting the leak. That works well during development but has two problems in production:

- `Debug.Fail` is decorated with `[Conditional("DEBUG")]`. In Release builds the call is compiled away entirely, so the finalizer becomes a silent no-op and real-world leaks are invisible.
- Even in Debug, the signal is confined to `Debug.Listeners` — it cannot be consumed by an out-of-process diagnostic tool like `dotnet-trace`, PerfView, ETW, or an OpenTelemetry exporter.

We already recommend the decorator pattern for `ILogger<T>` and `System.Diagnostics.Metrics.Meter` integration in the README, but leaks are detected by the *finalizer*, which runs outside any user decorator: the handle has by definition escaped the caller. A first-party, library-owned signal is therefore needed.

## Decision

Introduce an internal `EventSource` with the provider name **`AdaskoTheBeAsT.Interop.COM`** (`ComInteropEventSource`). It exposes a single warning-level event:

- **Event ID 1 — `HandleLeaked(string typeName)`** — emitted from the finalizer of `ComObjectHandle<T>` when `IsReleased` is `false`. Payload: the short name of the generic type argument `T`.

The finalizer now does two things when a leak is detected:

1. `ComInteropEventSource.Log.HandleLeaked(typeof(T).Name)` — always runs. Survives in Release builds and is observable externally.
2. `Debug.Fail(...)` — guarded by `Debugger.IsAttached`. It only fires when a developer is actively debugging, so it cannot deadlock an unattended test runner or a production host on a stray `TraceListener` configuration while still giving developers the breakpoint-level signal they expect.

Consumers can subscribe out-of-process:

```bash
dotnet-trace collect --providers AdaskoTheBeAsT.Interop.COM --process-id <pid>
```

…or in-process via `EventListener`:

```csharp
sealed class MyListener : EventListener
{
    protected override void OnEventSourceCreated(EventSource source)
    {
        if (source.Name == "AdaskoTheBeAsT.Interop.COM")
            EnableEvents(source, EventLevel.Warning);
    }
    protected override void OnEventWritten(EventWrittenEventArgs e) { /* ... */ }
}
```

The event source is `internal` because:

- `EventPipe` and ETW subscribe by provider **name**, not by CLR type, so accessibility is irrelevant to external tooling.
- Keeping it internal avoids committing to a public shape for the future; we can add payloads, tags, or more events without an API break.
- The test assembly already has `InternalsVisibleTo`, so unit tests can still exercise the source directly.

The `EventSource.IsEnabled()` guard around `WriteEvent` ensures there is no string-boxing or payload allocation when no listener is attached — the hot (finalizer) path stays allocation-free.

## Consequences

- Positive: production leaks are now observable via `dotnet-trace`, PerfView, OpenTelemetry, Application Insights, and any standard .NET diagnostics pipeline.
- Positive: the finalizer has real Release-build code (the `WriteEvent` call), so the `CA1821` suppression on "empty finalizer" is no longer needed.
- Positive: subscribers are zero-cost when inactive — `IsEnabled()` short-circuits emission.
- Positive: payload carries the generic type name, so a single listener can distinguish leaks of different COM types.
- Negative: the new type `ComInteropEventSource` adds a small amount of surface to the assembly (though not to the public API).
- Negative: consumers who already rely on `Debug.Listeners` will see two signals now (`Debug.Fail` + `EventWrittenEventArgs`) in a Debug build. This is additive, not breaking.
- Negative: `EventSource` has strict method-shape rules (`[Event]` attribute, event ID monotonicity). Anyone adding a new event must follow them or the BCL self-test will fail at runtime.

## Alternatives Considered

- **Keep only `Debug.Fail`** — rejected: invisible in Release, out-of-process unreachable.
- **Use `ILogger<T>` directly from the finalizer** — rejected: the handle has no DI scope reference, logger lifetime is not guaranteed during finalization, and logging from a finalizer is discouraged. `EventSource` is explicitly designed for this scenario.
- **Write to `Trace.TraceWarning`** — rejected: no structured payload, inconsistent default listeners across host types, no ETW / EventPipe integration.
- **Throw in the finalizer** — rejected: throwing from a finalizer terminates the process on modern .NET and destroys the very data needed to diagnose the leak.
- **Attempt native cleanup in the finalizer** — rejected: apartment-affine handles cannot be freed from the finalizer thread without potential deadlock / heap corruption, which is precisely why the finalizer exists only as a diagnostic.

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComInteropEventSource.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComObjectHandle.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/ComInteropEventSourceTest.cs`
- [ADR-0012](0012-com-object-handle-idisposable.md)
