# ADR-0026: Introduce Persistent Multi-Coclass Activation Scopes via `IComActivationScope` and Derived `IComActivationExecutor`

- Status: Accepted
- Date: 2026-11-15
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0
- Extends [ADR-0010](0010-com-object-handle-lifetime-api.md), [ADR-0012](0012-com-object-handle-idisposable.md), [ADR-0013](0013-icomexecutor-abstraction.md), [ADR-0018](0018-eventsource-for-leaked-handles.md), and [ADR-0019](0019-safe-activation-context-cleanup.md)

## Context

`Executor.Execute` manages activation contexts only for the duration of a single synchronous callback, while `Executor.Create<T>` / `ComObjectHandle<T>` binds an activation stack to a single root RCW:

1. **Multi-coclass sessions:** Real-world COM calculation engines frequently instantiate several coclasses across one or more manifests during a single session (`Engine`, `Parameters`, `Formatter`, `Recordset`), invoke them in arbitrary order, and tear down all owned RCWs together at session end. Creating separate `ComObjectHandle<T>` instances per coclass duplicates activation-context frames and enforces rigid per-object LIFO release order even when all objects belong to the same logical session.
2. **`IComExecutor` interface compatibility across legacy TFMs:** Adding new abstract methods (`Activate`, `Deactivate`, or `ComActivationOptions` overloads) directly to `IComExecutor` would break existing consumer decorators, adapters, and Moq/hand-written test doubles — and C# default interface implementations (`DIM`) are not supported on `.NET Framework 4.7.2`–`4.8.1`.

## Decision

1. **Keep `IComExecutor` unchanged** with its five existing members.
2. **Introduce `IComActivationExecutor : IComExecutor` and `ComExecutorExtensions`:**
   - Implement `IComActivationExecutor` on `ComExecutor`, exposing `Execute(..., ComActivationOptions)`, `Create(..., ComActivationOptions)`, `Activate(...)`, and `Deactivate(IComActivationScope)`.
   - Provide matching extension methods on `IComExecutor` in `ComExecutorExtensions` that dispatch to `IComActivationExecutor` when implemented and otherwise delegate cleanly without breaking existing `IComExecutor` implementations, plus matching static `Executor.Activate` / `Executor.Deactivate` / `ComActivationOptions` overloads.
3. **Introduce `IComActivationScope` (`IDisposable`), `ComActivationOptions`, and `ComCreateResult<T>`:**
   - An active scope holds the pushed Win32 SxS cookies/handles (or Nano Server fallback registrations), acquired module and registry leases, thread identity, loaded `HMODULE` handles (`LoadedModuleHandles`), fallback indicator (`IsUsingNanoServerFallback`), bounded `PumpPendingMessages(int maxMessages = 256)`, and an ordered list of tracked RCWs registered via `TrackObject<T>(T comObject)`.
   - `Deactivate` validates owner-thread and LIFO activation-stack invariants **before** mutating state (leaving the scope intact for a valid retry on the owner thread if rejected), then releases tracked RCWs in reverse registration order via `Marshal.FinalReleaseComObject`, tears down contexts/registrations, balances any library-acquired COM apartment frame, and releases module/registry leases.
   - Add `ComInteropEventSource` diagnostic events `ScopeLeaked` (ID 3) and `ScopeReleaseFailed` (ID 4) mirroring `HandleLeaked` (ID 1) and `HandleReleaseFailed` (ID 2).

## Consequences

- Positive: Callers can instantiate and use multiple coclasses in any order within one persistent `IComActivationScope` while keeping deterministic owner-thread RCW cleanup and LIFO stack validation.
- Positive: Existing `IComExecutor` implementations, decorators, and mocks compile without modification across `.NET Framework 4.7.2`–`4.8.1` and `.NET 10`–`.NET 11`.
- Negative: Custom `IComExecutor` decorators that also want to intercept new `Activate` / `ComActivationOptions` calls should implement `IComActivationExecutor`.

## Alternatives Considered

- Adding new methods directly to `IComExecutor`: rejected because `.NET Framework` runtimes do not support default interface methods and existing third-party decorators/mocks would break at compile time.
- Requiring callers to nest multiple `ComObjectHandle<T>` instances for multi-coclass sessions: rejected because coclasses within one session often cross-reference one another and are not strictly hierarchical in creation/teardown order.

## References

- `src/AdaskoTheBeAsT.Interop.COM/IComExecutor.cs`
- `src/AdaskoTheBeAsT.Interop.COM/IComActivationExecutor.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComExecutorExtensions.cs`
- `src/AdaskoTheBeAsT.Interop.COM/IComActivationScope.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComActivationScope.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComActivationOptions.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComInteropEventSource.cs`
