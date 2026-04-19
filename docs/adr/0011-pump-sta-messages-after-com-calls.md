# ADR-0011: Pump Pending STA Windows Messages Around COM Calls

- Status: Accepted
- Date: 2026-04-07
- Deciders: @AdaskoTheBeAsT
- Release: v2.1.0
- Related commit(s): `5fa56fa`

## Context

COM objects declared with the `Apartment` threading model run on an STA thread and talk to proxies/stubs that deliver cross-apartment calls through the Windows message queue. If the host thread never pumps messages, those callbacks queue up indefinitely and can cause deadlocks, missed events, and stalled finalization. Developers who are used to writing console or worker code frequently forget that creating an STA COM object implicitly obliges them to pump.

## Decision

The library exposes an internal helper `NativeMethods.PumpPendingMessages` that drains the current thread's queue with `PeekMessage(PM_REMOVE)` + `TranslateMessage` + `DispatchMessage`. Every public path that runs user code inside an activation context calls this helper once the user delegate (or factory) returns:

- After the delegate in `Execute` (both overloads)
- After the factory in `Create` (both overloads)
- After releasing the COM object in `Free`

The pump is best-effort: it drains whatever is currently queued without blocking, so it never stalls the caller.

## Consequences

- Positive: callers benefit from correct STA semantics without needing to know about `PeekMessage`/`DispatchMessage`.
- Positive: the cost is bounded (one drain per API call) and negligible when the queue is empty.
- Negative: if user code spawns timers or asynchronous callbacks that *require* sustained pumping, one drain is not enough; such scenarios still need a long-running pump loop (for example, via `AdaskoTheBeAsT.Interop.Threading`).
- Negative: the library silently assumes the calling thread is STA. Callers on an MTA thread will not gain anything from the pump, but also will not fail.

## Alternatives Considered

- Not pumping at all: rejected because it is a common source of COM deadlocks.
- Running a dedicated message loop thread: rejected as out of scope; this belongs to the companion `AdaskoTheBeAsT.Interop.Threading` package and would force callers to pay for a thread they might not want.

## References

- `src/AdaskoTheBeAsT.Interop.COM/NativeMethods.cs` (`PumpPendingMessages`)
- `src/AdaskoTheBeAsT.Interop.COM/Executor.cs`
