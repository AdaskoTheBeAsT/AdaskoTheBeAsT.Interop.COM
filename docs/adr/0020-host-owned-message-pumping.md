# ADR-0020: Support Host-Owned Message Pumping and Explicit RCW Ownership

- Status: Accepted
- Date: 2026-09-09
- Release: Unreleased
- Extends ADR-0019

## Context

An unconditional Win32 pump removes custom thread messages without delivering them to a host handler.
It also bypasses UI-framework filtering, accelerators, and dialog preprocessing. The older P/Invoke
branch selected ANSI entry points, unlike the Unicode calls used on modern .NET.

The lifetime API uses `Marshal.FinalReleaseComObject`. Managed aliases are not independent owners:
releasing one shared RCW invalidates every alias.

## Decision

- Keep existing signatures and default pumping behavior for source and binary compatibility.
- Add an immutable `ComExecutor(bool pumpPendingMessages)` policy and corresponding static
  `Execute`/`Create` overloads. Leave `IComExecutor` unchanged so existing implementations still compile.
- Capture the policy in each handle. Both `Free` and `Dispose` honor the creation policy, regardless of
  which executor releases the handle. Do not introduce mutable global or thread-local pump settings.
- Use `PeekMessageW` and `DispatchMessageW` on every target.
- Require factories to transfer exclusive ownership of their RCW. Document alias invalidation,
  separately owned child RCWs, and the prohibition on borrowed/shared wrappers. Do not try to infer
  ownership from RCW reference counts.

## Consequences

Hosts can keep control of their message queues, but must provide any pumping their component needs.
Disabling the explicit pump does not prevent COM's own synchronous-call pumping. Native callback and
window-procedure reentrancy remain possible.

Native integration tests must observe actual object destruction, callbacks and release-thread identity,
and run with DLLs matching the test process architecture.

The destruction-callback tests run with explicit pumping enabled and disabled. They verify the active
context and owner thread during the callback, reject reentrant release, and confirm that disabling
the pump leaves queued window messages for the host.
