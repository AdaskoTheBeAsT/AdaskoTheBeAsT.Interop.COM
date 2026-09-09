# ADR-0019: Enforce Activation-Context Lifetime and Bound Message Pumping

- Status: Accepted
- Date: 2026-09-08
- Release: Unreleased
- Supersedes lifetime and pumping details in ADR-0010, ADR-0011, and ADR-0012

## Context

Context creation could leak earlier handles when a later manifest failed. Setup failures escaped the documented result boundary. Handle release could reach `DeactivateActCtx` with an invalid thread or stack order, which Windows reports through native exceptions. The message pump could consume `WM_QUIT` or keep draining indefinitely while messages arrived.

## Decision

- Validate arguments before allocating native resources. Keep context creation and activation inside the result boundary and retain each acquired handle and cookie for cleanup, including partial setup.
- Use the directory containing the supplied DLL as the private assembly probing base.
- Track activated cookies per thread across both `Execute` and `Create`. Store managed and native creator-thread identities on each handle. Before releasing an object, verify thread ownership, reverse activation order, native current context, and absence of a reentrant release.
- Check native deactivation results and remove cookies only after successful deactivation. Report cleanup failures through `Result.Exception`, aggregating them with an earlier operation failure when necessary.
- Reject invalid handle release without releasing its COM object. Retain unfinished cleanup state for retry. Mark a handle released and suppress finalization only after its cleanup completes. `Dispose` continues reporting failure through diagnostics rather than throwing.
- Each message pump processes at most 256 messages. On `WM_QUIT`, repost the same exit code and stop. This bounds the number of dispatches, not their execution time; a window procedure can still block or reenter the application.
- Keep execution on the caller's thread. COM initialization and apartment setup remain caller responsibilities, and components needing a sustained message loop require a host-provided loop or scheduler.

## Consequences

- Nested handles must be released in reverse creation order on the creating thread. A rejected `Free` can be retried after nested work completes.
- Handles created inside a callback or factory must be disposed before that delegate returns. Callers using raw activation-context APIs must restore the native stack before returning control to the library.
- Diagnostic finalization remains diagnostic only. It cannot repair an abandoned thread-affine context stack.
- Message-pump budgeting may leave messages queued for the host loop or a later call. The library does not promise progress for asynchronous COM operations through one pump alone.
- Public signatures remain unchanged. Missing or malformed manifests now return failed results; null actions throw argument exceptions instead of silently succeeding.

## Validation

Regression tests cover setup failures, private assembly probing, null delegates and collection elements, wrong-thread and out-of-order release, nested execution, owner-thread retries, preservation of caller apartment state, `WM_QUIT`, and pump budgeting.

An internal, operation-scoped native API boundary enables deterministic failures without global hooks.
Tests delegate successful calls to Windows and inject partial creation, activation, and deactivation
failures. They verify retained cleanup state, release retries, reference accounting, and aggregation
of callback and cleanup exceptions. Each test uses an isolated STA and restores native state on exit.
