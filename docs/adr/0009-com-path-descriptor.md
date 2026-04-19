# ADR-0009: Introduce `ComPathDescriptor` for Multi-Context Activation

- Status: Accepted
- Date: 2025-11-22
- Deciders: @AdaskoTheBeAsT
- Release: v2.0.0
- Related commit(s): `ab14934`

## Context

Real applications frequently need more than one COM component to be simultaneously active in the same activation-context stack (for example, a main component and one or more helper DLLs, each with its own manifest). The v1.x API accepted only a single `(comAssemblyPath, manifestPath)` pair, so supporting two components required two nested `Execute` calls and two nested try/finally blocks — awkward and error-prone because the correct LIFO deactivation order has to be preserved manually.

## Decision

A small, immutable value type `ComPathDescriptor` is introduced that pairs a COM assembly path with a manifest path and validates both in its constructor. A second overload `Executor.Execute(ICollection<ComPathDescriptor>, Action)` accepts an ordered collection of descriptors, activates them in order, runs the callback, and deactivates them in reverse order. The same pattern is repeated later for `Create<T>` (see ADR-0010).

## Consequences

- Positive: multi-component activation is expressed as a single call with a natural input type.
- Positive: activation/deactivation ordering is centralized and verified once.
- Negative: a new public type enters the API surface and must be preserved from now on.
- Negative: misuse is still possible — callers can pass duplicate descriptors, or pass descriptors that reference mutually incompatible manifests; the library cannot detect such conflicts.

## Alternatives Considered

- Tuple of `(string, string)`: rejected because it provides no validation, no XML docs, and no discoverability.
- Overloading with `params` arrays of strings: rejected because it encourages positional errors (swapping DLL/manifest).

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComPathDescriptor.cs`
- `src/AdaskoTheBeAsT.Interop.COM/Executor.cs` (`Execute`/`Create` collection overloads)
