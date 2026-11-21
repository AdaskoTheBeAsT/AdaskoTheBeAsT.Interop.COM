# ADR-0032: Explicit WinSxS Bypass (`ForceInProcessFallback`) and Apartment-Scoped Class-Factory Reference Counting

- Status: Accepted
- Date: 2026-11-17
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0
- Extends [ADR-0023](0023-nanoserver-inproc-class-factory-fallback.md), [ADR-0027](0027-reference-counted-module-leases-and-apartment-aware-fallback.md)

## Context

Isolated parallel worker hosts frequently stage distinct on-disk copies of a stateful native COM server (`worker_01\Engine.dll`, `worker_02\Engine.dll`) inside one Windows process so concurrent STA worker threads do not share mutable `.data`/`.bss` globals.

On full Desktop Windows and Windows Server Core, `CreateActCtxW` succeeds and hands activation to the OS Side-by-Side (`sxs.dll`) cache. When staged worker manifests carry identical `<assemblyIdentity>` tuples (or bind to assemblies already resolved in the process SxS cache), WinSxS can coalesce resolution across worker directories so `CoCreateInstance` silently binds both workers to the first loaded `HMODULE`. Meanwhile, the library's existing in-process class-factory registration path (`LoadLibraryExW` + `DllGetClassObject` + `combase!CoRegisterClassObject(..., CLSCTX_INPROC_SERVER, REGCLS_MULTIPLEUSE)`) resolves per-apartment on STA threads (`OXID`-scoped) and bypasses WinSxS identity caching completely, yet was previously reachable only when `CreateActCtxW` failed on headless Windows Nano Server.

In addition, when nested STA scopes or concurrent MTA scopes activate the same `(CLSID, DllPath)` under the in-process fallback, calling `CoRevokeClassObject` unconditionally when the inner or earlier scope deactivates tears down the outer/peer scope's class registration prematurely, while allowing two concurrent scopes in the same apartment to register *different* `DllPath` binaries for the same `CLSID` would cause silent cross-worker hijacking.

## Decision

1. **Add opt-in `ComActivationOptions.ForceInProcessFallback` (default `false`):**
   - When `ForceInProcessFallback` is `true`, `ActivationContextApi` bypasses `CreateActCtxW` even on Desktop Windows and Windows Server Core, routing activation directly through manifest-driven `ComManifestParser`, `ModuleLease` loading, `LoadTypeLibEx`, optional `HKCU` registry leases, and `CoRegisterClassObject`.
   - If `ForceInProcessFallback` is `true` while `NanoServerOptions.EnableSxsFallback` has been explicitly set to `false`, activation fails closed in `Result.Exception` with `InvalidOperationException` rather than silently calling `CreateActCtxW`.
2. **Reference-count fallback `CoRegisterClassObject` leases per `(ApartmentRegistrationKey, Clsid)`:**
   - Partition fallback class-object registrations by COM apartment: per native Win32 thread ID (`GetCurrentThreadId`) for STA threads, and a shared process-wide MTA key for MTA/Implicit-MTA threads.
   - Nested STA scopes or concurrent MTA scopes activating the *same* normalized `DllPath` for an already-leased `(Apartment, Clsid)` increment `SharedClassRegistrationLease.LeaseCount` and reuse the active registration token; `CoRevokeClassObject` and `Marshal.Release(classFactory)` execute only when `LeaseCount` reaches `0`.
   - Attempting to activate a *different* `DllPath` for an already-leased `(Apartment, Clsid)` fails closed with `InvalidOperationException` without clobbering the active registration.

## Consequences

- Positive: Parallel STA worker threads on Desktop Windows and Server Core can simultaneously isolate identical `CLSID`s into distinct per-worker `HMODULE` copies without mutating `<assemblyIdentity>` or fighting OS SxS caching.
- Positive: Nested STA scopes and concurrent MTA fallback scopes safely share and release `CoRegisterClassObject` registrations deterministically, while conflicting `DllPath` collisions inside one apartment fail fast.
- Negative: Within the process-wide MTA, `CoRegisterClassObject` remains apartment-wide; isolating identical `CLSID`s across concurrent workers inside one process still requires dedicated STA worker threads (or direct `ComClassFactory` instantiation).

## Alternatives Considered

- Relying exclusively on rewriting `<assemblyIdentity>` for Desktop SxS isolation: rejected as the sole mechanism because vendor manifests with nested `<dependentAssembly>` graphs or custom internal `CoCreateInstance` calls often still benefit from apartment-scoped `CoRegisterClassObject` routing identical to Nano Server containers.
- Allowing a newer scope in the same apartment to overwrite an active `CLSID`'s `DllPath`: rejected because concurrent MTA callers or outer STA scopes would silently instantiate coclasses from the wrong worker binary.

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComActivationOptions.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ActivationContextApi.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/ForceInProcessFallbackAndClassFactoryTest.cs`
