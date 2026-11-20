# ADR-0035: Bidirectional `ProgID` / `VersionIndependentProgID` Registry Leases and `HKCR`-Presence Short-Circuiting

- Status: Accepted
- Date: 2026-11-17
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0
- Extends [ADR-0028](0028-manifest-normalization-proxy-stub-handling-and-hkcu-leases.md)

## Context

Two practical gaps surfaced when running legacy VB6, Office/VBA, and ATL COM components with `ComActivationOptions.RegisterCurrentUserProgIdsOnFallback = true` (and `RegisterCurrentUserTypeLibsOnFallback = true`):

1. **Unidirectional ProgID registration broke `ProgIDFromCLSID` and version-independent ProgIDs:** `CurrentUserRegistryLeaseManager` previously wrote only the forward `HKCU\Software\Classes\<ProgId>\CLSID` key for a single `progid`. Components that call `ole32!ProgIDFromCLSID` (requiring `HKCU\Software\Classes\CLSID\{CLSID}\ProgID` and `VersionIndependentProgID`) or `CLSIDFromProgID` on a `<comClass versionIndependentProgId="...">` / secondary child `<progid>` failed with `REGDB_E_CLASSNOTREG` (`0x80040154`).
2. **Redundant `HKCU` registry churn on workstations with existing `HKLM` registrations:** When a developer workstation or Windows Server uses `ForceInProcessFallback = true` (for per-worker DLL isolation) and the host OS already has matching `ProgID <-> CLSID` or `TypeLib` entries in `HKLM\Software\Classes` (merged into `HKCR`), writing and deleting identical keys under `HKCU\Software\Classes` on every worker scope adds avoidable registry lock contention and process-crash cleanup exposure.

## Decision

1. **Extend `ComManifestParser` and `CurrentUserRegistryLeaseManager` for full bidirectional `ProgID` + `VersionIndependentProgID` mappings:**
   - Parse `progid` and `versionIndependentProgId` attributes plus all `<progid>` child elements on `<comClass>`, resolving deterministic `PrimaryProgId` and `VersionIndependentProgId` metadata while rejecting conflicting ProgID-to-CLSID declarations across the manifest graph.
   - When `RegisterCurrentUserProgIdsOnFallback` is `true`, acquire reference-counted leases for every forward key (`HKCU\Software\Classes\<ProgId>\CLSID` -> `{CLSID}`) and reverse key (`HKCU\Software\Classes\CLSID\{CLSID}\ProgID` -> `PrimaryProgId`, plus `VersionIndependentProgID` when declared), preserving pre-existing user values and rolling back only library-created keys when the last lease drops to `0`.
2. **Add `ComActivationOptions.SkipCurrentUserRegistrationWhenPresentInClassesRoot` (default `true`):**
   - Before writing any individual `ProgID`, reverse `CLSID`, or `TypeLib` entry to `HKCU\Software\Classes`, probe `Registry.ClassesRoot` read-only. If `HKCR` already maps that exact subkey/value (case-insensitive GUID / ProgID / TypeLib path comparison), return a reference-counted no-op lease for that entry with zero `HKCU` writes.
   - If a concurrent scope acquires the same lease key with `SkipCurrentUserRegistrationWhenPresentInClassesRoot = false`, atomically promote the lease to a materialized `HKCU` write until all overlapping leases release.

## Consequences

- Positive: Both `CLSIDFromProgID` (versioned and version-independent) and `ProgIDFromCLSID` succeed in fallback scopes without `regsvr32` or administrator privileges, and `HKCU` remains completely clean after the last scope deactivates.
- Positive: Desktop/Server worker hosts with existing `HKLM` metadata incur zero `HKCU` registry writes by default while Nano Server containers (where `HKCR` lacks the entries) automatically populate `HKCU\Software\Classes`.
- Negative: `HKCU\Software\Classes` remains per-user hive state shared across threads and processes of the same user account; registry leases remain strictly opt-in (`RegisterCurrentUserProgIdsOnFallback = false` by default).

## Alternatives Considered

- Skipping all `HKCU` writes whenever *any* key for the `CLSID` exists in `HKCR`: rejected in favor of per-key value equivalence checks so stale or mismatched `HKLM` TypeLib paths or re-pointed `CLSID` mappings do not falsely short-circuit a needed `HKCU` override.
- Enabling `RegisterCurrentUserProgIdsOnFallback` by default: rejected so default activation stays 100% registry-write-free.

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComActivationOptions.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComManifestParser.cs`
- `src/AdaskoTheBeAsT.Interop.COM/CurrentUserRegistryLeaseManager.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/ComErrorDiagnosticsAndBidirectionalRegistryTest.cs`
