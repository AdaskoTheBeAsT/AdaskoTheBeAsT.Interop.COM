# ADR-0028: Pre-Activation Manifest Normalization, Declared Proxy/Stub Differentiation, and Reference-Counted `HKCU` Registry Leases

- Status: Accepted
- Date: 2026-11-15
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0
- Extends [ADR-0002](0002-registration-free-com-via-activation-contexts.md) and [ADR-0023](0023-nanoserver-inproc-class-factory-fallback.md)

## Context

Real-world VB6, Delphi, and ATL/MIDL side-by-side manifests exposed four limitations in v4.1–v4.2's fallback parser:

1. **Late failure on malformed or conflicting manifests:** Malformed GUIDs, conflicting duplicate `<comClass>` / `<progid>` / `<typelib>` declarations across nested manifests, or unsafe `<file name="...">` paths (relative traversal `..`, UNC shares, Win32 device paths `\\?\`, root-relative `\...`, or drive-relative `C:...`) must be rejected **before** loading any native DLL or touching registry keys.
2. **Licensee metadata leakage:** Commercial COM manifests sometimes embed customer or vendor license strings inside `<comClass description="...">` or `<typelib>` descriptions; retaining or emitting these descriptions into exception messages or logs leaks sensitive licensee strings.
3. **Proxy/stub classes vs. unavailable coclasses:** Manifests generated for MIDL/ATL servers frequently declare `<comInterfaceProxyStub>` / `PSFactoryBuffer` CLSIDs that return `CLASS_E_CLASSNOTAVAILABLE` (`0x80040111`) from `DllGetClassObject` when unused in-apartment. However, blanket-ignoring `CLASS_E_CLASSNOTAVAILABLE` on all `<comClass>` entries would mask genuinely broken or non-creatable coclasses.
4. **ProgID (`CLSIDFromProgID`) and TypeLib (`LoadRegTypeLib`) lookups on Nano Server:** Legacy VB6 (`CreateObject("Lib.Class")`) and Automation/dual-interface servers query `ProgID` and `TypeLib` entries that SxS normally virtualizes. Writing unscoped `HKCU\Software\Classes` keys would clobber pre-existing machine/user registrations or break concurrent overlapping scopes when one scope exits.

## Decision

1. **Normalize and validate manifests upfront in `ComManifestParser`:**
   - Parse `<file>`, `<comClass>`, `progid` attributes and child `<progid>` elements, `<typelib>`, `<comInterfaceProxyStub>`, `<comInterfaceExternalProxyStub>`, and nested `<dependentAssembly>` files using `XmlReader` (`DtdProcessing.Prohibit`, `XmlResolver = null`).
   - Reject malformed `clsid`/`tlbid`/`iid` GUIDs, conflicting duplicate declarations (allowing only idempotent identical re-declarations), and unsafe `<file>` paths before loading any native module or writing any registry entry.
   - Never retain, store, or log `<description>` attributes from manifests.
2. **Distinguish declared proxy/stub classes from ordinary coclasses:**
   - Identify proxy/stub classes exclusively from explicit manifest metadata (`<comInterfaceProxyStub>`, `<comInterfaceExternalProxyStub>`, or a `<comClass>` whose CLSID matches a declared `proxyStubClsid32`).
   - When `ComActivationOptions.IgnoreUnavailableProxyStubClasses = true` (default), tolerate `CLASS_E_CLASSNOTAVAILABLE` (`0x80040111`) **only** for declared proxy/stub classes; an unavailable ordinary coclass remains a hard activation failure.
3. **Keep activation registry-free by default, with opt-in reference-counted `HKCU\Software\Classes` leases (`CurrentUserRegistryLeaseManager`):**
   - Load `<typelib>` entries in-memory via `oleaut32!LoadTypeLibEx(..., REGKIND_NONE)` when `LoadTypeLibsOnFallback = true`.
   - Only write `HKCU\Software\Classes\<ProgId>\CLSID` (`RegisterCurrentUserProgIdsOnFallback = true`) or `HKCU\Software\Classes\TypeLib\{tlbid}\...` (`RegisterCurrentUserTypeLibsOnFallback = true`) when explicitly enabled by the caller.
   - Snapshot pre-existing `HKCU` key/value state before first write, reference-count identical concurrent scope leases, conflict-fail overlapping incompatible targets, and on release roll back or delete keys **only** if their current registry values still match the values written by the lease (preserving external edits and pre-existing registrations).

## Consequences

- Positive: Invalid, traversal, or conflicting manifests fail cleanly with zero native module loads or registry side effects, and licensee `<description>` text never enters memory or logs.
- Positive: MIDL `PSFactoryBuffer` entries no longer abort Nano Server fallback activation, while genuinely missing coclasses still fail deterministically.
- Positive: Legacy `CLSIDFromProgID` and `LoadRegTypeLib` callers work under unprivileged accounts (`ContainerUser`) when opted in, with safe concurrent lease counting and rollback.
- Negative: Opt-in `HKCU\Software\Classes` leases remain per-user/process-visible while active rather than per-thread like kernel WinSxS.

## Alternatives Considered

- Blanket-ignoring `CLASS_E_CLASSNOTAVAILABLE` for every `<comClass>`: rejected because typos or uncreatable coclasses would silently succeed during `Activate` and fail unpredictably later.
- Enabling `HKCU` ProgID/TypeLib writes by default: rejected to preserve strict registry-free semantics unless explicitly requested by the consumer.

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComManifestParser.cs`
- `src/AdaskoTheBeAsT.Interop.COM/CurrentUserRegistryLeaseManager.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ActivationContextApi.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/ComManifestValidationAndFallbackOptionsTest.cs`
