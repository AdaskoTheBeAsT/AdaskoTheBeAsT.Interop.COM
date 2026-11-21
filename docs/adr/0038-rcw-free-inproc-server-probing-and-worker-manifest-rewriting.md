# ADR-0038: RCW-Free In-Proc Server Provenance Probing and Validated Worker SxS Manifest Rewriting

- Status: Accepted
- Date: 2026-11-17
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0
- Extends [ADR-0028](0028-manifest-normalization-proxy-stub-handling-and-hkcu-leases.md), [ADR-0029](0029-vtable-module-isolation-diagnostics.md)

## Context

Two operational friction points remained in isolated per-worker COM staging and verification pipelines (`worker_00\Server_00.dll`, `worker_01\Server_01.dll`):

1. **Managed RCW allocation during pre-flight `CLSID` provenance checks:** `ComModuleDiagnostics.VerifyInstanceModule(object comObject, ...)` ([ADR-0029](0029-vtable-module-isolation-diagnostics.md)) requires an already-materialized managed RCW. When a worker host wants a quick pre-flight check right after `Executor.Activate` to confirm that Win32 `CoCreateInstance(clsid, ..., CLSCTX_INPROC_SERVER, ...)` resolves to the isolated worker DLL rather than a global `HKLM` registration, calling `Activator.CreateInstance(Type.GetTypeFromCLSID(clsid))` allocates a managed RCW in the CLR wrapper table. Even if `FinalReleaseComObject` is omitted or missed on an exception path, uncollected RCWs keep the native DLL's `DllCanUnloadNow()` returning `S_FALSE` until a blocking `GC.Collect()` + `GC.WaitForPendingFinalizers()` cycle runs, defeating deterministic `UnloadModulesOnDeactivate = true` worker recycling.
2. **Ad-hoc XML text replacement when staging renamed worker COM binaries:** Worker pools frequently rename staged server DLLs per worker slot (`Engine.dll` -> `Engine_w00.dll`) so Windows debugger module lists, ETW traces, and SxS assembly identities stay distinct, prune unused companion `<file>` tags not copied into the worker folder, and strip vendor `<description>` strings or XML comments containing licensee PII. Performing those edits with regexes or unvalidated DOM scripts risks path traversal, dangling `<comInterfaceExternalProxyStub tlbid="...">` references, empty manifests lacking `<comClass>`/`<typelib>` entries, BOM mismatches, or accidental in-place corruption of the template manifest.

## Decision

1. **Add RCW-free `CLSID` probing to `ComModuleDiagnostics`:**
   - Expose `GetRegisteredInprocServerPath(Guid clsid, Guid? interfaceId = null, int methodSlotIndex = 0)`, `InspectRegisteredInprocServer(...)`, and `VerifyRegisteredInprocServer(...)`.
   - Invoke raw `ole32/combase!CoCreateInstance(ref clsid, IntPtr.Zero, CLSCTX_INPROC_SERVER, ref iidIUnknown, out IntPtr rawUnknown)` directly, optionally `Marshal.QueryInterface` for a direct native `interfaceId` (`Marshal.Release` in `finally`), resolve the vtable function pointer's `HMODULE` and file path via `GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT)`, and release `rawUnknown` via `Marshal.Release` in `finally`—creating **zero** managed RCWs so `Deactivate` + `FreeLibrary` succeeds immediately with zero `GC.Collect()` calls.
2. **Introduce `ComManifestRewriter` backed by `System.Xml.Linq` and `ComManifestParser` invariants:**
   - Parse manifests via `XmlReader.Create(..., DtdProcessing.Prohibit, XmlResolver = null)` (`FromFile` / `FromXml`) using in-box `System.Xml.Linq` (`XDocument`), keeping zero external NuGet dependencies across all five TFMs.
   - Support fluent, fail-closed transformations: `WithAssemblyIdentityName` (updating root `<assemblyIdentity name="...">`), `RenameFile` / `RenameFiles` (validating safe bare filenames, matching case-insensitively with idempotent acceptance of `replacementFileName`, and throwing `InvalidOperationException` when neither exists), `RemoveFile` (optionally pruning orphan `<comInterfaceExternalProxyStub>` elements whose `tlbid` was owned exclusively by the removed `<file>`), and `StripDescriptions()` (removing all `<description>` elements and `XComment` nodes).
   - Validate output via `ComManifestParser.TryValidateXmlContent` (including rejecting manifests left with zero `<comClass>` and zero `<typelib>` declarations with `ERROR_SXS_CANT_GEN_ACTCTX`) before `ToUtf8Bytes()` or `SaveIfChanged(destinationManifestPath)` emits UTF-8 without BOM, refuse overwriting `sourceManifestPath` in-place, and write atomically via a same-directory temp file only when existing destination bytes differ.

## Consequences

- Positive: Isolated worker hosts can verify `CoCreateInstance` `CLSID` provenance before work begins and still unload worker DLLs immediately on `Deactivate` without forcing garbage collection.
- Positive: Staged worker manifests are rewritten deterministically, stripped of `<description>` / comment PII, validated against the exact same rules `Executor.Activate` enforces, and skipped on disk (`SaveIfChanged == false`) when warm worker directories already match byte-for-byte.
- Negative: `net472`, `net48`, and `net481` reference the in-box framework assembly `System.Xml.Linq` (`<Reference Include="System.Xml.Linq" />`), adding zero NuGet package dependencies.

## Alternatives Considered

- Reading `HKCR\CLSID\{...}\InprocServer32` from the registry in `GetRegisteredInprocServerPath`: rejected because registry inspection does not observe active WinSxS activation contexts (`ActivateActCtx`) or in-process `CoRegisterClassObject` registrations, whereas raw `CoCreateInstance` + vtable address resolution tests the exact dispatch path COM uses at runtime.
- Regex/string substitution for worker manifest rewriting: rejected because XML comments, CDATA, namespace prefixes, orphan `tlbid` proxy/stub references, and structural invariants cannot be enforced reliably without an XML parser and `ComManifestParser` validation.

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComModuleDiagnostics.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComManifestRewriter.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComManifestParser.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/StaPreWarmingLocaleQuirksDiagnosticsAndManifestRewriterTest.cs`
