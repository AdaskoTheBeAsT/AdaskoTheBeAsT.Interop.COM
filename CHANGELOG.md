# Changelog

Notable changes to `AdaskoTheBeAsT.Interop.COM`.

Dates shown for historical releases refer to repository tags, not NuGet publication dates.

See the [migration guide](MIGRATION.md) for required upgrade steps and the [README](README.md) for current usage.

## [5.1.0]

Adds thread-affine COM instance leases and opt-in x86 Nano Server legacy data access setup.

## [5.0.0]

### Breaking changes

- Upgrade the SDK in `global.json` to `.NET 11.0.100` (`LangVersion` `14.0`), add `net11.0`, and remove `net8.0` and `net9.0` package assets. Supported targets are now `net11.0`, `net10.0`, `net481`, `net48`, and `net472`.
- Validate and normalize manifests prior to activation (`ComManifestParser`): malformed GUIDs, duplicate/conflicting CLSID/ProgID/TypeLib declarations, and unsafe `<file>` paths (path traversal, UNC, Win32 device prefixes, root-relative, or drive-relative paths) now fail deterministically through `Result.Exception` before loading native modules or touching registry keys.
- Honor manifest `threadingModel` (`Apartment` / `Single` -> `STA`, `Free` / `Both` / `Neutral` -> `MTA`) or explicit `ComActivationOptions.ApartmentState` when initializing uninitialized threads during Nano Server fallback instead of unconditionally forcing `COINIT_MULTITHREADED`. Existing threads with a conflicting apartment propagate `RPC_E_CHANGED_MODE` (`0x80010106`) rather than silently mutating apartment state.

### Added

- Persistent multi-coclass activation scopes via `IComActivationScope`, `ComActivationOptions`, `ComCreateResult<T>`, `IComActivationExecutor`, `ComExecutorExtensions`, and `Executor.Activate` / `Executor.Deactivate`, while keeping `IComExecutor` binary- and source-compatible for existing implementations and mocks.
- Opt-in reference-counted native module leases (`ComActivationOptions.PreloadModules`, `ComActivationOptions.UnloadModulesOnDeactivate`), `LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS` (`0x00001100`) sibling DLL resolution, `DllCanUnloadNow() == S_OK` + `CoFreeUnusedLibrariesEx(0, 0)` unload gating after tracked RCW release and factory revocation, and `Executor.TryUnloadUnusedModules()` for sweeping deferred modules once leaked native references drop to zero.
- Manifest `progid`, `versionIndependentProgId`, and child `<progid>`, `<typelib>`, `<comInterfaceProxyStub>`, and `<comInterfaceExternalProxyStub>` support on in-process fallback, including `oleaut32!LoadTypeLibEx`, declared proxy/stub `CLASS_E_CLASSNOTAVAILABLE` tolerance (`IgnoreUnavailableProxyStubClasses`) without masking unavailable ordinary coclasses, bidirectional `HKCU\Software\Classes\<ProgId>\CLSID` and `HKCU\Software\Classes\CLSID\{CLSID}\ProgID` + `VersionIndependentProgID` reference-counted leases (`RegisterCurrentUserProgIdsOnFallback`, `RegisterCurrentUserTypeLibsOnFallback`), and read-only `HKCR` presence short-circuiting (`SkipCurrentUserRegistrationWhenPresentInClassesRoot = true` by default) to avoid redundant `HKCU` writes when matching machine registrations already exist. Manifest `<description>` licensee text is never retained or emitted to diagnostics.
- Explicit WinSxS bypass (`ComActivationOptions.ForceInProcessFallback`) for Desktop Windows and Windows Server Core isolated STA worker pools, plus apartment-scoped (`(Apartment, Clsid)`) `CoRegisterClassObject` reference counting across nested STA scopes and concurrent MTA scopes with fail-closed conflicting `DllPath` detection.
- Direct zero-manifest / zero-registry `HMODULE` instantiation via `ComClassFactory` (`CreateInstance`, `CreateInstance<TInterface>`, `CreateInstanceFromDll`, `CreateInstanceFromDll<TInterface>`) and scope-tracked `IComActivationScope.CreateInstance` / `ComActivationScopeExtensions.CreateInstance` helpers.
- `ComModuleDiagnostics` (`GetInstanceModuleHandle`, `InspectInstanceModule`, `VerifyInstanceModule`, RCW-free `GetRegisteredInprocServerPath`, `InspectRegisteredInprocServer`, `VerifyRegisteredInprocServer`, `ComModuleInspectionResult`, `ComModuleVerificationStatus`) to inspect vtable slot ownership by `HMODULE` and path on live instances or raw `CoCreateInstance` probes without allocating managed RCWs, distinguishing concrete worker-DLL mismatches from `Inconclusive` shared runtime/proxy modules (`combase.dll`, `ole32.dll`, `oleaut32.dll`, `rpcrt4.dll`, `msvbvm60.dll`, CLR).
- `ComErrorDiagnostics` and `ComErrorInfo` (`ConsumeCurrentErrorInfo`, `PeekCurrentErrorInfo`, `CreateExceptionForHr`, `TryGetVisualBasicErrorNumber`, `TryGetInterfaceErrorOffset`) for destructive and non-destructive thread `IErrorInfo` inspection, enriched `COMException` creation (`Data["ComErrorInfo"]`), and VB6 `FACILITY_CONTROL` (`0x800Axxxx`) / `FACILITY_ITF` (`0x80040200+`) `HRESULT` unpacking.
- `ComAutomationLocaleScope` (`Enter`, `EnterInvariant`, `UsEnglishLocaleId = 0x0409`, `UkEnglishLocaleId = 0x0809`, `InvariantLocaleId = 0x007F`) for thread-affine synchronization of Win32 kernel `GetThreadLocale`/`SetThreadLocale` and `CultureInfo.CurrentCulture` across OLE Automation calls.
- Automatic CoreCLR native STA `OleMainThreadWndClass` pre-warming via `NanoServerOptions.EnsureCurrentStaThreadInitialized()` on STA entry points (`Executor.Activate`, `Execute`, `Create`), and opt-in per-module W^X unbalanced uninitialize guards via `ComModuleQuirks.SuppressUninitializeImports` / `RestoreUninitializeImports` (rejecting OS, COM, and CLR system modules).
- `ComManifestRewriter` (`FromFile`, `FromXml`, `WithAssemblyIdentityName`, `RenameFile`, `RenameFiles`, `RemoveFile`, `StripDescriptions`, `ToXmlString`, `ToUtf8Bytes`, `SaveIfChanged`) for validated, idempotent, PII-sanitized worker SxS manifest staging.
- `SafeArrayMarshal` utilities for borrowed vs. owned unmanaged `SAFEARRAY*` pointers (including Borland Delphi arrays omitting `FADF_HAVEVARTYPE`), ownership-aware `BSTR` and `VARIANT` arrays, `0`-based / `1`-based / negative lower-bound normalization (`NormalizeVariantArray<T>`), and VB6 `1`-based array creation (`CreateOneBasedArray<T>`).
- Opt-in W^X Nano Server compatibility providers for older container images: `NanoServerOptions.EnableLegacyOle2Shim` (`ADASKO_NANOSERVER_ENABLE_LEGACY_OLE2_SHIM`) patching `OleRun`, `OleSaveToStream`, `OleLoadFromStream`, `CoInitialize`, `OleInitialize`, and `OleUninitialize` IAT slots, and `NanoServerOptions.EnableDispenserManagerFallback` (`ADASKO_NANOSERVER_ENABLE_DISPENSER_FALLBACK`) registering a static-lifetime pure native x86/x64 `CLSID_DispenserManager` provider (`PAGE_READWRITE` -> `PAGE_EXECUTE_READ`), automatically deferring when OS `nanoserver_sta_shim.dll` advertises capabilities via `NanoServerStaShim_GetCapabilities`.
- Idempotent startup preparation (`NanoServerOptions.PrepareProcess()`) and lock-free `Volatile.Read` fast paths for multi-STA worker hosts.
- `ComInteropEventSource` events `ScopeLeaked` (ID 3) and `ScopeReleaseFailed` (ID 4).
- [ADR-0025](docs/adr/0025-update-target-frameworks-to-net11-net10-and-netframework.md), [ADR-0026](docs/adr/0026-activation-scopes-and-derived-icomactivationexecutor.md), [ADR-0027](docs/adr/0027-reference-counted-module-leases-and-apartment-aware-fallback.md), [ADR-0028](docs/adr/0028-manifest-normalization-proxy-stub-handling-and-hkcu-leases.md), [ADR-0029](docs/adr/0029-vtable-module-isolation-diagnostics.md), [ADR-0030](docs/adr/0030-safearray-marshal-for-vb6-and-delphi-interop.md), [ADR-0031](docs/adr/0031-wx-enforcement-legacy-ole2-dispenser-manager-and-startup-prep.md), [ADR-0032](docs/adr/0032-force-inproc-fallback-and-apartment-class-factory-refcounting.md), [ADR-0033](docs/adr/0033-direct-hmodule-class-factory-instantiation.md), [ADR-0034](docs/adr/0034-thread-ierrorinfo-diagnostics-and-vb6-hresult-unpacking.md), [ADR-0035](docs/adr/0035-bidirectional-progid-leases-and-hkcu-short-circuiting.md), [ADR-0036](docs/adr/0036-coreclr-sta-prewarming-and-per-module-couninitialize-guard.md), [ADR-0037](docs/adr/0037-thread-affine-automation-locale-scope.md), and [ADR-0038](docs/adr/0038-rcw-free-inproc-server-probing-and-worker-manifest-rewriting.md), plus expanded x86/x64 native C++ and managed unit tests covering all new activation, marshaling, diagnostic, manifest-rewriting, and W^X Nano Server subsystems.

### Security

- Eliminate `PAGE_EXECUTE_READWRITE` (`0x40`) from `NanoServerStaShim` IAT patching in favor of temporary `PAGE_READWRITE` (`0x04`) and restoration of the original page protection via `VirtualProtect`, enforcing strict Write-Xor-Execute (W^X) across the entire library.

## [4.2.0]

### Added

- Automatic in-process Single-Threaded Apartment (`STA` / `COINIT_APARTMENTTHREADED`) and user-mode Win32 window & message-queue shim (`NanoServerStaShim` + `NanoServerWindowManager`) on **Windows Nano Server 2025 (`ltsc2025`)** x86 and x64 when `win32u.dll` is absent, enabling `ApartmentState.STA`, `[STAThread]`, and `pumpPendingMessages: true` without `E_NOTIMPL` (`0x80004001`).
- Public `NanoServerOptions` configuration API (`EnableSxsFallback`, `EnableStaShim`, `ForceInProcessStaShim`, `IsInProcessStaShimActive`, `RefreshStaState()`, `ResetToEnvironmentDefaults()`) and environment-variable toggles (`ADASKO_NANOSERVER_DISABLE_SXS_FALLBACK`, `ADASKO_NANOSERVER_DISABLE_STA_SHIM`, `NANOSERVER_DISABLE_STA_SHIM`) for runtime diagnostics and testing.
- Seamless coordination with the OS-level native `nanoserver_sta_shim.dll` in `adaskothebeast/nanoserver` container images (`NanoServerStaShim_Refresh` / `NanoServerStaShim_RestoreOriginal`) and automatic `%SystemRoot%\System32\forwarders` (`SysWOW64\forwarders` on x86) `PATH` initialization via `[ModuleInitializer]`.
- [ADR-0024](docs/adr/0024-nanoserver-sta-shim-and-options.md) and expanded container and host unit test coverage across all STA shim toggle modes.

## [4.1.0]

### Added

- Automatic in-process COM class-factory fallback (`DllGetClassObject` + `combase!CoRegisterClassObject` / `CoRevokeClassObject`) when Windows Side-by-Side (SxS) activation contexts (`CreateActCtxW`) are unavailable on the host OS (`ERROR_INVALID_FUNCTION`, `ERROR_NOT_SUPPORTED`, or `ERROR_CALL_NOT_IMPLEMENTED`), enabling registration-free COM in **Windows Nano Server 2025 (`ltsc2025`)** x86 and x64 containers.
- Recursive manifest parsing for `<file>` and `<dependentAssembly>` elements in fallback mode with cycle protection and process-lifetime native module (`HMODULE`) caching.
- Graceful handling of missing `user32.dll` / `PeekMessageW` entry points during message pumping on headless OneCore / Nano Server images.
- [ADR-0023](docs/adr/0023-nanoserver-inproc-class-factory-fallback.md) and unit test suite (`NanoServerFallbackTest`) covering Nano Server fallback execution, LIFO handle semantics, dependent manifests, and native failure modes.

## [4.0.1]

### Fixed

- Preserve `Platform=AnyCPU` on the test project's `ProjectReference` during `AnyCPU` solution builds so MSBuild evaluates `AdaskoTheBeAsT.Interop.COM` once instead of twice in parallel, preventing `NU5026` `GeneratePackageOnBuild` races during `--no-incremental` rebuilds.
- Populate NuGet `<PackageReleaseNotes>` metadata in `AdaskoTheBeAsT.Interop.COM.csproj`.

### Build, tests, and documentation

- Update the SDK selected by `global.json` to `10.0.402`.
- Update analyzer and test dependencies (`StyleCop.Analyzers 1.3.151-beta`, `SonarAnalyzer.CSharp 10.35.0.4138`, `Meziantou.Analyzer 3.0.316`, `AdaskoTheBeAsT.CodeCracker.CSharp 1.1.0.270`, `Microsoft.CodeAnalysis.NetAnalyzers 10.0.402`, `Microsoft.VisualStudio.Threading.Analyzers 18.12.1`, `Roslynator.* 5.0.1`, `Microsoft.NET.Test.Sdk 18.11.0`, `Microsoft.Testing.Extensions.GitHubActionsReport 2.5.1`, `Moq 4.21.0`, `xunit.v3 4.0.2`, `xunit.analyzers 2.2.0`, `xunit.runner.visualstudio 4.0.1`) and satisfy all newly enabled analyzer diagnostics (`SA1121`, `CC0105`, `CC0001`, `CA1060`, `S2094`).
- Consolidate shared test helpers (`TestPaths`, `CapturingEventListener`) and expand unit tests for `ActCtxWrongSizeException` binary serialization on .NET Framework and `Executor.ValidateReleaseOrder` edge cases.
- Add [ADR-0021](docs/adr/0021-consolidate-supported-target-frameworks.md) and [ADR-0022](docs/adr/0022-explicit-intptr-and-analyzer-hygiene.md).

## [4.0.0]

### Breaking changes

- Remove `netstandard2.0`, `net462`, `net47`, and `net471` assets. The supported targets are now `net8.0`, `net9.0`, `net10.0`, `net472`, `net48`, and `net481`.
- Enforce handle release on the creating managed and native thread, in reverse activation order. Wrong-thread, reentrant, and out-of-order release returns a failed `Result` before releasing the COM object.
- Return operational activation-context setup failures through `Result.Exception`. Callers that previously caught setup exceptions must also check the result. Invalid arguments still throw.
- Bound each automatic message pump to 256 messages instead of draining the queue indefinitely. Applications must not rely on an API call emptying their message queue.

### Added

- `ComExecutor(bool pumpPendingMessages)` for an immutable message-pumping policy.
- Additive `Executor.Execute` and `Executor.Create` overloads with a final `bool pumpPendingMessages` parameter, for both single and multiple manifests.
- Handles retain the policy selected at creation for `Free` and `Dispose`, including release through another executor. Existing signatures still enable pumping by default; `IComExecutor` has no new members.
- Source-built native lifetime tests for actual COM destruction, callbacks, alias invalidation, and reentrant release.
- Regression tests for activation failures, release ordering, thread ownership, Unicode messages, `WM_QUIT`, bounded pumping, and pump policies.

### Fixed

- Release already-created activation contexts when a later manifest fails to load, and clean up after partial activation failures.
- Validate null actions and null entries in descriptor collections before allocating native resources.
- Use the COM DLL's containing directory, rather than the DLL path itself, as the assembly directory for activation-context probing.
- Check native deactivation results and preserve unfinished handle cleanup state for retry. Combine work and cleanup failures in `AggregateException` when both occur.
- Keep leak diagnostics enabled after rejected or incomplete disposal instead of suppressing finalization unconditionally.
- Preserve `WM_QUIT` and its exit code for the host message loop.
- Use Unicode message retrieval and dispatch on .NET Framework as well as modern .NET.
- Select native test payloads by explicit process architecture and isolate managed outputs by architecture to avoid reusing incompatible assemblies.

### Build and documentation

- Update the SDK selected by `global.json` to `10.0.401` and select Microsoft.Testing.Platform for `dotnet test`.
- Update analyzer and test dependencies, including xUnit v3 across the remaining test targets.
- Document caller-owned COM initialization, STA setup, message processing, and exclusive RCW ownership.
- Reorganize the README and add standalone changelog, migration, advanced usage, and native build guides.
- Add [ADR-0019](docs/adr/0019-safe-activation-context-cleanup.md) and [ADR-0020](docs/adr/0020-host-owned-message-pumping.md).

## [3.0.0] - 2026-04-19

### Added

- `IComExecutor` and its `ComExecutor` implementation for dependency injection and testing. Existing static `Executor` entry points remain available.
- `IDisposable` on `ComObjectHandle<T>` and public `IsReleased` state.
- Diagnostic-only handle finalization and the internal `AdaskoTheBeAsT.Interop.COM` event source: `HandleLeaked` (ID 1) and `HandleReleaseFailed` (ID 2).
- Windows platform annotations on `Executor`, `IComExecutor`, `ComExecutor`, and `ComObjectHandle<T>` for modern .NET targets. Consumers may need to address new `CA1416` warnings.
- Dedicated .NET Framework targets: `net462`, `net47`, `net471`, `net472`, `net48`, and `net481`, alongside existing `netstandard2.0`, `net8.0`, `net9.0`, and `net10.0` assets.
- Shared GitHub Actions CI workflow, architecture decision records, and expanded managed tests.

### Changed

- Build the managed library as AnyCPU instead of x86. Native COM components and their host processes must still have matching architectures.

### Fixed

- Compatibility of native declarations and exception code across the target framework matrix.

## [2.1.0] - 2026-04-07

### Added

- `Executor.Create<T>` / `Executor.Free<T>`, `ComObjectHandle<T>`, and `ComObjectCreationResult<T>` for explicit object lifetime management.
- Automatic pending Windows message pumping around COM operations.
- Lifetime and STA scheduling examples in the README.

The tagged package targets `netstandard2.0`, `net8.0`, `net9.0`, and `net10.0`; its managed assembly is x86.

## Earlier history

Older README and ADR entries use version labels that do not line up consistently with repository tags. They are preserved here as development milestones rather than assigned unverified release versions or dates:

- **2025-08-10 to 2025-11-23:** multi-context activation through `ComPathDescriptor`, modern .NET targets, source-generated P/Invoke on modern .NET, and expanded package documentation. The repository's [`v1.0.0` tag](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/tree/v1.0.0) points to the 2025-11-23 commit, not the initial implementation. There is no `v2.0.0` tag.
- **2023-12-09:** initial implementation with static `Executor`, `Result`, a `netstandard2.0` target, and internal Win32 activation-context wrappers.

The [architecture decision index](docs/adr/README.md) provides additional historical context.

[5.0.0]: https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/compare/v4.2.0...v5.0.0
[4.2.0]: https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/compare/v4.1.0...v4.2.0
[4.1.0]: https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/compare/v4.0.1...v4.1.0
[4.0.1]: https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/compare/v4.0.0...v4.0.1
[4.0.0]: https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/compare/v3.0.0...v4.0.0
[3.0.0]: https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/compare/v2.1.0...v3.0.0
[2.1.0]: https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/releases/tag/v2.1.0
