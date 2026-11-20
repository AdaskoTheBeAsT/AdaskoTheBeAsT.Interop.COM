# ADR-0031: Enforce Strict W^X Memory Protection, Opt-In Legacy OLE2 & `CLSID_DispenserManager` Fallbacks, and Startup Preparation on Nano Server

- Status: Accepted
- Date: 2026-11-15
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0
- Extends [ADR-0024](0024-nanoserver-sta-shim-and-options.md)

## Context

Three production security, compatibility, and startup-ordering requirements arose on headless Windows Nano Server 2025 containers:

1. **Strict Write-Xor-Execute (W^X) compliance:** `NanoServerStaShim` previously temporarily marked IAT pages with `0x40` (`PAGE_EXECUTE_READWRITE`) during pointer patching. Modern security baselines and Arbitrarily Executable Memory (`ACG` / Defender Exploit Guard) audits forbid any page from being simultaneously writable and executable, even transiently.
2. **Older Nano Server container compatibility (`OLE2` delay-load stubs and `CLSID_DispenserManager`):**
   - On Nano Server images whose OS-level `nanoserver_sta_shim.dll` precedes capability support (or stock Microsoft Nano Server images), `ole32.dll` delay-load stubs for `OleRun`, `OleSaveToStream`, `OleLoadFromStream`, `CoInitialize`, `OleInitialize`, and `OleUninitialize` fail unless redirected in loaded modules' Import Address Tables, and `CLSID_DispenserManager` (`{ECABB0C0-7F19-11D2-978E-0000F8757E2A}`, normally provided by `comsvcs.dll`) is absent for MDAC/ADO/OLE DB/Jet 4.0 resource pooling.
   - Crucially, background native OLE DB / Jet worker threads detach and invoke `IHolder::Close()` / `IUnknown::Release()` during `DLL_THREAD_DETACH` or process shutdown **after** the .NET CLR has begun shutting down, where reverse-P/Invoke managed delegates would fault. Any in-process `CLSID_DispenserManager` fallback must therefore reside in static-lifetime pure native machine code (`PAGE_READWRITE` -> `PAGE_EXECUTE_READ`) and must yield automatically when `nanoserver_sta_shim.dll` advertises native OS support via a capability query rather than brittle version-string parsing.
3. **Concurrent STA worker startup ordering:** Initializing `%SystemRoot%\System32\forwarders` (`SysWOW64\forwarders` on x86) `PATH` entries, preloading base DLLs, and patching IATs lazily under contended locks when dozens of STA worker threads start simultaneously adds avoidable lock contention on partially initialized STA threads.

## Decision

1. **Eliminate simultaneous write+execute memory protection across the entire library:**
   - Update `NanoServerStaShim` and `NanoServerOle2Shim` IAT slot patching to transition target IAT pages temporarily to `PAGE_READWRITE` (`0x04`) and immediately restore `oldProtect` via `VirtualProtect`.
2. **Select Nano Server OLE2 and `CLSID_DispenserManager` providers via capability export (`NanoServerStaShim_GetCapabilities`):**
   - Query `nanoserver_sta_shim.dll!NanoServerStaShim_GetCapabilities()` bit flags (`0x01` STA/Windowing, `0x02` Legacy OLE2, `0x04` `CLSID_DispenserManager`). When the OS shim advertises a capability (and `ForceInProcessStaShim` is false), defer cleanly to the OS provider without double-patching or overriding registrations.
   - Expose opt-in `NanoServerOptions.EnableLegacyOle2Shim` (`ADASKO_NANOSERVER_ENABLE_LEGACY_OLE2_SHIM`) backed by `NanoServerOle2Shim` (`OleRun` via `QI(IRunnableObject)->Run`, `OleSaveToStream` via `QI(IPersistStream)->GetClassID` + `WriteClassStm` + `Save`, `OleLoadFromStream` via `ReadClassStm` + `CoCreateInstance` + `QI(IPersistStream)->Load`, and `CoInitialize`/`OleInitialize`/`OleUninitialize` forwarding to `combase.dll`).
   - Expose opt-in `NanoServerOptions.EnableDispenserManagerFallback` (`ADASKO_NANOSERVER_ENABLE_DISPENSER_FALLBACK`) backed by `NanoServerDispenserManager`, which allocates separate data (`PAGE_READWRITE`) and code (`PAGE_READWRITE` -> `FlushInstructionCache` -> `PAGE_EXECUTE_READ`) pages for static process lifetime implementing `IClassFactory`, `IDispenserManager` (`RegisterDispenser`, `GetContext`), and pass-through `IHolder` (`AllocResource` -> `CreateResource`, `FreeResource` -> `DestroyResource`, `TrackResource`, `TrackResourceS`, `UntrackResource`, `Close`) in pure x86 (`__stdcall`) and x64 (`fastcall`) native code with zero managed callbacks.
3. **Add idempotent `NanoServerOptions.PrepareProcess()` and lock-free hot-path checks:**
   - Expose `NanoServerOptions.PrepareProcess()` so hosts can perform all forwarder `PATH` initialization, base DLL preloading, `NanoServerMainStaHost` startup, and configured shim installation once on the startup thread before spawning concurrent STA workers, while guarding hot-path activation reads with lock-free `Volatile.Read`.

## Consequences

- Positive: Zero RWX memory allocations or `VirtualProtect` calls exist anywhere in `AdaskoTheBeAsT.Interop.COM`.
- Positive: Legacy OLE2 stream persistence and MDAC/Jet 4.0 `CLSID_DispenserManager` pooling work on both older and newer Nano Server images across x86 and x64 without risking CLR-shutdown reverse-P/Invoke crashes.
- Positive: Concurrent STA worker pools avoid first-call initialization contention by calling `NanoServerOptions.PrepareProcess()` at startup.
- Negative: `NanoServerDispenserManager`'s fallback `IHolder` intentionally performs pass-through resource creation/destruction (`CreateResource` / `DestroyResource`) rather than idle connection pooling when `comsvcs.dll` is absent, trading connection reuse for deterministic leak-free native shutdown safety.

## Alternatives Considered

- Implementing `IDispenserManager` / `IHolder` via `Marshal.GetFunctionPointerForDelegate` reverse-P/Invoke delegates: rejected because native MDAC/Jet background threads invoke `IHolder::Close` and `Release` during process teardown when managed execution can no longer run safely.
- Detecting OS shim support by parsing `nanoserver_sta_shim.dll` file version strings: rejected in favor of `NanoServerStaShim_GetCapabilities()` bitmask probing so capability detection stays accurate across custom builds and backports.

## References

- `src/AdaskoTheBeAsT.Interop.COM/NanoServerOptions.cs`
- `src/AdaskoTheBeAsT.Interop.COM/NanoServerStaShim.cs`
- `src/AdaskoTheBeAsT.Interop.COM/NanoServerOle2Shim.cs`
- `src/AdaskoTheBeAsT.Interop.COM/NanoServerDispenserManager.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/NanoServerWxOle2DispenserAndStartupTest.cs`
