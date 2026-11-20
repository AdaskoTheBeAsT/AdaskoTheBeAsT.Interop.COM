# ADR-0033: Direct `HMODULE` Class-Factory Instantiation via `ComClassFactory` and Activation-Scope Helpers

- Status: Accepted
- Date: 2026-11-17
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0
- Extends [ADR-0026](0026-activation-scopes-and-derived-icomactivationexecutor.md)

## Context

Even with registration-free SxS manifests or `CoRegisterClassObject`, instantiating a known root coclass via `new MyComClass()` or `CoCreateInstance` traverses thread activation contexts, process apartment tables, or `HKCR\CLSID` registry keys. When a host already holds a loaded worker `HMODULE` (for example via `IComActivationScope.LoadedModuleHandles` with `PreloadModules = true`, or a single utility COM DLL without an external `.manifest`), calling the module's exported `DllGetClassObject` entry point and `IClassFactory::CreateInstance` directly guarantees that the instance originates from that exact `HMODULE` with zero SxS, `HKCR`, or `CoRegisterClassObject` indirection.

However, manual `GetProcAddress` + vtable slot invocation from application code is error-prone: callers frequently leak the intermediate `IClassFactory*` or raw object `IntPtr`, forget to fall back to `IID_IUnknown` when a legacy factory returns `E_NOINTERFACE` for a derived `IID` on `IClassFactory::CreateInstance`, leak `LoadLibraryExW` handles when `CreateInstance` fails, or forget to register created RCWs with an active `IComActivationScope` before `UnloadModulesOnDeactivate = true` unloads the DLL.

## Decision

1. **Introduce `ComClassFactory` following the library's `ComCreateResult<T>` pattern:**
   - `CreateInstance(IntPtr moduleHandle, Guid clsid, Guid? interfaceId = null)` and `CreateInstance<TInterface>(IntPtr moduleHandle, Guid clsid)` resolve `DllGetClassObject` from `moduleHandle`, request `IID_IClassFactory` (falling back to `IID_IUnknown`), invoke `IClassFactory::CreateInstance` (slot `3`, retrying with `IID_IUnknown` if a specific `IID` returns `E_NOINTERFACE`), materialize the RCW via `Marshal.GetObjectForIUnknown`, validate `TInterface` castability (calling `FinalReleaseComObject` if the cast fails), and release all raw native pointers in `finally`.
   - `CreateInstanceFromDll` and `CreateInstanceFromDll<TInterface>` load `dllPath` via the library's safe search-order `LoadLibraryExW` chain (`LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS` fallback chain) and automatically `FreeLibrary` + notify `NanoServerOle2Shim.OnModuleUnloaded` if instantiation fails, returning the loaded `moduleHandle` only on success.
2. **Expose scope-bound `CreateInstance` / `CreateInstance<TInterface>` on `IComActivationScope` and `ComActivationScopeExtensions`:**
   - Enforce creating-thread affinity (`ManagedThreadId` and native `GetCurrentThreadId`), non-released scope state, `PreloadModules = true` availability (`LoadedModuleHandles.Count > 0`), and valid `moduleIndex`, then instantiate from `LoadedModuleHandles[moduleIndex]` and automatically register the resulting RCW via `TrackObject(...)` so `Deactivate` releases it prior to `DllCanUnloadNow()` / `FreeLibrary`.

## Consequences

- Positive: Callers can instantiate root coclasses directly from a specific `HMODULE` without `.manifest` files, registry lookups, or apartment class-table contention, and scope-created RCWs participate automatically in deterministic module-unload tracking.
- Positive: Operational COM `HRESULT` failures (`CLASS_E_CLASSNOTAVAILABLE`, `E_NOINTERFACE`, missing `DllGetClassObject`) return cleanly in `ComCreateResult<T>.Exception` enriched via `ComErrorDiagnostics.CreateExceptionForHr` without leaking native pointers or `HMODULE`s.
- Negative: Internal `CoCreateInstance` calls made *inside* the native COM server for secondary helper coclasses still require an active SxS or fallback activation scope (`Executor.Activate`); direct `ComClassFactory` calls bypass COM's SCM only for the call site itself.

## Alternatives Considered

- Adding new members to `IComExecutor`: rejected to preserve strict binary and source compatibility for existing `IComExecutor` implementations and test mocks ([ADR-0013](0013-icomexecutor-abstraction.md), [ADR-0026](0026-activation-scopes-and-derived-icomactivationexecutor.md)).
- Throwing exceptions on `CLASS_E_CLASSNOTAVAILABLE` or `E_NOINTERFACE`: rejected to stay consistent with the library's `Result` / `ComCreateResult<T>` contract ([ADR-0003](0003-result-pattern-instead-of-exceptions.md)).

## References

- `src/AdaskoTheBeAsT.Interop.COM/ComClassFactory.cs`
- `src/AdaskoTheBeAsT.Interop.COM/IComActivationScope.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComActivationScope.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComActivationScopeExtensions.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/ForceInProcessFallbackAndClassFactoryTest.cs`
