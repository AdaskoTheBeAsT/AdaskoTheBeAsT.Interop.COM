# Migration guide

Upgrade guidance for `AdaskoTheBeAsT.Interop.COM`. Read the [changelog](CHANGELOG.md) for the full change list and the [README](README.md) for setup and deployment.

- [4.x to 5.0](#4x-to-50)
- [3.x to 4.0](#3x-to-40)
- [2.x to 3.0](#2x-to-30)
- [Upgrade verification checklist](#upgrade-verification-checklist)

## 4.x to 5.0

This section covers upgrading from **4.0.x–4.2.0** to **5.0.0**. Existing `Executor` and `IComExecutor` method signatures remain intact, and `IComExecutor` adds no abstract members, so existing decorators and test mocks compile without changes. However, supported target frameworks, pre-activation manifest validation, and uninitialized-thread apartment selection on Nano Server fallback change in 5.0.0.

### 1. Check your target framework

| Your current target | Action for 5.0 |
| --- | --- |
| `net11.0`, `net10.0` | Supported directly; prefer `net11.0-windows` or `net10.0-windows` for Windows-only applications |
| `net472`, `net48`, `net481` | Supported directly; no framework change required |
| `net8.0`, `net9.0` | Retarget to `net10.0` or `net11.0`, or stay on `4.2.0` |

`AdaskoTheBeAsT.Interop.COM` 5.0.0 continues to ship zero external runtime NuGet dependencies across all supported targets.

### 2. Existing `IComExecutor` implementations remain valid; adopt `IComActivationExecutor` or extension methods for scopes

`IComExecutor` is unchanged. `ComExecutor` additionally implements `IComActivationExecutor`, and `ComExecutorExtensions` exposes `Execute(..., ComActivationOptions)`, `Create(..., ComActivationOptions)`, `Activate(...)`, and `Deactivate(...)` on any `IComExecutor`:

```csharp
IComExecutor executor = new ComExecutor();

var options = new ComActivationOptions
{
    PreloadModules = true,
    UnloadModulesOnDeactivate = true,
    PumpPendingMessages = false,
};

var activation = executor.Activate(dll, manifest, options);
if (!activation.Success)
{
    throw new InvalidOperationException("COM scope activation failed.", activation.Exception);
}

using var scope = activation.Value!;
var first = scope.TrackObject(new FirstComClass());
var second = scope.TrackObject(new SecondComClass());

// Use both coclasses in any order on the owning thread...

var release = executor.Deactivate(scope);
if (!release.Success)
{
    throw new InvalidOperationException("COM scope deactivation failed.", release.Exception);
}
```

If you wrap `IComExecutor` in a custom decorator and want options-bearing calls to flow through your decorator rather than falling back to a temporary `ComExecutor`, implement `IComActivationExecutor` on the decorator as well.

### 3. Review apartment initialization on Nano Server fallback

In 4.1.0–4.2.0, Nano Server fallback initialized uninitialized threads with `COINIT_MULTITHREADED` (`MTA`). In 5.0.0:

- If the calling thread already initialized COM (`STA` or `MTA`), the library keeps that apartment unless an explicit `ComActivationOptions.ApartmentState` requests the opposite mode, in which case activation fails with `RPC_E_CHANGED_MODE` (`0x80010106`) rather than silently changing apartment mode.
- If the calling thread has not initialized COM (`CO_E_NOTINITIALIZED`), Nano Server fallback selects `COINIT_APARTMENTTHREADED` (`STA`) when `ComActivationOptions.ApartmentState` is `ApartmentState.STA` or when `ApartmentState.Unknown` and the manifest declares `threadingModel="Apartment"` or `"Single"`; otherwise it selects `COINIT_MULTITHREADED` (`MTA`).
- Only `CoInitializeEx` frames acquired by the library are balanced with `CoUninitialize` on the creating thread.

### 4. Fix invalid or conflicting manifests surfaced by pre-activation validation

5.0.0 validates and normalizes manifests before loading native DLLs or writing registry entries:

- Malformed `clsid`, `tlbid`, or `iid` GUID attributes fail immediately in `Result.Exception`.
- Duplicate CLSIDs, ProgIDs, or TypeLibs with conflicting target DLLs or metadata fail immediately (idempotent identical declarations across nested manifests remain allowed).
- Relative path traversal (`..`), UNC paths (`\\server\share\...`), Win32 device paths (`\\?\...`, `\\.\...`), root-relative paths (`\evil.dll`), and drive-relative paths (`C:evil.dll`) in `<file name="...">` are rejected.

### 5. Opt in explicitly to module unloading, HKCU registration, legacy Nano Server shims, and startup preparation

Default activation remains registry-free and retains loaded native modules for process lifetime so uncollected RCWs cannot trigger finalizer access violations. Opt in only to the behaviors your workload requires:

- **Worker DLL unloading and Desktop SxS bypass:** Set `UnloadModulesOnDeactivate = true` (paired with `PreloadModules = true` on Desktop Windows, or `ForceInProcessFallback = true` when isolating identical `<assemblyIdentity>` worker clones across parallel STA threads without WinSxS cache coalescing) and register created RCWs via `scope.TrackObject(...)` or `scope.CreateInstance<T>(clsid)` so `Deactivate` releases them before checking `DllCanUnloadNow()` and calling `FreeLibrary`. Use `ComModuleDiagnostics.VerifyRegisteredInprocServer(...)` for RCW-free `CLSID` provenance checks so no `GC.Collect()` is required before unloading, and `ComManifestRewriter` when staging renamed worker DLL manifests (`Server_00.dll`) with `<description>` PII stripped.
- **Direct class-factory instantiation, locale guards, error diagnostics, and module quirks:** Use `ComClassFactory` when instantiating a coclass directly from a loaded `HMODULE` or DLL path with zero SxS/registry lookups, `ComAutomationLocaleScope.Enter()` to synchronize Win32 `GetThreadLocale` (`0x0409`) and `CultureInfo.CurrentCulture` around locale-sensitive `oleaut32` calls, `ComErrorDiagnostics` to peek/consume thread `IErrorInfo` or unpack VB6 `0x800Axxxx` error codes, and `ComModuleQuirks.SuppressUninitializeImports(moduleHandle)` only for specific legacy third-party DLLs that emit unbalanced `CoUninitialize` / `OleUninitialize` calls.
- **ProgID / TypeLib HKCU leases on fallback:** Set `RegisterCurrentUserProgIdsOnFallback = true` or `RegisterCurrentUserTypeLibsOnFallback = true` on `ComActivationOptions` only when legacy servers call `CLSIDFromProgID`, `ProgIDFromCLSID`, or `LoadRegTypeLib`. Both forward (`<ProgId>\CLSID`) and reverse (`CLSID\{CLSID}\ProgID` and `VersionIndependentProgID`) leases are reference-counted, pre-existing `HKCU\Software\Classes` values are preserved, and `SkipCurrentUserRegistrationWhenPresentInClassesRoot = true` (default) skips redundant `HKCU` writes when matching entries already exist in `Registry.ClassesRoot`.
- **Older Nano Server images (`< 1.8.0`):** Set `NanoServerOptions.EnableLegacyOle2Shim = true` (`ADASKO_NANOSERVER_ENABLE_LEGACY_OLE2_SHIM=1`) and/or `NanoServerOptions.EnableDispenserManagerFallback = true` (`ADASKO_NANOSERVER_ENABLE_DISPENSER_FALLBACK=1`) before loading legacy OLE2 / MDAC / Jet 4.0 consumers on images whose OS `nanoserver_sta_shim.dll` does not yet advertise those capabilities.
- **Concurrent STA worker pools:** Call `NanoServerOptions.PrepareProcess()` once during single-threaded host startup before spawning parallel STA worker threads (`Executor` also pre-warms lazy CoreCLR native STA `OleMainThreadWndClass` initialization automatically on STA entry points via `NanoServerOptions.EnsureCurrentStaThreadInitialized()`).

## 3.x to 4.0

This section covers upgrading to **4.0.0**. If you are moving directly from 2.x to 4.0, apply both migration sections.

Existing public method signatures remain available, and `IComExecutor` has no new members. That does **not** make this a drop-in upgrade: supported frameworks and observable cleanup, failure, and message-processing behavior change.

### 1. Check your target framework

| Your current target | Action for 4.0 |
| --- | --- |
| `net8.0`, `net9.0`, `net10.0` | No framework change required; use a Windows-qualified target for a Windows-only application |
| `net472`, `net48`, `net481` | No framework change required |
| `net462`, `net47`, `net471` | Retarget to .NET Framework 4.7.2 or later, or stay on a compatible earlier package |
| `netstandard2.0` library | Retarget/multi-target to the supported concrete frameworks, or retain a compatible earlier package for that target |
| A runtime that used the `netstandard2.0` fallback, such as .NET 6 | Retarget to a supported modern .NET target, or stay on a compatible earlier package |

There is no .NET Standard fallback in 4.0. A reusable library with unsupported targets may need conditional package references instead of one package version across all targets. Do not manually copy the 4.0 DLL into an unsupported application.

### 2. Check results, not just exceptions

Operational setup failures, such as missing or malformed manifests, are now captured in `Result.Exception`. A catch block alone will not detect them:

```csharp
var result = executor.Execute(dll, manifest, Work);

if (!result.Success)
{
    throw new InvalidOperationException("COM execution failed.", result.Exception);
}
```

Apply the same check to `Create` and `Free`. Check `creation.Success` before accessing `creation.Value`.

Invalid arguments still throw, including null callbacks, blank paths, empty descriptor collections, and null collection entries. If work and cleanup both fail, inspect the exceptions inside the returned `AggregateException`; do not assume every failure is a single `Win32Exception`.

### 3. Keep handles on their owning thread

Handle release now validates both managed and native thread identity. This is not safe:

```csharp
// Incorrect: the continuation may run on another thread.
var creation = executor.Create(dll, manifest, CreateComObject);
using var handle = creation.Value!;
await DoOtherWorkAsync();
```

Keep the complete create/use/release scope synchronous on the owning thread instead:

```csharp
// Run this entire block on the thread/apartment required by the component.
var creation = executor.Create(dll, manifest, CreateComObject);
if (!creation.Success)
{
    throw new InvalidOperationException("COM creation failed.", creation.Exception);
}

using (var handle = creation.Value
    ?? throw new InvalidOperationException("No COM handle was returned."))
{
    Use(handle.ComObject
        ?? throw new InvalidOperationException("No COM object was returned."));
}

// Other asynchronous work can follow after the COM scope is released.
```

The functions and paths in these snippets represent your application's code. If you use an STA scheduler for a longer-lived object, it must guarantee that every operation, including final release, runs on the **same** thread, not merely on any STA thread.

The library still does not initialize COM or establish an STA. That responsibility was always the caller's; an `[STAThread]` entry point is suitable for a synchronous console sample.

### 4. Release nested handles in reverse order

If you create A and then B, release B before A. A wrong-order release now fails without releasing A's RCW.

Use nested scopes, checking each creation result:

```csharp
var first = executor.Create(dllA, manifestA, CreateFirst);
if (!first.Success)
{
    throw new InvalidOperationException("First COM creation failed.", first.Exception);
}

using (var a = first.Value
    ?? throw new InvalidOperationException("No first handle was returned."))
{
    var second = executor.Create(dllB, manifestB, CreateSecond);
    if (!second.Success)
    {
        throw new InvalidOperationException("Second COM creation failed.", second.Exception);
    }

    using (var b = second.Value
        ?? throw new InvalidOperationException("No second handle was returned."))
    {
        // Use a and b here. Dispose b before leaving a's scope.
    }
}
```

Finish nested `Execute` calls before releasing an outer handle. Dispose handles created inside an `Execute` callback or `Create` factory before that delegate returns. Do not release an outer handle reentrantly from a COM callback.

### 5. Handle incomplete release explicitly

`Dispose` remains non-throwing for release failures and emits `HandleReleaseFailed` (event ID 2). In 4.0, rejected or incomplete disposal no longer suppresses leak diagnostics.

Use `Free` if your application needs to react to a failure:

```csharp
var release = executor.Free(handle);
if (!release.Success)
{
    // Keep the handle reachable for recovery on its owning thread.
    // Log or propagate release.Exception according to application policy.
    // Do not assume handle.ComObject is still usable.
}
```

This is a failure-handling fragment, not a complete ownership implementation. A longer-lived owner should retain its handle in a field until `handle.IsReleased` is true.

| Outcome | State and recovery |
| --- | --- |
| Wrong thread, reentrant release, or wrong activation order | Object remains unchanged; correct the thread/order or finish the callback, then retry |
| Native cleanup fails after the RCW is released | `ComObject` may be null while `IsReleased` is false; keep the handle to retry context cleanup, not to resume COM calls |
| Already released | `Free` returns success without additional work |

Do not blindly loop on failure, drop the handle, or overwrite it by creating a replacement. Retry only after addressing the reported cause. Avoid throwing from a `finally` block in a way that hides an earlier application exception; preserve both failures if needed.

### 6. Audit RCW ownership

Release still uses `Marshal.FinalReleaseComObject`. This is not new, but 4.0 documentation makes the ownership contract explicit:

- A factory must transfer exclusive ownership of its returned RCW.
- Do not return a borrowed object, cached singleton, or RCW already owned by another handle.
- Every alias to the same RCW becomes invalid when the handle releases it.
- Do not call `Marshal.ReleaseComObject` or `Marshal.FinalReleaseComObject` independently on a handle-owned RCW.
- Child RCWs returned by properties or methods are your responsibility; the handle does not recursively release them.
- `Execute` only manages activation contexts. It does not release arbitrary COM objects created inside its callback.

These rules cannot be inferred from managed alias counts. Review object sharing rather than adding reference-count checks.

### 7. Choose who processes messages

Default pumping is still enabled, but each pump now processes at most 256 messages, preserves `WM_QUIT` and its exit code, and uses Unicode APIs on all supported targets. Do not assume a COM operation drains the entire queue.

When an existing UI loop or STA scheduler owns message processing, opt out:

```csharp
IComExecutor executor = new ComExecutor(pumpPendingMessages: false);
```

With DI:

```csharp
services.AddSingleton<IComExecutor>(
    _ => new ComExecutor(pumpPendingMessages: false));
```

With the static API:

```csharp
var execution = Executor.Execute(dll, manifest, Work, pumpPendingMessages: false);
var creation = Executor.Create(dll, manifest, CreateComObject, pumpPendingMessages: false);
```

Check both results as usual. A created handle retains this policy through `Dispose`, `Executor.Free`, or `Free` on a differently configured executor. There is no new member to implement on mocks or custom `IComExecutor` implementations.

The built-in pump bypasses host-specific filters, accelerators, and dialog preprocessing. It can consume custom `PostThreadMessage` messages without delivery. Disabling it prevents only the library's explicit pump; COM can still pump during synchronous calls. Your host must provide whatever message processing the component requires.

### 8. Check deployment and repository tooling

The activation context's assembly directory now uses the directory containing the COM DLL. Test with absolute DLL/manifest paths, especially when the manifest lives elsewhere or private dependencies are involved.

For contributors building this repository:

- Install the SDK selected by `global.json`.
- Use Microsoft.Testing.Platform syntax: `dotnet test --solution AdaskoTheBeAsT.Interop.COM.slnx` or `dotnet test --project <path-to-csproj>`.
- Replace test commands targeting removed frameworks, for example `net462`, with a remaining target such as `net472`.
- Install the required native C++ build tools and architecture-specific runtimes. See [build and test](README.md#build-and-test).

These SDK/test-runner settings apply to this repository; consuming applications do not need to switch their own test runner.

## 2.x to 3.0

Version 3.0.0 retains the static API and adds injectable execution, disposable handles, platform annotations, and AnyCPU packaging. These steps also apply when upgrading directly from 2.x to 4.0, subject to the 4.0 framework removals above.

### Adopt `IComExecutor` when useful

Before:

```csharp
var result = Executor.Execute(dll, manifest, Work);
```

After:

```csharp
IComExecutor executor = new ComExecutor();
var result = executor.Execute(dll, manifest, Work);
```

Application services can receive `IComExecutor` through constructor injection. Register it with `services.AddSingleton<IComExecutor, ComExecutor>()` if using Microsoft.Extensions.DependencyInjection. No separate DI package is shipped by this library.

Existing static calls remain valid; converting all call sites is not required.

### Use disposable handles, but do not discard failure handling

The lifetime API introduced in 2.1 required explicit `Executor.Free(handle)`. In 3.0, `ComObjectHandle<T>` also implements `IDisposable`:

```csharp
var creation = executor.Create(dll, manifest, CreateComObject);
if (!creation.Success)
{
    throw new InvalidOperationException("COM creation failed.", creation.Exception);
}

using var handle = creation.Value
    ?? throw new InvalidOperationException("No COM handle was returned.");
Use(handle.ComObject
    ?? throw new InvalidOperationException("No COM object was returned."));
```

`using` attempts the same release operation, but does not return the failure information exposed by `Free`. Keep explicit `Free` if your application needs that result. Monitor disposal diagnostics when relying on `using`; the finalizer diagnoses leaks rather than cleaning them up.

### Address Windows platform warnings

On modern .NET, public COM entry points are annotated with `[SupportedOSPlatform("windows")]`. Choose the appropriate approach:

1. Target `net8.0-windows`, `net9.0-windows`, or `net10.0-windows` for a Windows-only application.
2. Guard calls with `OperatingSystem.IsWindows()` in a cross-platform host.
3. Annotate a Windows-only calling member or class with `[SupportedOSPlatform("windows")]`.

Do not broadly suppress `CA1416` to imply cross-platform support.

### Pin the host architecture and review package assets

The managed library changed from x86 to AnyCPU. Explicitly configure the consumer as x86 for a 32-bit native component, or deploy an x64 native component and matching manifest for a 64-bit process. Do not rely on the package to make the application run as 32-bit.

3.0 added dedicated .NET Framework assets from 4.6.2 through 4.8.1. `net10.0` was already present in the tagged 2.1.0 project. NuGet selects compatible assets automatically; 4.0 subsequently removes the older targets listed above.

## Upgrade verification checklist

- [ ] Restore against the intended package/build and confirm all application targets are supported.
- [ ] Match the process architecture, COM DLL, manifest, and native dependencies.
- [ ] Activate the real COM class on a clean Windows environment where it is not globally registered.
- [ ] Check `Success` and `Exception` for activation, creation, execution, and release.
- [ ] Exercise failure paths: bad manifest, factory exception, and COM method exception.
- [ ] Verify same-thread use/release and reverse-order disposal for nested handles.
- [ ] Confirm no shared RCWs or independently released handle-owned RCWs.
- [ ] Verify UI/thread-message delivery and shutdown with the chosen pump policy.
- [ ] Retain incompletely released handles and scopes for recovery, without retaining aliases to released RCWs.
- [ ] Check for `HandleLeaked`, `HandleReleaseFailed`, `ScopeLeaked`, and `ScopeReleaseFailed` diagnostics in representative workloads.
- [ ] Call `NanoServerOptions.PrepareProcess()` once at process startup before launching concurrent STA worker pools on headless Nano Server containers.
