# Advanced usage

These examples supplement the [README](../README.md) and target version **5.0.0**. See the [migration guide](../MIGRATION.md) when upgrading.

Unless a complete class is shown, snippets assume your application supplies the paths, COM types, callbacks, DI registrations, and error reporting. Keep all handle operations on their creating thread.

## Multiple activation contexts

Pass an ordered collection of `ComPathDescriptor` instances to `Execute` or `Create`:

```csharp
var descriptors = new List<ComPathDescriptor>
{
    new ComPathDescriptor(@"C:\MyApp\First.dll", @"C:\MyApp\First.manifest"),
    new ComPathDescriptor(@"C:\MyApp\Second.dll", @"C:\MyApp\Second.manifest"),
};

var result = executor.Execute(descriptors, () =>
{
    // Instantiate and use your COM classes while both contexts are active.
    // Release any exclusively owned objects before this callback returns.
});

if (!result.Success)
{
    throw new InvalidOperationException("COM execution failed.", result.Exception);
}
```

Use `System.Collections.Generic` for `List<T>`. Contexts activate in collection order and deactivate in reverse order. If multiple manifests describe the same class, activation-stack precedence matters; avoid relying on ambiguous duplicate declarations.

`Execute` manages activation contexts only. Use `Create` for an exclusively owned RCW that needs a handle, or `Activate` (`IComActivationScope`) when multiple coclasses share one session lifetime, and follow the [lifetime rules](../README.md#lifetime-and-threading-rules).

## Multi-coclass activation scopes and isolated worker recycling

When one calculation or workflow instantiates several coclasses across one or more manifests — or recycles per-worker COM DLL copies in isolated directories to reset mutable native `.data`/`.bss` state — use `Activate` / `Deactivate` with `ComActivationOptions`:

```csharp
var options = new ComActivationOptions
{
    PreloadModules = true,
    UnloadModulesOnDeactivate = true,
    PumpPendingMessages = false,
};

var activation = executor.Activate(descriptors, options);
if (!activation.Success)
{
    throw new InvalidOperationException("Scope activation failed.", activation.Exception);
}

using var scope = activation.Value!;
var engine = scope.TrackObject(new CalcEngineClass());
var formatter = scope.TrackObject(new CalcFormatterClass());

// Verify the live COM object's direct native interface resolves to the isolated worker DLL
// rather than a stray global HKLM registration:
if (scope.LoadedModuleHandles.Count > 0)
{
    ComModuleDiagnostics.VerifyInstanceModule(
        engine,
        scope.LoadedModuleHandles[0],
        expectedDllPath: descriptors[0].ComAssemblyPath,
        interfaceId: typeof(ICalcEngine).GUID);
}

// Perform cross-coclass work on the owning thread...

var release = executor.Deactivate(scope);
if (!release.Success)
{
    throw new InvalidOperationException("Scope deactivation failed.", release.Exception);
}
```

During `Deactivate`, tracked RCWs are released in reverse registration order via `Marshal.FinalReleaseComObject`, activation contexts / class-factory registrations are torn down, `CoFreeUnusedLibrariesEx(0, 0)` runs, and each module whose lease count reaches zero is freed via `FreeLibrary` only if `DllCanUnloadNow()` returns `S_OK` (`0`). If an untracked native reference temporarily reports `S_FALSE`, the module stays mapped safely until a later `Executor.TryUnloadUnusedModules()` sweep.

### Staging renamed worker manifests, Desktop SxS bypass, RCW-free probing, and direct `HMODULE` instantiation

When staging dedicated per-worker directories (`worker_00\CalcEngine_00.dll`) across parallel STA threads on Desktop Windows, Windows Server Core, or Nano Server:

1. **Rewrite and sanitize staged SxS manifests idempotently (`ComManifestRewriter`):**

```csharp
bool wrote = ComManifestRewriter
    .FromFile(templateManifestPath)
    .WithAssemblyIdentityName("CalcEngine.Worker00")
    .RenameFile("CalcEngine.dll", "CalcEngine_00.dll")
    .RemoveFile("OptionalDesignerAddin.dll")
    .StripDescriptions()
    .SaveIfChanged(workerManifestPath);
```

2. **Bypass Desktop WinSxS identity caching (`ForceInProcessFallback = true`), probe `CLSID` provenance without allocating a managed RCW, or instantiate directly via `DllGetClassObject` (`scope.CreateInstance<T>` / `ComClassFactory`):**

```csharp
var workerOptions = new ComActivationOptions
{
    ForceInProcessFallback = true,
    PreloadModules = true,
    UnloadModulesOnDeactivate = true,
    PumpPendingMessages = false,
};

var workerActivation = executor.Activate(workerDllPath, workerManifestPath, workerOptions);
if (!workerActivation.Success)
{
    throw new InvalidOperationException("Worker activation failed.", workerActivation.Exception);
}

using var workerScope = workerActivation.Value!;

// Pre-flight CoCreateInstance provenance check with ZERO managed RCW allocations
// (so Deactivate + FreeLibrary never requires GC.Collect()):
ComModuleDiagnostics.VerifyRegisteredInprocServer(
    typeof(CalcEngineClass).GUID,
    workerScope.LoadedModuleHandles[0],
    expectedDllPath: workerDllPath,
    interfaceId: typeof(ICalcEngine).GUID);

// Direct DllGetClassObject -> IClassFactory::CreateInstance from LoadedModuleHandles[0],
// automatically registered with workerScope.TrackObject(...):
var directCreate = workerScope.CreateInstance<ICalcEngine>(typeof(CalcEngineClass).GUID);
if (!directCreate.Success)
{
    throw new InvalidOperationException("Direct coclass creation failed.", directCreate.Exception);
}
```

## Automation `LCID` guard, `IErrorInfo` diagnostics, and per-module `CoUninitialize` quirks

Legacy VB6, Delphi, and Office/OLE Automation servers rely on thread-local Win32 state that managed `CultureInfo` or default CLR exception marshaling does not normalize automatically:

```csharp
// 1. Force Win32 GetThreadLocale() = 0x0409 (en-US) for oleaut32 VarR8FromStr / IDispatch
//    while setting CultureInfo.CurrentCulture = InvariantCulture on the owning thread:
using (ComAutomationLocaleScope.Enter())
{
    int hr = nativeEngine.CalculateRate("1234.56", out double rate);
    if (hr < 0)
    {
        // 2. Capture rich thread IErrorInfo (Description, Source, HelpFile, HelpContext, IID)
        //    and unpack VB6 FACILITY_CONTROL (0x800Axxxx) error numbers:
        if (ComErrorDiagnostics.TryGetVisualBasicErrorNumber(hr, out int vbErr))
        {
            // e.g. vbErr == 13 for VB6 Type Mismatch (0x800A000D)
        }

        throw ComErrorDiagnostics.CreateExceptionForHr(hr);
    }
}

// 3. Opt-in per-HMODULE W^X IAT guard ONLY for defective legacy DLLs that emit an extra
//    unbalanced CoUninitialize() / OleUninitialize() and would otherwise tear down the STA apartment:
ComModuleQuirks.SuppressUninitializeImports(workerScope.LoadedModuleHandles[0]);
```

## VB6 1-based and Delphi borrowed `SAFEARRAY` marshaling

Use `SafeArrayMarshal` when interoperating with VB6 non-zero lower-bound arrays or Borland Delphi raw `SAFEARRAY*` pointers lacking `FADF_HAVEVARTYPE`:

```csharp
// 1. Copy an unmanaged numeric array from a caller-owned/borrowed Delphi SAFEARRAY* without destroying it:
int[] borrowedValues = SafeArrayMarshal.CopyFromBorrowedSafeArray<int>(borrowedSafeArrayPtr);

// 2. Copy and deterministically SafeArrayDestroy a callee-allocated SAFEARRAY*:
double[] ownedValues = SafeArrayMarshal.ConsumeOwnedSafeArray<double>(ownedSafeArrayPtr);

// 3. Normalize a 1-based VB6 System.Array (or VT_VARIANT array) returned as object into a 0-based T[]:
double[] normalized = SafeArrayMarshal.NormalizeVariantArray<double>(vb6VariantResult);

// 4. Create a 1-based SAFEARRAY-compatible System.Array for VB6 ByRef array inputs:
Array vb6Input = SafeArrayMarshal.CreateOneBasedArray(new[] { 10.0, 20.0, 30.0 });
```

Ownership-aware `CopyBorrowedBstrArray` / `ConsumeOwnedBstrArray` and `CopyBorrowedVariantArray` / `ConsumeOwnedVariantArray` helpers are also available for string and `VARIANT` payloads.

## Process startup preparation for concurrent STA worker pools

On headless Windows Nano Server containers, call `NanoServerOptions.PrepareProcess()` once during single-threaded host startup before spawning concurrent STA worker threads:

```csharp
// Optional compatibility shims for older Nano Server container images (< 1.8.0):
// NanoServerOptions.EnableLegacyOle2Shim = true;
// NanoServerOptions.EnableDispenserManagerFallback = true;

NanoServerOptions.PrepareProcess();
```

This ensures `%SystemRoot%\System32\forwarders` (`SysWOW64\forwarders` on x86) `PATH` setup, base DLL preloading, W^X IAT patching, and the background `NanoServerMainStaHost` thread are initialized before parallel workers start.

## Dependency injection and decorators

`ComExecutor` is sealed. Wrap `IComExecutor` to add logging or metrics without changing the execution and ownership rules.

For example, this decorator logs failed operations. The consuming project needs Microsoft.Extensions.Logging abstractions:

```csharp
using System;
using System.Collections.Generic;
using AdaskoTheBeAsT.Interop.COM;
using Microsoft.Extensions.Logging;

public sealed class LoggingComExecutor : IComExecutor
{
    private readonly IComExecutor _inner;
    private readonly ILogger<LoggingComExecutor> _logger;

    public LoggingComExecutor(
        IComExecutor inner,
        ILogger<LoggingComExecutor> logger)
    {
        _inner = inner;
        _logger = logger;
    }

    public Result Execute(string comAssemblyPath, string manifestPath, Action action)
        => Report(_inner.Execute(comAssemblyPath, manifestPath, action));

    public Result Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action)
        => Report(_inner.Execute(comPathDescriptors, action));

    public ComObjectCreationResult<T> Create<T>(
        string comAssemblyPath, string manifestPath, Func<T> factory)
        where T : class
        => Report(_inner.Create(comAssemblyPath, manifestPath, factory));

    public ComObjectCreationResult<T> Create<T>(
        ICollection<ComPathDescriptor> comPathDescriptors, Func<T> factory)
        where T : class
        => Report(_inner.Create(comPathDescriptors, factory));

    public Result Free<T>(ComObjectHandle<T> comObjectHandle)
        where T : class
        => Report(_inner.Free(comObjectHandle));

    private TResult Report<TResult>(TResult result)
        where TResult : Result
    {
        if (!result.Success)
        {
            _logger.LogError(result.Exception, "COM operation failed.");
        }

        return result;
    }
}
```

Register the implementation and decorator separately:

```csharp
services.AddLogging();
services.AddSingleton<ComExecutor>();
services.AddSingleton<IComExecutor>(provider =>
    new LoggingComExecutor(
        provider.GetRequiredService<ComExecutor>(),
        provider.GetRequiredService<ILogger<LoggingComExecutor>>()));
```

When the host owns the message loop, replace the concrete registration with:

```csharp
services.AddSingleton(_ => new ComExecutor(pumpPendingMessages: false));
```

The decorator does not catch argument-validation exceptions. It also cannot intercept `handle.Dispose()`, which invokes the static release path directly. Use event-source diagnostics for disposal failures and leaks.

For metrics, apply the same pattern using `System.Diagnostics.Metrics` in your application. Avoid high-cardinality or sensitive tags such as arbitrary manifest paths. The library itself does not publish a `Meter`.

### Keyed registrations

With Microsoft.Extensions.DependencyInjection 8.0 or later, keyed services can distinguish workloads or pump policies:

```csharp
services.AddKeyedSingleton<IComExecutor, ComExecutor>("background");
services.AddKeyedSingleton<IComExecutor>(
    "ui",
    (_, _) => new ComExecutor(pumpPendingMessages: false));
```

Store per-component paths separately, for example as keyed options or `ComPathDescriptor` instances. An executor contains a pump policy, not DLL paths or COM objects.

## Unit testing without COM

Substitute `IComExecutor` to test delegation and result handling. For example, given this application service:

```csharp
public sealed class ComWorkRunner
{
    private readonly IComExecutor _executor;

    public ComWorkRunner(IComExecutor executor) => _executor = executor;

    public void Run(string dll, string manifest, Action work)
    {
        var result = _executor.Execute(dll, manifest, work);
        if (!result.Success)
        {
            throw new InvalidOperationException("COM work failed.", result.Exception);
        }
    }
}
```

A Moq/xUnit test can verify failure propagation without activating COM:

```csharp
var executor = new Mock<IComExecutor>();
var failure = new InvalidOperationException("Simulated activation failure.");
executor
    .Setup(e => e.Execute(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<Action>()))
    .Returns(new Result { Success = false, Exception = failure });

var runner = new ComWorkRunner(executor.Object);
Action work = () => throw new InvalidOperationException("The mock must not run this.");

var error = Assert.Throws<InvalidOperationException>(
    () => runner.Run("component.dll", "component.manifest", work));

Assert.Same(failure, error.InnerException);
executor.Verify(e => e.Execute("component.dll", "component.manifest", work), Times.Once);
```

Import `Moq` and `Xunit` in the test. Only invoke a captured callback in a unit test when the callback is itself COM-free. Mocking the executor does not replace COM constructors inside that callback.

For success-path business logic that calls COM methods, introduce an application-specific abstraction, such as a string-concatenation service. Mock that abstraction in unit tests and test its real COM implementation separately with the native DLL and manifest.

## STA scheduling and asynchronous callers

`IComExecutor` is synchronous. It does not create a thread, initialize COM, enforce a timeout, or cancel an in-flight COM call.

For applications using an STA scheduler, including a compatible version of the optional [`AdaskoTheBeAsT.Interop.Threading`](https://www.nuget.org/packages/AdaskoTheBeAsT.Interop.Threading/) package:

1. Schedule the entire create/use/release scope on one STA thread for a one-off operation.
2. For a persistent handle, verify that every scheduled operation returns to the exact same thread.
3. Serialize operations on that handle and release all newer handles before disposing it.
4. Keep the owning thread alive until release succeeds. Do not cancel the final cleanup merely because a caller canceled its request.
5. Disable the COM executor's built-in pump if the scheduler owns message processing.
6. Keep a handle after incomplete release, but clear any alias once `handle.ComObject` is null.

Check your scheduler's own API and timeout guarantees. Timing out the caller's wait does not necessarily stop a native COM call. An `IAsyncDisposable` wrapper is safe only if it marshals final release back to the owner thread and retains state when cleanup fails.
