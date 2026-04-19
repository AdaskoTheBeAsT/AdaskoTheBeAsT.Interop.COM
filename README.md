# AdaskoTheBeAsT.Interop.COM
# AdaskoTheBeAsT.Interop.COM

> Registration-free COM interop for .NET — skip `regsvr32`, skip the registry, skip the drama.

[![NuGet](https://img.shields.io/nuget/v/AdaskoTheBeAsT.Interop.COM.svg?label=AdaskoTheBeAsT.Interop.COM&logo=nuget)](https://www.nuget.org/packages/AdaskoTheBeAsT.Interop.COM/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/AdaskoTheBeAsT.Interop.COM.svg?logo=nuget)](https://www.nuget.org/packages/AdaskoTheBeAsT.Interop.COM/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
![TFMs](https://img.shields.io/badge/TFMs-net10.0%20%7C%20net9.0%20%7C%20net8.0%20%7C%20net4.6.2%E2%80%93net4.8.1%20%7C%20netstandard2.0-512BD4?logo=dotnet)
![Windows](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows)
![Warnings](https://img.shields.io/badge/warnings--as--errors-on-green)
![Deterministic](https://img.shields.io/badge/deterministic%20build-on-blue)
![Tests](https://img.shields.io/badge/tests-567%20across%209%20TFMs-brightgreen)
![Coverage](https://img.shields.io/badge/local%20coverage-92.8%25-brightgreen)

### 🔬 Code quality — SonarCloud

> Badges reflect the shared umbrella project `AdaskoTheBeAsT_AdaskoTheBeAsT.Interop` which hosts the combined analysis for every `AdaskoTheBeAsT.Interop.*` library in the monorepo.

[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop)
[![Coverage](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=coverage)](https://sonarcloud.io/component_measures?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=coverage)
[![Maintainability Rating](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=sqale_rating)](https://sonarcloud.io/component_measures?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=sqale_rating)
[![Reliability Rating](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=reliability_rating)](https://sonarcloud.io/component_measures?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=reliability_rating)
[![Security Rating](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=security_rating)](https://sonarcloud.io/component_measures?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=security_rating)
[![Bugs](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=bugs)](https://sonarcloud.io/component_measures?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=bugs)
[![Vulnerabilities](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=vulnerabilities)](https://sonarcloud.io/component_measures?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=vulnerabilities)
[![Code Smells](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=code_smells)](https://sonarcloud.io/component_measures?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=code_smells)
[![Duplicated Lines (%)](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=duplicated_lines_density)](https://sonarcloud.io/component_measures?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=duplicated_lines_density)
[![Technical Debt](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=sqale_index)](https://sonarcloud.io/component_measures?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=sqale_index)
[![Lines of Code](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=ncloc)](https://sonarcloud.io/component_measures?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop&metric=ncloc)

---

## 👋 Hello, COM-wrangler

So you've got a COM component. Maybe it's a vendor ActiveX control from 2005 that still refuses to die. Maybe it's an in-house ATL library that your colleague shipped and then went on sabbatical. Maybe it's a third-party SDK that cheerfully assumes `regsvr32` ran as Administrator on every machine that will ever exist. 🙃

All you wanted to do was call a single method on it. From a .NET process. Ideally without:

- 🚫 polluting the Windows registry
- 🔐 requiring Administrator rights to install
- 🧟 breaking when another version of the same COM component is registered globally
- 🎲 guessing whether your thread is in the right apartment
- 💀 leaking activation contexts because the API that frees them is `extern "C" void WINAPI` and nobody documented what happens if you forget

`AdaskoTheBeAsT.Interop.COM` is the small, focused library that turns *all of that* into a single call:

```csharp
executor.Execute(comDllPath, manifestPath, () =>
{
    var obj = new YourCom.SomeClass();
    obj.DoTheThing();
});
```

The activation context is created from your manifest, pushed on the current thread, your `Action` runs in a real STA with message pumping, and everything is torn down cleanly even when you throw. ✨

---

## ✨ Why you'll love this

- 🚀 **No COM registration ever.** `regsvr32` stays uninstalled. No admin rights. Your build server thanks you.
- 📜 **Manifest-based side-by-side activation.** Standard Windows SxS — battle-tested since Windows XP SP2, used by every in-box OS component.
- 🧵 **Real STA, real message pump.** `CreateActCtx` → `ActivateActCtx` → `CoInitialize(STA)` → run work → `PeekMessage`/`TranslateMessage`/`DispatchMessage` pump → `DeactivateActCtx` → `ReleaseActCtx`. You just write the inner `Action`. ([ADR-0011](docs/adr/0011-pump-sta-messages-after-com-calls.md))
- 🧹 **`ComObjectHandle<T>` is `IDisposable`.** `using var handle = creation.Value!;` — that's your whole cleanup story. Forget to dispose? A diagnostic-only finalizer emits a `HandleLeaked` event on an `EventSource` so you find the bug instead of crashing later. ([ADR-0012](docs/adr/0012-com-object-handle-idisposable.md), [ADR-0018](docs/adr/0018-eventsource-for-leaked-handles.md))
- 🧩 **Drop-in DI.** `services.AddSingleton<IComExecutor, ComExecutor>();` and inject `IComExecutor` anywhere. Unit tests replace it with a `Mock<IComExecutor>` in one line. ([ADR-0013](docs/adr/0013-icomexecutor-abstraction.md))
- 🖥️ **10 TFMs, all green.** `net10.0`, `net9.0`, `net8.0`, `net481`, `net48`, `net472`, `net471`, `net47`, `net462`, `netstandard2.0` — `TreatWarningsAsErrors=true` on every cell. ([ADR-0007](docs/adr/0007-multi-target-frameworks.md))
- 🏛️ **AnyCPU library.** Consumers no longer force 32-bit; a single NuGet works for both x86 and x64 host processes. ([ADR-0016](docs/adr/0016-remove-x86-platform-target.md))
- 🪟 **`[SupportedOSPlatform("windows")]` on the public surface.** Static analysis catches cross-platform call-sites at compile time on TFMs that understand the attribute. ([ADR-0014](docs/adr/0014-supported-os-platform-windows.md))
- 🔭 **Built-in observability.** `EventSource` named `AdaskoTheBeAsT.Interop.COM` emits `HandleLeaked` (Event ID `1`). Subscribe with `dotnet-trace --providers AdaskoTheBeAsT.Interop.COM` or an in-process `EventListener`.
- ✏️ **Source Link + snupkg.** F11 steps into this library from your debugger. No guessing which version is deployed.
- 📚 **19 ADRs** documenting every meaningful design choice.
- 🧪 **567 test invocations** (63 tests × 9 TFMs) plus **92.8 % local line coverage** of the shipped assembly. 9 of 10 source files at 100 %.

---

## 📦 Package

| Package | What it gives you |
| --- | --- |
| [`AdaskoTheBeAsT.Interop.COM`](https://www.nuget.org/packages/AdaskoTheBeAsT.Interop.COM/) | ⚓ `IComExecutor` + `ComExecutor`, static `Executor`, `ComObjectHandle<T>`, `ComPathDescriptor`, `ComObjectCreationResult<T>`, `Result`, `ComInteropEventSource`. One assembly, zero runtime dependencies beyond the BCL. |

### ⬇️ Install

```powershell
dotnet add package AdaskoTheBeAsT.Interop.COM
```

```powershell
# or, via the Package Manager Console
Install-Package AdaskoTheBeAsT.Interop.COM
```

Symbols ship as `.snupkg` with Source Link and embedded untracked sources. Step in. Look around. It's fine.

---

## 🗺️ Target framework matrix

| TFM | Status | Notes |
| --- | :-: | --- |
| `net10.0` | ✅ | Primary target; `LibraryImport` source-gen P/Invoke, `[SupportedOSPlatform]`, `AllowUnsafeBlocks`. |
| `net9.0` | ✅ | Primary target. |
| `net8.0` | ✅ | Primary target. |
| `net481` | ✅ | Windows desktop. |
| `net48` | ✅ | Windows desktop. |
| `net472` | ✅ | Windows desktop. |
| `net471` | ✅ | Windows desktop. |
| `net47` | ✅ | Windows desktop. |
| `net462` | ✅ | Windows desktop. |
| `netstandard2.0` | ✅ | Classic fallback; `[SupportedOSPlatform]` downgrades to no-op. |

Every cell is built with `TreatWarningsAsErrors=true`, `ContinuousIntegrationBuild=true`, `Deterministic=true`, and exercised by CI. Consumers on a single TFM always win via NuGet TFM precedence — you don't have to know this matrix exists.

---

## 🚀 Quick Start

> **Recommended API:** `IComExecutor` (implemented by `ComExecutor`). Inject it into your types and it will be trivially testable. The static `Executor` class remains available for backward compatibility with v2.x code, but **all new code should use `IComExecutor`**.

### Register `ComExecutor` Once

```csharp
using AdaskoTheBeAsT.Interop.COM;
using Microsoft.Extensions.DependencyInjection;

var services = new ServiceCollection();
services.AddSingleton<IComExecutor, ComExecutor>();
// ... resolve IComExecutor wherever you need to run COM code.
```

### Basic Usage (recommended)

```csharp
using AdaskoTheBeAsT.Interop.COM;

public sealed class MyComRunner
{
    private readonly IComExecutor _executor;

    public MyComRunner(IComExecutor executor) => _executor = executor;

    public void Run(string comDllPath, string manifestPath)
    {
        var result = _executor.Execute(comDllPath, manifestPath, () =>
        {
            var comObject = new MyCom.MyComClass();
            var output = comObject.DoSomething();
            Console.WriteLine(output);
        });

        if (!result.Success)
        {
            Console.WriteLine($"Error: {result.Exception?.Message}");
        }
    }
}
```

### Basic Usage (static, legacy — v2.x compatible)

Use this form only if you have existing code that depends on the static entry point; no new code should do this.

```csharp
using AdaskoTheBeAsT.Interop.COM;

var result = Executor.Execute(comDllPath, manifestPath, () =>
{
    var comObject = new MyCom.MyComClass();
    Console.WriteLine(comObject.DoSomething());
});
```

### Real-World Example

```csharp
using AdaskoTheBeAsT.Interop.COM;

public sealed class ComStringProcessor
{
    private readonly IComExecutor _executor;

    public ComStringProcessor(IComExecutor executor) => _executor = executor;

    public string ConcatenateStrings(string str1, string str2)
    {
        var comDllPath = Path.Combine(AppContext.BaseDirectory, "NativeCOM.dll");
        var manifestPath = Path.Combine(AppContext.BaseDirectory, "NativeCOM.manifest");

        string? result = null;

        var executionResult = _executor.Execute(comDllPath, manifestPath, () =>
        {
            var concatenator = new NativeCOM.StringConcatenatorClass();
            result = concatenator.ConcatStrings(str1, str2);
        });

        if (!executionResult.Success)
        {
            throw new InvalidOperationException(
                "COM execution failed",
                executionResult.Exception);
        }

        return result ?? string.Empty;
    }
}
```

#### Unit-testing `ComStringProcessor`

Because `IComExecutor` is an interface, the class above is trivial to test without any real COM on the box:

```csharp
var executor = new Mock<IComExecutor>();
executor
    .Setup(e => e.Execute(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<Action>()))
    .Callback<string, string, Action>((_, _, action) => action())
    .Returns(new Result { Success = true });

var sut = new ComStringProcessor(executor.Object);
// assert against sut.ConcatenateStrings(...) without loading the COM DLL.
```

### Create a COM Object and Release It Later

Use `IComExecutor.Create(...)` when the COM object should outlive a single callback and you want to release it explicitly.

#### Recommended: `using` statement with `IComExecutor`

`ComObjectHandle<T>` implements `IDisposable`, so the simplest way to guarantee release is a `using` block:

```csharp
using AdaskoTheBeAsT.Interop.COM;

public sealed class ComSessionRunner
{
    private readonly IComExecutor _executor;

    public ComSessionRunner(IComExecutor executor) => _executor = executor;

    public void Run(string comDllPath, string manifestPath)
    {
        var creation = _executor.Create(
            comDllPath,
            manifestPath,
            () => new NativeCOM.StringConcatenatorClass());

        if (!creation.Success)
        {
            throw new InvalidOperationException("COM object creation failed.", creation.Exception);
        }

        using var handle = creation.Value
            ?? throw new InvalidOperationException("The COM handle was not created.");

        var concatenator = handle.ComObject
            ?? throw new InvalidOperationException("The COM object was not created.");

        var output = concatenator.ConcatStrings("Hello", "World!");
        Console.WriteLine(output);
        // handle.Dispose() is called automatically at the end of scope,
        // which in turn releases the COM object and activation contexts.
    }
}
```

#### Explicit `IComExecutor.Free(...)` (when you cannot use `using`)

If the handle has to outlive a local scope, call `Free` explicitly:

```csharp
var creation = _executor.Create(
    comDllPath,
    manifestPath,
    () => new NativeCOM.StringConcatenatorClass());

if (!creation.Success)
{
    throw new InvalidOperationException("COM object creation failed.", creation.Exception);
}

var handle = creation.Value
    ?? throw new InvalidOperationException("The COM handle was not created.");

try
{
    var concatenator = handle.ComObject
        ?? throw new InvalidOperationException("The COM object was not created.");

    var output = concatenator.ConcatStrings("Hello", "World!");
    Console.WriteLine(output);
}
finally
{
    var release = _executor.Free(handle);

    if (!release.Success)
    {
        throw new InvalidOperationException("COM object release failed.", release.Exception);
    }
}
```

> The static equivalents (`Executor.Create`, `Executor.Free`) still work unchanged for existing v2.x code.

### Multiple COM Contexts

When you need to work with multiple COM components simultaneously:

```csharp
var descriptors = new List<ComPathDescriptor>
{
    new ComPathDescriptor(@"C:\MyApp\ComLib1.dll", @"C:\MyApp\ComLib1.manifest"),
    new ComPathDescriptor(@"C:\MyApp\ComLib2.dll", @"C:\MyApp\ComLib2.manifest"),
};

var result = _executor.Execute(descriptors, () =>
{
    // Both COM libraries are now active
    var obj1 = new ComLib1.MyClass();
    var obj2 = new ComLib2.AnotherClass();

    // Use both objects together
    var data = obj1.GetData();
    obj2.ProcessData(data);
});
```

### Using with AdaskoTheBeAsT.Interop.Threading

When you also need a dedicated STA thread, timeout handling, or a reusable STA scheduler, combine this package with `AdaskoTheBeAsT.Interop.Threading`.

```bash
dotnet add package AdaskoTheBeAsT.Interop.Threading
```

#### One-Off COM Call with Timeout

```csharp
using AdaskoTheBeAsT.Interop.COM;
using AdaskoTheBeAsT.Interop.Threading;

public sealed class TimedComStringProcessor
{
    private readonly IComExecutor _executor;
    private readonly string _comDllPath;
    private readonly string _manifestPath;

    public TimedComStringProcessor(IComExecutor executor, string comDllPath, string manifestPath)
    {
        _executor = executor;
        _comDllPath = comDllPath;
        _manifestPath = manifestPath;
    }

    public Task<string> ConcatAsync(string left, string right, CancellationToken cancellationToken)
        => SingleThreadedApartmentTask.RunWithTimeoutAsync(
            TimeSpan.FromSeconds(10),
            () =>
            {
                string output = string.Empty;

                var execution = _executor.Execute(_comDllPath, _manifestPath, () =>
                {
                    var concatenator = new NativeCOM.StringConcatenatorClass();
                    output = concatenator.ConcatStrings(left, right);
                });

                if (!execution.Success)
                {
                    throw new InvalidOperationException("The COM call failed.", execution.Exception);
                }

                return output;
            },
            cancellationToken);
}
```

#### Reuse One COM Object on a Single STA Thread

```csharp
using AdaskoTheBeAsT.Interop.COM;
using AdaskoTheBeAsT.Interop.Threading;

public sealed class ScheduledComStringProcessor : IAsyncDisposable
{
    private readonly IComExecutor _executor;
    private readonly string _comDllPath;
    private readonly string _manifestPath;
    private ComObjectHandle<NativeCOM.StringConcatenatorClass>? _handle;
    private NativeCOM.StringConcatenatorClass? _concatenator;

    public ScheduledComStringProcessor(IComExecutor executor, string comDllPath, string manifestPath)
    {
        _executor = executor;
        _comDllPath = comDllPath;
        _manifestPath = manifestPath;
    }

    public Task InitializeAsync(CancellationToken cancellationToken)
        => SingleThreadedApartmentTaskScheduler.RunAsync(
            () =>
            {
                var creation = _executor.Create(
                    _comDllPath,
                    _manifestPath,
                    () => new NativeCOM.StringConcatenatorClass());

                if (!creation.Success)
                {
                    throw new InvalidOperationException("Failed to create the COM object.", creation.Exception);
                }

                _handle = creation.Value
                    ?? throw new InvalidOperationException("The COM handle was not created.");
                _concatenator = _handle.ComObject
                    ?? throw new InvalidOperationException("The COM object was not created.");

                return 0;
            },
            cancellationToken);

    public Task<string> ConcatAsync(string left, string right, CancellationToken cancellationToken)
        => SingleThreadedApartmentTaskScheduler.RunAsync(
            () =>
            {
                if (_concatenator is null)
                {
                    throw new InvalidOperationException("The COM object is not initialized.");
                }

                return _concatenator.ConcatStrings(left, right);
            },
            cancellationToken);

    public ValueTask DisposeAsync()
        => new(
            SingleThreadedApartmentTaskScheduler.RunAsync(
                () =>
                {
                    if (_handle is not null)
                    {
                        var release = _executor.Free(_handle);

                        _concatenator = null;
                        _handle = null;

                        if (!release.Success)
                        {
                            throw new InvalidOperationException("Failed to release the COM object.", release.Exception);
                        }
                    }

                    return 0;
                },
                CancellationToken.None));
}
```

## 🧩 Dependency Injection

This library does **not** ship a separate `*.DependencyInjection` NuGet package (see [ADR-0017](docs/adr/0017-no-dependency-injection-package.md)). `ComExecutor` is stateless and needs no configuration, so the canonical wiring is a one-liner using `Microsoft.Extensions.DependencyInjection`. Logging, metrics, and named/keyed resolution are achieved by composing `IComExecutor` with standard DI building blocks — no library-specific helpers required.

### Register as a singleton (canonical)

```csharp
using AdaskoTheBeAsT.Interop.COM;
using Microsoft.Extensions.DependencyInjection;

var services = new ServiceCollection();
services.AddSingleton<IComExecutor, ComExecutor>();

// Consume:
public sealed class MyService
{
    private readonly IComExecutor _executor;
    public MyService(IComExecutor executor) => _executor = executor;
}
```

`ComExecutor` holds no state, so a singleton is always correct. There is no reason to register it as scoped or transient.

### Register as a keyed singleton (.NET 8+)

When an application has several logical COM workloads — different manifests, different DLLs, different configuration — keyed registrations let each component resolve "its" executor by a string key. This is only meaningful when you also keep per-key `ComPathDescriptor` data somewhere (e.g. keyed options, a dictionary), because `ComExecutor` itself carries no state.

```csharp
services.AddKeyedSingleton<IComExecutor, ComExecutor>("primary");
services.AddKeyedSingleton<IComExecutor, ComExecutor>("reporting");

public sealed class ReportGenerator
{
    private readonly IComExecutor _executor;

    public ReportGenerator([FromKeyedServices("reporting")] IComExecutor executor)
        => _executor = executor;
}

// Manual resolution:
var primary = serviceProvider.GetRequiredKeyedService<IComExecutor>("primary");
```

If you also want to inject the paths/manifests by key, pair the keyed executor with a keyed `ComPathDescriptor`:

```csharp
services.AddKeyedSingleton<IComExecutor, ComExecutor>("reporting");
services.AddKeyedSingleton(
    "reporting",
    new ComPathDescriptor(@"C:\COM\Reporting.dll", @"C:\COM\Reporting.manifest"));
```

### Add logging via a decorator

`ComExecutor` is `sealed`. To add cross-cutting concerns, wrap it with your own `IComExecutor` decorator — this is idiomatic .NET DI composition and needs no library-specific plumbing:

```csharp
using Microsoft.Extensions.Logging;

public sealed class LoggingComExecutor : IComExecutor
{
    private readonly IComExecutor _inner;
    private readonly ILogger<LoggingComExecutor> _logger;

    public LoggingComExecutor(IComExecutor inner, ILogger<LoggingComExecutor> logger)
    {
        _inner = inner;
        _logger = logger;
    }

    public Result Execute(string comAssemblyPath, string manifestPath, Action action)
    {
        _logger.LogDebug("COM Execute starting for {Manifest}", manifestPath);
        var result = _inner.Execute(comAssemblyPath, manifestPath, action);
        if (!result.Success)
        {
            _logger.LogError(result.Exception, "COM Execute failed for {Manifest}", manifestPath);
        }
        return result;
    }

    public Result Execute(ICollection<ComPathDescriptor> descriptors, Action action)
        => _inner.Execute(descriptors, action);

    public ComObjectCreationResult<T> Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory)
        where T : class
        => _inner.Create(comAssemblyPath, manifestPath, factory);

    public ComObjectCreationResult<T> Create<T>(ICollection<ComPathDescriptor> descriptors, Func<T> factory)
        where T : class
        => _inner.Create(descriptors, factory);

    public Result Free<T>(ComObjectHandle<T> handle)
        where T : class
        => _inner.Free(handle);
}
```

Register the decorator. Without Scrutor:

```csharp
services.AddSingleton<ComExecutor>();
services.AddSingleton<IComExecutor>(sp =>
    new LoggingComExecutor(
        sp.GetRequiredService<ComExecutor>(),
        sp.GetRequiredService<ILogger<LoggingComExecutor>>()));
```

With [Scrutor](https://www.nuget.org/packages/Scrutor):

```csharp
services.AddSingleton<IComExecutor, ComExecutor>();
services.Decorate<IComExecutor, LoggingComExecutor>();
```

### Publish metrics via `System.Diagnostics.Metrics.Meter`

The decorator pattern composes just as cleanly with `Meter` / `IMeterFactory` (available in .NET 8+, or via `System.Diagnostics.DiagnosticSource` on older TFMs):

```csharp
using System.Diagnostics;
using System.Diagnostics.Metrics;

public sealed class MetricsComExecutor : IComExecutor, IDisposable
{
    public const string MeterName = "AdaskoTheBeAsT.Interop.COM";

    private readonly IComExecutor _inner;
    private readonly Meter _meter;
    private readonly Counter<long> _executions;
    private readonly Counter<long> _executionFailures;
    private readonly Histogram<double> _executionDurationMs;

    public MetricsComExecutor(IComExecutor inner, IMeterFactory meterFactory)
    {
        _inner = inner;
        _meter = meterFactory.Create(MeterName);
        _executions = _meter.CreateCounter<long>("com.execute.count");
        _executionFailures = _meter.CreateCounter<long>("com.execute.failures");
        _executionDurationMs = _meter.CreateHistogram<double>(
            "com.execute.duration",
            unit: "ms");
    }

    public Result Execute(string comAssemblyPath, string manifestPath, Action action)
    {
        var tag = new KeyValuePair<string, object?>("manifest", manifestPath);
        var sw = Stopwatch.StartNew();
        var result = _inner.Execute(comAssemblyPath, manifestPath, action);
        sw.Stop();

        _executions.Add(1, tag);
        _executionDurationMs.Record(sw.Elapsed.TotalMilliseconds, tag);
        if (!result.Success)
        {
            _executionFailures.Add(1, tag);
        }
        return result;
    }

    // Delegate the rest of the interface identically, instrumenting as needed.
    public Result Execute(ICollection<ComPathDescriptor> descriptors, Action action)
        => _inner.Execute(descriptors, action);

    public ComObjectCreationResult<T> Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory)
        where T : class
        => _inner.Create(comAssemblyPath, manifestPath, factory);

    public ComObjectCreationResult<T> Create<T>(ICollection<ComPathDescriptor> descriptors, Func<T> factory)
        where T : class
        => _inner.Create(descriptors, factory);

    public Result Free<T>(ComObjectHandle<T> handle)
        where T : class
        => _inner.Free(handle);

    public void Dispose() => _meter.Dispose();
}
```

Register (logging + metrics, chained):

```csharp
services.AddLogging();
services.AddMetrics();

services.AddSingleton<ComExecutor>();
services.AddSingleton<IComExecutor>(sp =>
{
    IComExecutor exec = sp.GetRequiredService<ComExecutor>();
    exec = new LoggingComExecutor(exec, sp.GetRequiredService<ILogger<LoggingComExecutor>>());
    exec = new MetricsComExecutor(exec, sp.GetRequiredService<IMeterFactory>());
    return exec;
});
```

Or with Scrutor:

```csharp
services.AddSingleton<IComExecutor, ComExecutor>();
services.Decorate<IComExecutor, MetricsComExecutor>();
services.Decorate<IComExecutor, LoggingComExecutor>(); // runs first (outermost)
```

Consume the metrics via your preferred exporter — OpenTelemetry, Application Insights, `dotnet-counters`, etc. — by subscribing to the meter name `"AdaskoTheBeAsT.Interop.COM"`.

## ⚙️ How It Works

1. **Activation Context Creation** - Creates Windows activation context from your manifest file
2. **Context Activation** - Activates the context to enable registration-free COM
3. **COM Execution** - Runs your code with the COM object in STA
4. **Message Pumping** - Processes Windows messages for COM callbacks
5. **Cleanup** - Automatically deactivates and releases contexts

## 📜 Creating COM Manifests

### Using ManifestMaker (Commercial Tool)

You can generate manifests for your COM DLLs using **[ManifestMaker](https://www.manifestmaker.com/)**, a commercial desktop tool:

1. **Download and install** ManifestMaker from [https://www.manifestmaker.com/](https://www.manifestmaker.com/)
2. **Load** your COM DLL file or enter COM details manually
3. **Configure** settings:
   - Assembly name and version
   - COM class details (CLSID, threading model, ProgID)
   - File dependencies
4. **Generate** and download your manifest file
5. **Place** the manifest file in the same directory as your COM DLL

**Benefits:**
- ✨ Automatically extracts COM metadata from DLL
- 🎯 Generates proper manifest structure
- 🔧 Supports type libraries and dependencies

**Note:** ManifestMaker is a paid service. Check their website for current pricing and licensing terms.

### Manual Manifest Creation

If you prefer to create manifests manually or need custom configuration:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity
    type="win32"
    name="NativeCOM"
    version="1.0.0.0"/>
  <file name="NativeCOM.dll">
    <comClass
      clsid="{YOUR-CLSID-HERE}"
      threadingModel="Apartment"
      progid="NativeCOM.StringConcatenator"/>
  </file>
</assembly>
```

### Finding Your COM CLSID

If you need to find the CLSID of your COM component:

**Option 1: Using OleView (Windows SDK)**
```bash
# Open OleView.exe from Windows SDK
# Navigate to "All Objects" and search for your component
```

**Option 2: Using Registry (if previously registered)**
```bash
# Search in Registry Editor (regedit)
HKEY_CLASSES_ROOT\CLSID
HKEY_CLASSES_ROOT\YourProgId
```

**Option 3: Using PowerShell**
```powershell
# Load the DLL and inspect COM registration
$assembly = [System.Reflection.Assembly]::LoadFile("C:\path\to\your.dll")
$types = $assembly.GetTypes() | Where-Object { $_.IsComVisible -eq $true }
$types | ForEach-Object { $_.GUID }
```

**Option 4: Using TlbImp/RegAsm**
```bash
# For .NET COM components
regasm YourAssembly.dll /regfile:output.reg
# Inspect output.reg for CLSID values

# For native COM
tlbimp NativeCOM.dll /out:Interop.dll /verbose
```

### Manifest Threading Models

Choose the appropriate threading model for your COM component:

| Threading Model | Description | Use When |
|----------------|-------------|----------|
| `Apartment` | Single-threaded apartment (STA) | UI components, most legacy COM |
| `Free` | Multi-threaded apartment (MTA) | Thread-safe, no UI interaction |
| `Both` | Supports both STA and MTA | Flexible components |
| `Neutral` | Thread-neutral | Advanced scenarios |

**Recommendation:** Use `Apartment` for most scenarios, especially if unsure.

### Manifest File Example with Type Library

For COM components with type libraries:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity
    type="win32"
    name="NativeCOM"
    version="1.0.0.0"/>
  <file name="NativeCOM.dll">
    <comClass
      clsid="{12345678-1234-1234-1234-123456789012}"
      threadingModel="Apartment"
      progid="NativeCOM.StringConcatenator.1"
      description="String Concatenator COM Component"/>
    <typelib
      tlbid="{87654321-4321-4321-4321-210987654321}"
      version="1.0"
      helpdir=""
      flags="HASDISKIMAGE"/>
  </file>
</assembly>
```

### Validating Your Manifest

After creating your manifest, validate it:

1. **Well-formed XML** - Ensure proper XML syntax
2. **Correct CLSID** - Match your COM component's CLSID
3. **File name** - Must match the actual DLL name
4. **Threading model** - Appropriate for your component
5. **Test** - Run with this library to verify it works

```csharp
// Quick validation test (using injected IComExecutor)
var result = executor.Execute(@"C:\path\to\COM.dll", @"C:\path\to\COM.manifest", () =>
{
    Console.WriteLine("Manifest loaded successfully!");
});

if (!result.Success)
{
    Console.WriteLine($"Manifest error: {result.Exception?.Message}");
}
```

## 📖 API Reference

### `IComExecutor` Interface (recommended)

The primary entry point for new code. Implemented by `ComExecutor`. Register `ComExecutor` as a singleton with your DI container and depend on the interface; every method below is testable via standard mocking tools.

```csharp
[SupportedOSPlatform("windows")]
public interface IComExecutor
{
    Result Execute(string comAssemblyPath, string manifestPath, Action action);
    Result Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action);

    ComObjectCreationResult<T> Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory)
        where T : class;
    ComObjectCreationResult<T> Create<T>(ICollection<ComPathDescriptor> comPathDescriptors, Func<T> factory)
        where T : class;

    Result Free<T>(ComObjectHandle<T> comObjectHandle) where T : class;
}
```

#### `Execute(string comAssemblyPath, string manifestPath, Action action)`

Executes an action with a single COM activation context.

**Parameters:**
- `comAssemblyPath` - Full path to the COM DLL
- `manifestPath` - Full path to the manifest file
- `action` - Code to execute in the activation context

**Returns:** `Result` object containing success status and any exception

#### `Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action)`

Executes an action with multiple COM activation contexts, activated in order and deactivated in reverse order.

#### `Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory)`

Creates a COM object inside a registration-free activation context and returns a `ComObjectHandle<T>`.
Release the handle later by calling `handle.Dispose()` (via `using`) or `executor.Free(handle)` on the same thread.

#### `Create<T>(ICollection<ComPathDescriptor> comPathDescriptors, Func<T> factory)`

Creates a COM object inside multiple registration-free activation contexts and returns a `ComObjectHandle<T>`.

#### `Free<T>(ComObjectHandle<T> comObjectHandle)`

Releases a handle created by one of the `Create` overloads. Idempotent; a second call on an already-released handle is a no-op and returns success.

### `ComExecutor` Class (recommended implementation)

Stateless default implementation of `IComExecutor`. Register as a singleton:

```csharp
services.AddSingleton<IComExecutor, ComExecutor>();
```

```csharp
[SupportedOSPlatform("windows")]
public sealed class ComExecutor : IComExecutor { /* delegates to the static Executor */ }
```

### `Executor` Static Class (legacy; kept for v2.x compatibility)

The original static entry point, kept so existing code compiles unchanged. **Prefer `IComExecutor` + `ComExecutor` for all new code.** `ComExecutor` itself is a thin forwarder to this class, so the two APIs have identical behaviour.

```csharp
[SupportedOSPlatform("windows")]
public static class Executor
{
    public static Result Execute(string comAssemblyPath, string manifestPath, Action action);
    public static Result Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action);
    public static ComObjectCreationResult<T> Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory)
        where T : class;
    public static ComObjectCreationResult<T> Create<T>(ICollection<ComPathDescriptor> comPathDescriptors, Func<T> factory)
        where T : class;
    public static Result Free<T>(ComObjectHandle<T> comObjectHandle) where T : class;
}
```

### Result Class

```csharp
public class Result
{
    public bool Success { get; set; }
    public Exception? Exception { get; set; }
}
```

### ComObjectCreationResult<T> Class

```csharp
public sealed class ComObjectCreationResult<T> : Result
    where T : class
{
    public ComObjectHandle<T>? Value { get; set; }
}
```

### ComObjectHandle<T> Class

Starting with v3.0.0 the handle implements `IDisposable`, so the simplest way to release it is a `using` block.
Calling `Dispose()` is equivalent to calling `executor.Free(handle)` and is idempotent.

```csharp
[SupportedOSPlatform("windows")]
public sealed class ComObjectHandle<T>
    : IDisposable
    where T : class
{
    public T? ComObject { get; internal set; }
    public bool IsReleased { get; }
    public void Dispose();
}
```

### ComPathDescriptor Class

```csharp
public sealed class ComPathDescriptor
{
    public ComPathDescriptor(string comAssemblyPath, string comManifestPath);
    public string ComAssemblyPath { get; }
    public string ComManifestPath { get; }
}
```

### ActCtxWrongSizeException Class

Thrown when the internal activation-context structure size does not match the current process architecture.

## 🎯 Best Practices

1. ✅ **Always check `Result.Success`** before assuming COM execution succeeded
2. ✅ **Use absolute paths** for COM DLL and manifest files
3. ✅ **Keep manifests alongside DLLs** for easy deployment
4. ✅ **Handle exceptions** - COM calls can fail in various ways
5. ✅ **Test on target platform** - COM behavior can vary by Windows version
6. ⚠️ **Avoid long-running operations** in the action delegate
7. ⚠️ **Be aware of STA threading model** requirements

## 📋 Requirements

- **OS**: Windows (uses Windows-specific APIs: kernel32.dll, user32.dll). The public surface is annotated with `[SupportedOSPlatform("windows")]` on .NET 8+ so cross-platform analyzers (CA1416) will flag misuse.
- **Framework**: .NET Standard 2.0, .NET 8.0 / 9.0 / 10.0, .NET Framework 4.6.2 / 4.7 / 4.7.1 / 4.7.2 / 4.8 / 4.8.1
- **Architecture**: x86, x64 (defined in your project configuration)

## 🩺 Troubleshooting

### Common Issues

**Issue**: `ActCtxWrongSizeException`  
**Solution**: This indicates a platform mismatch. Ensure your target platform (x86/x64) matches your COM DLL architecture.

**Issue**: `Win32Exception` when creating context  
**Solution**: Check that:
- COM DLL path is valid and file exists
- Manifest file path is valid and file exists
- Manifest XML is well-formed
- CLSID in manifest matches the COM component

**Issue**: COM object creation fails silently  
**Solution**: Verify the manifest `progid` and `clsid` match your COM component registration.

### Detecting leaked `ComObjectHandle<T>` in production

If a caller forgets to `Dispose()` / `Executor.Free` a `ComObjectHandle<T>`, its finalizer emits a warning-level event on the `System.Diagnostics.Tracing.EventSource` named **`AdaskoTheBeAsT.Interop.COM`** (Event ID `1`, `HandleLeaked`, payload: short type name). The signal is visible in Release builds and to out-of-process diagnostic tools.

Collect from the command line with `dotnet-trace`:

```bash
dotnet tool install --global dotnet-trace
dotnet-trace collect --providers AdaskoTheBeAsT.Interop.COM --process-id <pid>
```

Or subscribe in-process via `EventListener`:

```csharp
using System.Diagnostics.Tracing;

public sealed class LeakedHandleListener : EventListener
{
    private readonly ILogger<LeakedHandleListener> _logger;

    public LeakedHandleListener(ILogger<LeakedHandleListener> logger) => _logger = logger;

    protected override void OnEventSourceCreated(EventSource source)
    {
        if (source.Name == "AdaskoTheBeAsT.Interop.COM")
        {
            EnableEvents(source, EventLevel.Warning);
        }
    }

    protected override void OnEventWritten(EventWrittenEventArgs e)
    {
        if (e.EventId == 1 && e.Payload is { Count: > 0 })
        {
            _logger.LogWarning("Leaked ComObjectHandle<{Type}>", e.Payload[0]);
        }
    }
}
```

Any non-zero rate of `HandleLeaked` events is a bug in the consuming code — either add a `using` statement around the handle or call `Executor.Free` explicitly on the creating thread.

## 🧪 Building from Source

```bash
git clone https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM.git
cd AdaskoTheBeAsT.Interop.COM
dotnet restore
dotnet build
dotnet test
```

The project includes:
- C# library (`src/AdaskoTheBeAsT.Interop.COM`)
- Native COM example (`src/NativeCOM`)
- Unit tests (`test/unit/AdaskoTheBeAsT.Interop.COM.Test`)

## ✅ Code Quality

This project maintains high code quality standards with:
- 20+ static analyzers (StyleCop, Roslynator, SonarAnalyzer, etc.)
- Nullable reference types enabled
- Comprehensive documentation
- Unit test coverage
- Strict warning-as-error policy

## 🙋 Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Ensure all analyzers pass
4. Add unit tests for new functionality
5. Submit a pull request

## 📜 Changelog

Historical design decisions are captured as ADRs in [`docs/adr/`](docs/adr/README.md).

### v3.0.0

Breaking only in the sense that new analyzer warnings may surface in consumer code; the runtime API is source-compatible with v2.1.0.

- **Added** — `ComObjectHandle<T>` implements `IDisposable`. Prefer `using var handle = creation.Value!;` over the explicit `Executor.Free(handle)` call. A diagnostic-only finalizer emits a warning-level `HandleLeaked` event on the `AdaskoTheBeAsT.Interop.COM` `EventSource` in all build flavours if a handle was never released, and additionally raises `Debug.Fail` when a debugger is attached. (ADR-0012, ADR-0018)
- **Added** — `ComInteropEventSource` (`EventSource` name: `AdaskoTheBeAsT.Interop.COM`). Subscribe with `dotnet-trace collect --providers AdaskoTheBeAsT.Interop.COM` or an in-process `EventListener` to detect leaked `ComObjectHandle<T>` in production. (ADR-0018)
- **Added** — `IComExecutor` interface and `ComExecutor` class, now the **recommended** API for all new code. Register as `services.AddSingleton<IComExecutor, ComExecutor>();` and depend on `IComExecutor`. The static `Executor` class is retained for source-level compatibility with v2.x but is demoted to "legacy" in the documentation. (ADR-0013)
- **Added** — `[SupportedOSPlatform("windows")]` on `Executor`, `ComObjectHandle<T>`, `IComExecutor`, and `ComExecutor` for TFMs that support it. Removed the in-method `#pragma warning disable CA1416`. (ADR-0014)
- **Added** — Target frameworks: `.NET 10.0`, `.NET Framework 4.6.2`, `4.7`, `4.7.1`, `4.7.2`, `4.8`, `4.8.1`. The full matrix is now `netstandard2.0`, `net462`, `net47`, `net471`, `net472`, `net48`, `net481`, `net8.0`, `net9.0`, `net10.0`. (ADR-0007)
- **Added** — `ComObjectHandle<T>.IsReleased` is now `public` (was `internal`).
- **Added** — GitHub Actions CI workflow using the shared reusable pipeline. (ADR-0015)
- **Added** — Expanded test suite from 7 to 63 tests across 9 test classes (`ComPathDescriptorTest`, `ActCtxWrongSizeExceptionTest`, `ResultTest`, `ExecuteValidationTest`, `ComObjectHandleDisposeTest`, `ComExecutorTest`, `ComInteropEventSourceTest`, `MessagePumpTest`). Local line coverage of `AdaskoTheBeAsT.Interop.COM.dll` is now **92.8 %** (the remaining ~7 % are native-failure `catch (Exception ex)` branches reachable only via fault-injection).
- **Changed** — The managed library now builds as `AnyCPU` (previously `x86`). 64-bit host processes can finally load the package. The existing `ActCtxWrongSizeException` guard continues to verify the correct `ACTCTX` layout at runtime for the current bitness. (ADR-0016)
- **Fixed** — `NativeMethods` and `ActCtxWrongSizeException` now compile cleanly on every non-NET8+ target (previously a recent TFM addition broke the .NET Framework builds).
- **Docs** — Added `docs/adr/` with architecture decision records for every significant choice since v1.0.0.

### v2.1.0 (2026-04-07)

- **Added** — `Executor.Create<T>` / `Executor.Free<T>` lifetime API, `ComObjectHandle<T>`, `ComObjectCreationResult<T>`. (ADR-0010)
- **Added** — STA message pumping via `NativeMethods.PumpPendingMessages` around every public API path. (ADR-0011)
- **Docs** — Expanded README with scheduler/Threading samples.

### v2.0.0 (2025-08-10 / 2025-11-22)

- **Added** — `ComPathDescriptor` plus collection overloads of `Execute`. (ADR-0009)
- **Added** — `.NET 8.0` / `.NET 9.0` target frameworks with `LibraryImport` source-generated P/Invoke. (ADR-0008)

### v1.0.0 (2023-12-09)

- Initial public release: static `Executor`, `Result` pattern, `netstandard2.0` target, internal `NativeMethods` wrapping `kernel32`/`user32`. (ADR-0002, ADR-0003, ADR-0004, ADR-0005)

## 🚚 Migration Guide: v2.x → v3.0

The public API is source-compatible. No symbol has been removed or renamed. The following notes help you adopt the v3.0 idioms and explain the observable changes.

### 1. Prefer `IComExecutor` over the static `Executor`

`IComExecutor` (implemented by `ComExecutor`) is now the primary API. It is functionally identical to the static `Executor` but can be injected and mocked. Migrate at your own pace — the static API stays available indefinitely — but every new class you write should depend on `IComExecutor`.

Old:

```csharp
var result = Executor.Execute(dll, manifest, Work);
```

New:

```csharp
public sealed class MyService
{
    private readonly IComExecutor _executor;
    public MyService(IComExecutor executor) => _executor = executor;

    public void Run() => _executor.Execute(dll, manifest, Work);
}

// Composition root:
services.AddSingleton<IComExecutor, ComExecutor>();
```

### 2. Prefer `using` over `executor.Free`

Old:

```csharp
var creation = _executor.Create(dll, manifest, () => new Foo());
var handle = creation.Value!;
try
{
    Use(handle.ComObject!);
}
finally
{
    _executor.Free(handle);
}
```

New (recommended):

```csharp
var creation = _executor.Create(dll, manifest, () => new Foo());
using var handle = creation.Value!;
Use(handle.ComObject!);
```

Both forms are valid and equivalent. `Dispose` simply calls `Executor.Free` under the hood.

### 3. Handle the new `CA1416` warning on non-Windows projects

Because `Executor`, `IComExecutor`, `ComExecutor`, and `ComObjectHandle<T>` are now annotated with `[SupportedOSPlatform("windows")]`, consumers that target frameworks like `net8.0` (without `TargetPlatformIdentifier=Windows`) will start to see `CA1416` warnings at the call-sites. Fix options, in order of preference:

1. Add `<TargetFramework>net8.0-windows</TargetFramework>` (or the equivalent platform-qualified TFM) to the calling project. This is the correct, honest expression of the dependency.
2. Guard calls with `if (OperatingSystem.IsWindows()) { ... }`.
3. Annotate the calling member with `[SupportedOSPlatform("windows")]`.

### 4. Unit-test code that depends on `IComExecutor`

Because `IComExecutor` is an interface, `Moq` / `NSubstitute` / your favourite substitution library can replace it directly:

```csharp
var mock = new Mock<IComExecutor>();
mock.Setup(e => e.Execute(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<Action>()))
    .Callback<string, string, Action>((_, _, a) => a())
    .Returns(new Result { Success = true });

var sut = new MyService(mock.Object);
// assert...
```

There is no pressure to migrate existing call-sites that use the static `Executor`; both styles coexist indefinitely.

### 5. Watch for `Debug.Fail` messages on leaked handles

If you forget to call `Dispose`/`Executor.Free`, the finalizer will emit `Debug.Fail("ComObjectHandle<...> was not disposed...")` in Debug builds. Treat these as bug reports, not as runtime errors — Release builds simply ignore the message.

### 6. Target framework changes

- If you were installing the `netstandard2.0` DLL into a .NET Framework 4.6.2-4.8.1 app, v3.0 now ships dedicated DLLs for each of those TFMs. NuGet will pick the best match automatically; no action required.
- `.NET 10.0` is available as a first-class target.

## 🏗️ Building native COM payloads for 32-bit and 64-bit hosts

Registration-free COM requires three artefacts that match the bitness of the host process: the **native COM DLL** (produced by your C++ project), the **side-by-side manifest** (`processorArchitecture="x86"` vs `"amd64"`), and the **managed interop assembly** (`Interop.<Lib>.dll`) produced by `TlbImp.exe`. The .NET library in this repository is `AnyCPU` — it will load in either bitness — but the native DLL is arch-specific. This section documents two patterns you can use when you need to ship or test your COM surface on both architectures.

The sample in `src/NativeCOM` (built through `NativeCOM.slnx`) and the test project (`test/unit/AdaskoTheBeAsT.Interop.COM.Test`) follow **Option A** out of the box.

### Option A - one agnostic interop assembly

`Interop.<Lib>.dll` is metadata-only. It contains COM interop attributes, IIDs, CLSIDs, and type declarations, but zero executable code. If you pass `/machine:Agnostic` to `TlbImp`, you get a pure MSIL assembly that loads from both 32-bit and 64-bit processes. Only the native DLL and the manifest differ between arches.

Generate the agnostic interop assembly once (the existing `generatelib.bat` in this repo does exactly this):

```bat
"C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\TlbImp.exe" ^
    .\x86\Debug\NativeCOM.dll ^
    /out:.\x86\Debug\Interop.NativeCOM.dll ^
    /namespace:NativeCOM ^
    /machine:Agnostic
```

Then route the arch-specific pieces with TFM conditions in your csproj:

```xml
<PropertyGroup Condition="$(TargetFramework.StartsWith('net4'))">
  <PlatformTarget>x86</PlatformTarget>
</PropertyGroup>

<!-- One MSIL-neutral interop assembly works for both 32-bit and 64-bit test hosts -->
<ItemGroup>
  <Reference Include="Interop.NativeCOM">
    <HintPath>..\..\..\x86\Debug\Interop.NativeCOM.dll</HintPath>
  </Reference>
</ItemGroup>

<!-- 32-bit native payload + x86 manifest for .NET Framework TFMs -->
<ItemGroup Condition="$(TargetFramework.StartsWith('net4'))">
  <None Include="..\..\..\x86\Debug\NativeCOM.dll">
    <Link>NativeCOM.dll</Link>
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </None>
  <None Include="..\..\..\manifest\NativeCOM\NativeCOM.manifest">
    <Link>NativeCOM.manifest</Link>
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </None>
</ItemGroup>

<!-- 64-bit native payload + amd64 manifest for .NET Core TFMs -->
<ItemGroup Condition="!$(TargetFramework.StartsWith('net4'))">
  <None Include="..\..\..\x64\Debug\NativeCOM.dll">
    <Link>NativeCOM.dll</Link>
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </None>
  <None Include="..\..\..\manifest\NativeCOM\NativeCOM.x64.manifest">
    <Link>NativeCOM.manifest</Link>
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </None>
</ItemGroup>
```

The two manifests differ only in `processorArchitecture`:

```xml
<!-- manifest\NativeCOM\NativeCOM.manifest (32-bit) -->
<assemblyIdentity name="NativeCOM" version="1.0.0.0" type="win32" processorArchitecture="x86"/>

<!-- manifest\NativeCOM\NativeCOM.x64.manifest (64-bit) -->
<assemblyIdentity name="NativeCOM" version="1.0.0.0" type="win32" processorArchitecture="amd64"/>
```

Verify the interop assembly is truly architecture-neutral:

```powershell
corflags .\x86\Debug\Interop.NativeCOM.dll
# Expected:
#   ILONLY    : 1
#   32BITREQ  : 0
#   32BITPREF : 0
```

**Pros** - one interop assembly to ship and track in source control. No arch-matching logic in the consuming csproj beyond the native DLL and manifest pair.
**Cons** - none of any consequence for metadata-only interop assemblies. If your `TlbImp`-generated assembly ever had to do platform-specific marshalling (extremely rare and usually a sign your IDL uses non-portable types like raw `LONG_PTR` in public signatures), the agnostic flag would hide the mismatch.

### Option B - two arch-specific interop assemblies

If you prefer explicit per-arch artefacts (sometimes required by compliance tooling or when your IDL genuinely has architecture-dependent types), build two separate `Interop.<Lib>.dll` copies - one per bitness. This requires the COM DLL's embedded type library to be emitted with the correct `SYSKIND`, otherwise `TlbImp /machine:X64` will fail with `TI2010: A single valid machine type compatible with the input type library must be specified.`

**Step 1 - make MIDL emit a 64-bit TLB for x64 configurations.** Open `NativeCOM.vcxproj` (or your own C++ project), locate the `<ItemDefinitionGroup>` entries for the x64 configurations, and add `<TargetEnvironment>X64</TargetEnvironment>` inside the `<Midl>` block:

```xml
<ItemDefinitionGroup Condition="'$(Configuration)|$(Platform)'=='Debug|x64'">
  <Midl>
    <TargetEnvironment>X64</TargetEnvironment>
    <PreprocessorDefinitions>_DEBUG;%(PreprocessorDefinitions)</PreprocessorDefinitions>
    <MkTypLibCompatible>false</MkTypLibCompatible>
    <!-- existing <HeaderFileName>, <InterfaceIdentifierFileName>, etc. -->
  </Midl>
  <!-- ... -->
</ItemDefinitionGroup>

<ItemDefinitionGroup Condition="'$(Configuration)|$(Platform)'=='Release|x64'">
  <Midl>
    <TargetEnvironment>X64</TargetEnvironment>
    <!-- ... -->
  </Midl>
</ItemDefinitionGroup>
```

For the `Win32` configurations either leave `<TargetEnvironment>` unset (the default produces a 32-bit TLB) or set it explicitly to `Win32`.

**Step 2 - rebuild the native DLL for both platforms.** `NativeCOM.slnx` in this repository carries both `x64` and `Win32` configurations; run the build in Visual Studio (or a Developer Command Prompt) for each arch so you end up with `x86\Debug\NativeCOM.dll` *and* `x64\Debug\NativeCOM.dll` whose embedded TLBs have the matching `SYSKIND`.

**Step 3 - run `TlbImp` twice, once per arch.** Point it at each arch-specific DLL and emit two arch-stamped interop assemblies:

```bat
"C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\TlbImp.exe" ^
    .\x86\Debug\NativeCOM.dll ^
    /out:.\x86\Debug\Interop.NativeCOM.dll ^
    /namespace:NativeCOM ^
    /machine:X86

"C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\TlbImp.exe" ^
    .\x64\Debug\NativeCOM.dll ^
    /out:.\x64\Debug\Interop.NativeCOM.dll ^
    /namespace:NativeCOM ^
    /machine:X64
```

Verify the two assemblies have the `32BITREQ` flag set correctly:

```powershell
corflags .\x86\Debug\Interop.NativeCOM.dll   # 32BITREQ : 1
corflags .\x64\Debug\Interop.NativeCOM.dll   # 32BITREQ : 0 (x64 only)
```

**Step 4 - TFM-condition the `Reference` as well.** Unlike Option A, the interop assembly now differs per arch, so extend the TFM conditions to both the native payload, manifest, **and** the managed reference:

```xml
<ItemGroup Condition="$(TargetFramework.StartsWith('net4'))">
  <Reference Include="Interop.NativeCOM">
    <HintPath>..\..\..\x86\Debug\Interop.NativeCOM.dll</HintPath>
  </Reference>
  <None Include="..\..\..\x86\Debug\NativeCOM.dll">
    <Link>NativeCOM.dll</Link>
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </None>
  <None Include="..\..\..\manifest\NativeCOM\NativeCOM.manifest">
    <Link>NativeCOM.manifest</Link>
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </None>
</ItemGroup>

<ItemGroup Condition="!$(TargetFramework.StartsWith('net4'))">
  <Reference Include="Interop.NativeCOM">
    <HintPath>..\..\..\x64\Debug\Interop.NativeCOM.dll</HintPath>
  </Reference>
  <None Include="..\..\..\x64\Debug\NativeCOM.dll">
    <Link>NativeCOM.dll</Link>
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </None>
  <None Include="..\..\..\manifest\NativeCOM\NativeCOM.x64.manifest">
    <Link>NativeCOM.manifest</Link>
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </None>
</ItemGroup>
```

**Pros** - explicit, auditable per-arch artefacts; each interop assembly is stamped with the expected `/machine:` value; catches TLB/SYSKIND drift early (if the C++ build is misconfigured, the `TI2010` error fires at `TlbImp` time rather than at runtime).
**Cons** - two interop assemblies to version-control and keep in sync; a vcxproj configuration mistake (forgetting `<TargetEnvironment>X64</TargetEnvironment>` on one configuration) surfaces as the `TI2010` error.

### Which option to pick

For most registration-free COM scenarios **Option A is the simpler choice** and is what this repository uses for its sample test harness. Pick Option B when:

- Your IDL contains architecture-dependent signatures (raw `LONG_PTR`, `SIZE_T`, inline pointer arithmetic exposed across COM boundaries).
- A downstream policy requires per-architecture managed artefacts (for example, a signed-assembly policy that stamps each binary with an explicit machine type).
- You want build-time detection of TLB/SYSKIND drift in your native project.

In both options the managed `AdaskoTheBeAsT.Interop.COM` library itself stays `AnyCPU` - only the native payloads, manifests, and interop assemblies (when using Option B) are split by arch.

## 📄 License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

## 👤 Credits

Created and maintained by [Adam Pluciński](https://github.com/AdaskoTheBeAsT).

## 🔗 Related Resources

- 📘 [Registration-Free COM Interop](https://learn.microsoft.com/en-us/dotnet/framework/interop/registration-free-com-interop)
- 📗 [Activation Contexts](https://learn.microsoft.com/en-us/windows/win32/sbscs/activation-contexts)
- 📙 [Side-by-Side Assemblies](https://learn.microsoft.com/en-us/windows/win32/sbscs/about-side-by-side-assemblies-)
- 📕 [`TlbImp.exe` reference](https://learn.microsoft.com/en-us/dotnet/framework/tools/tlbimp-exe-type-library-importer)
- 🧭 [This library's ADR index](docs/adr/README.md)

---

**Questions or issues?** 🐛 Open an issue on [GitHub](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/issues) or start a discussion — PRs welcome too. 💬

---

<p align="center">
  Built for the kind of code that calls into a <strong>20-year-old ActiveX control</strong>, <strong>forgets to register anything</strong>, and <strong>still ships on Monday</strong>. ✨<br/>
  Made with ❤️ (and a lot of coffee ☕) by <a href="https://github.com/AdaskoTheBeAsT">AdaskoTheBeAsT</a>.
</p>