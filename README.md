# AdaskoTheBeAsT.Interop.COM

> Registration-free COM interop for .NET. Skip `regsvr32`, skip the registry, skip the drama.

[![NuGet](https://img.shields.io/nuget/v/AdaskoTheBeAsT.Interop.COM.svg?logo=nuget)](https://www.nuget.org/packages/AdaskoTheBeAsT.Interop.COM/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/AdaskoTheBeAsT.Interop.COM.svg?logo=nuget)](https://www.nuget.org/packages/AdaskoTheBeAsT.Interop.COM/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/LICENSE)
![Windows](https://img.shields.io/badge/platform-Windows-0078D6)
![AnyCPU](https://img.shields.io/badge/managed%20library-AnyCPU-512BD4?logo=dotnet)
[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop.COM&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop.COM)
[![Coverage](https://sonarcloud.io/api/project_badges/measure?project=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop.COM&metric=coverage)](https://sonarcloud.io/component_measures?id=AdaskoTheBeAsT_AdaskoTheBeAsT.Interop.COM&metric=coverage)

## 👋 Hello, COM-wrangler

Maybe it's a vendor SDK. Maybe it's an in-house ATL library. Maybe it's that ActiveX component from 2005 that has now outlived three rewrites of the application around it.

You just want to call a method from .NET, not turn deployment into a registry troubleshooting session.

**Bring your DLL and a matching manifest.** `AdaskoTheBeAsT.Interop.COM` handles Windows activation contexts so you can use compatible COM components without registering them globally. Your application gets an explicit lifetime API; your deployment instructions lose a `regsvr32` step.

> 🎉 **Version 4.0.0:** host-owned message pumping, stricter lifetime checks, and a focused framework lineup. Upgrading? Check the [changelog](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/CHANGELOG.md) and [migration guide](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/MIGRATION.md). Still on 3.0? The [v3.0.0 README](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/v3.0.0/README.md) has you covered.

## ✨ Why you'll love this

- 📦 **Deployment without the registration ceremony.** Activate one or more manifests without registering the COM classes globally.
- 🧹 **A familiar `using` block.** `ComObjectHandle<T>` gives an exclusively owned COM object an explicit lifetime. Prefer `Free` when you need to inspect the cleanup result.
- 🧩 **DI-friendly, container-optional.** Inject `IComExecutor`, mock it in tests, or use the existing static `Executor`. No separate DI package to hunt down.
- 🧵 **Your thread stays your thread.** Work runs synchronously where you call it. No surprise background thread or apartment switch.
- 🪟 **Your message loop gets a say.** Use the bounded built-in pump, or disable it in 4.0 when your UI framework or STA scheduler owns message processing.
- 🔎 **Leaks don't have to be silent.** Event-source diagnostics expose leaked handles and failed disposal in Debug and Release builds.
- 🪶 **Small package, broad reach.** One AnyCPU managed assembly, modern .NET and .NET Framework targets, and no additional runtime package dependencies.

> **Small library, clear boundaries.** This package does not generate interop assemblies, register COM servers, initialize COM, create an STA thread, or replace your application's message loop. You supply the native DLL, its dependencies, a matching manifest, and the apartment your component requires.

## 🗺️ Find your way

**Here to make your first COM call?** Head to [Install](#install), then [Quick start](#quick-start). Upgrading an existing app? Start with the [migration guide](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/MIGRATION.md).

- 🧰 [Requirements](#requirements)
- 📦 [Install](#install)
- 🚀 [Quick start](#quick-start)
- 🧩 [Choose an API](#choose-an-api)
- 🧵 [Lifetime and threading rules](#lifetime-and-threading-rules)
- 🪟 [Message processing](#message-processing)
- 📜 [Manifests and deployment](#manifests-and-deployment)
- 🩺 [Diagnostics and troubleshooting](#diagnostics-and-troubleshooting)
- 🧪 [Build and test](#build-and-test)
- 📚 [Further documentation](#further-documentation)

<a id="requirements"></a>

## 🧰 Requirements

The short version: Windows, a compatible .NET target, and native bits that match your process.

| Requirement | Version 4.0.0 |
| --- | --- |
| Operating system | Windows only, including when called from a platform-neutral .NET target |
| Modern .NET targets | `net8.0`, `net9.0`, `net10.0` |
| .NET Framework targets | `net472`, `net48`, `net481` |
| Process architecture | x86 or x64, matching the native COM DLL and its dependencies |
| COM metadata | A managed interop assembly or equivalent COM interface declarations |
| Apartment | Supplied by the caller; STA when required by the COM component |

**4.0 removes `netstandard2.0`, `net462`, `net47`, and `net471` package assets.** Consumers on those targets must retarget or remain on a compatible earlier package version. The managed library being AnyCPU does not make a native COM DLL architecture-neutral.

<a id="install"></a>

## 📦 Install

One package to add:

```powershell
dotnet add package AdaskoTheBeAsT.Interop.COM --version 4.0.0
```

For a Windows-only modern .NET application, use a Windows-qualified target such as `net10.0-windows`. This also communicates the platform requirement to the compiler's platform compatibility analyzer.

<a id="quick-start"></a>

## 🚀 Quick start

Let's make COM say hello. This example uses the repository's `NativeCOM.StringConcatenatorClass`. For your component, substitute its interop type, method, DLL, and manifest.

Three things to line up before hitting Run:

1. Reference the managed `Interop.NativeCOM.dll` assembly.
2. Copy `NativeCOM.dll` and the matching x86 or x64 manifest to the application output directory. Name the copied manifest `NativeCOM.manifest`.
3. Set the application's `PlatformTarget` to match the native DLL, for example `x64`.

Then create it, call it, and let `using` attempt the cleanup. The synchronous STA entry point keeps the handle on one thread:

```csharp
using System;
using System.IO;
using AdaskoTheBeAsT.Interop.COM;

internal static class Program
{
    [STAThread]
    private static int Main()
    {
        IComExecutor executor = new ComExecutor();
        string dll = Path.Combine(AppContext.BaseDirectory, "NativeCOM.dll");
        string manifest = Path.Combine(AppContext.BaseDirectory, "NativeCOM.manifest");

        var creation = executor.Create(
            dll,
            manifest,
            () => new NativeCOM.StringConcatenatorClass());

        if (!creation.Success)
        {
            Console.Error.WriteLine(creation.Exception);
            return 1;
        }

        using var handle = creation.Value
            ?? throw new InvalidOperationException("No COM handle was returned.");

        var concatenator = handle.ComObject
            ?? throw new InvalidOperationException("No COM object was returned.");

        Console.WriteLine(concatenator.ConcatStrings("Hello, ", "COM!"));
        return 0;
    }
}
```

The payoff:

```text
Hello, COM!
```

No `regsvr32` invocation required. Just a COM call, with its activation context and ownership made explicit.

`Create` captures activation and factory failures in its result. Calls made on `handle.ComObject` afterward are ordinary COM calls and can throw. The `using` statement attempts cleanup in either case.

`Dispose` reports release failures through diagnostics rather than throwing. Use `Free` and inspect its result when you need to handle cleanup failures directly.

<a id="choose-an-api"></a>

## 🧩 Choose an API

One callback or a longer-lived object? That's the main choice:

| API | Use it when | Lifetime responsibility |
| --- | --- | --- |
| `Execute(dll, manifest, action)` | Work fits inside one synchronous callback | The library manages activation contexts only; you manage objects created inside the callback |
| `Create(dll, manifest, factory)` | You need an explicitly owned object and its activation contexts | Release the returned handle with `using` or `Free` |
| `Free(handle)` | You need the release result or the handle outlives a local scope | Check `Result.Success`; retain an incompletely released handle for a correct retry |
| Collection overloads | Several manifests must be active together | Contexts activate in collection order and deactivate in reverse order |

Prefer `IComExecutor` in application services so it can be substituted in unit tests. No DI container is required, and the existing static `Executor` signatures remain available.

Already using `Microsoft.Extensions.DependencyInjection`? The registration is a one-liner:

```csharp
using AdaskoTheBeAsT.Interop.COM;
using Microsoft.Extensions.DependencyInjection;

services.AddSingleton<IComExecutor, ComExecutor>();
```

`ComExecutor` stores only immutable configuration and can be a singleton. That does not make the COM objects or handles it creates safe to share between threads.

### Results and exceptions

COM can fail. The important part is knowing where to look:

- Check `Result.Success` after `Execute`, `Create`, and `Free`.
- Operational failures, including manifest activation and callback exceptions, appear in `Result.Exception`.
- If work and cleanup both fail, `Result.Exception` can be an `AggregateException`.
- Invalid arguments, such as null callbacks, blank paths, or empty descriptor collections, throw before native allocation.
- A successful empty `Execute` callback proves context activation, not that a particular COM class can be instantiated. Test the actual class too.

<a id="lifetime-and-threading-rules"></a>

## 🧵 Lifetime and threading rules

COM still has opinions about threads. This library makes those rules explicit rather than pretending they went away. The same rules apply to the instance and static APIs:

1. **Create, use, and release a handle on the same thread.** Do not let `await` or `Task.Run` move its use or disposal to another thread. For STA components, use an STA entry point or a scheduler that guarantees the same owning thread.
2. **Release handles in reverse creation order.** Nested `using` scopes enforce this naturally. Finish nested `Execute` calls before releasing an outer handle.
3. **Release handles created inside a callback or factory before that delegate returns.** Do not leave nested activation contexts active across the delegate boundary.
4. **Transfer exclusive ownership of the returned runtime callable wrapper (RCW).** Do not return borrowed, cached, or shared COM objects from a factory. `Free` / `Dispose` uses `Marshal.FinalReleaseComObject`, invalidating every managed alias to that RCW.
5. **Manage child RCWs separately.** Objects returned by COM properties or methods are not recursively released by the parent handle. Do not release the handle's RCW independently.
6. **Treat failed cleanup as unfinished work.** A rejected wrong-thread, reentrant, or out-of-order release leaves the object unchanged. A native cleanup failure may leave `ComObject` null while `IsReleased` remains false. Keep the handle for retry, but do not use an already released object.

Repeated release of an already released handle succeeds without doing additional work. Its finalizer only emits diagnostics; it cannot safely clean up thread-affine native state.

See the [migration guide](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/MIGRATION.md) for release-order and retry examples.

<a id="message-processing"></a>

## 🪟 Message processing

Already have a message loop? No need to give it an unexpected roommate.

Automatic pumping is enabled by default. After a successful callback or factory, and during handle release, the library processes a bounded batch of pending Windows messages:

- At most 256 messages per pump.
- Unicode retrieval and dispatch on all supported targets.
- `WM_QUIT` and its exit code are preserved for the host.
- Remaining messages are left for a later pump or the host loop.

The limit bounds message count, not execution time. A message handler or COM call can still block.

**New in 4.0:** disable automatic pumping when a UI framework or STA scheduler owns message processing:

```csharp
IComExecutor executor = new ComExecutor(pumpPendingMessages: false);
```

For DI, use an explicit factory:

```csharp
services.AddSingleton<IComExecutor>(
    _ => new ComExecutor(pumpPendingMessages: false));
```

Static callers can pass `pumpPendingMessages: false` to the new `Execute` and `Create` overloads. Each handle retains its creation policy, even when freed through another executor or the static API.

The built-in pump does not run host-specific accelerators, dialog preprocessing, or framework message filters. Custom `PostThreadMessage` messages have no window procedure and can be consumed without delivery. Disabling the pump leaves message processing to the host; it does not prevent COM itself from pumping during synchronous calls.

<a id="manifests-and-deployment"></a>

## 📜 Manifests and deployment

Think of the manifest as the introduction: "Windows, here's the COM class, and here's the DLL that provides it."

Use the component vendor's metadata or type library to create a manifest. The CLSID, file name, threading model, and architecture must describe the actual component. **Do not guess the threading model.** Setting `Apartment` in XML does not turn an arbitrary thread into an STA.

This minimal x64 example uses the repository sample's CLSID:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity
    name="NativeCOM"
    version="1.0.0.0"
    type="win32"
    processorArchitecture="amd64"/>
  <file name="NativeCOM.dll">
    <comClass
      clsid="{03d0aedd-2013-418a-8b92-e57c30f2ab26}"
      threadingModel="Apartment"/>
  </file>
</assembly>
```

Use `x86` for a 32-bit DLL and process. Components using type libraries or proxy/stub marshaling may need additional entries. The repository includes complete [x86](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/manifest/NativeCOM/NativeCOM.manifest) and [x64](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/manifest/NativeCOM/NativeCOM.x64.manifest) sample manifests.

Use absolute paths in application code, usually based on `AppContext.BaseDirectory`. The COM DLL's directory is used as the activation context's assembly directory. Deploy the native DLL, manifest, interop assembly, and required native dependencies together as appropriate for your component. Registration-free activation does not remove native runtime dependencies or guarantee that every vendor component supports this deployment model.

For inspecting native COM metadata, use the vendor's IDL/type library or Windows SDK tools such as OleView. `Assembly.LoadFile` reads managed assemblies, not arbitrary native COM DLLs.

See [native builds and architecture](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/docs/native-builds.md) for interop generation and x86/x64 deployment.

<a id="diagnostics-and-troubleshooting"></a>

## 🩺 Diagnostics and troubleshooting

When COM says no, start with `Result.Exception`. When cleanup goes missing, listen for the diagnostics.

The internal event source is observable by provider name **`AdaskoTheBeAsT.Interop.COM`**:

| Event ID | Name | Level | Meaning |
| --- | --- | --- | --- |
| 1 | `HandleLeaked` | Warning | A handle reached finalization without successful release |
| 2 | `HandleReleaseFailed` | Error | `Dispose` attempted release and received a failed result |

These events are available in Debug and Release builds. The provider is not a public API type; subscribe by name with an `EventListener`, ETW tooling, or, for modern .NET, `dotnet-trace`:

```powershell
dotnet-trace collect --providers AdaskoTheBeAsT.Interop.COM --process-id <pid>
```

Install the `dotnet-trace` tool separately if needed. Direct `Free` callers should inspect the returned result; the disposal-failure event is emitted by `Dispose`.

| Symptom | What to check |
| --- | --- |
| Failed result with `Win32Exception` | Inspect `NativeErrorCode`, manifest XML, paths, and native dependencies |
| Class not registered (`0x80040154`) | Verify the CLSID, active manifest, DLL architecture, and that construction occurs inside the active context |
| `BadImageFormatException` or native load failure | Match the process architecture to the native DLL and all dependencies |
| `ActCtxWrongSizeException` | This reports an internal `ACTCTX` structure-layout mismatch, not a direct check of the DLL's architecture; capture runtime details and report it |
| Failed release or `IsReleased == false` | Retry on the owner thread after newer handles and nested callbacks finish; inspect the failure before retrying |
| Missing UI/thread messages or unexpected reentrancy | Let the host own message processing by disabling the built-in pump |
| `CA1416` warning | Use a Windows-qualified TFM, guard the call with `OperatingSystem.IsWindows()`, or annotate the Windows-only caller |

<a id="build-and-test"></a>

## 🧪 Build and test

Want to peek under the hood or send a fix? From the repository root, on Windows:

```powershell
dotnet restore AdaskoTheBeAsT.Interop.COM.slnx
dotnet build AdaskoTheBeAsT.Interop.COM.slnx
dotnet test --solution AdaskoTheBeAsT.Interop.COM.slnx
```

Requirements:

- The .NET SDK selected by `global.json`.
- Matching .NET runtimes and .NET Framework targeting packs for the targets being built/tested.
- Visual Studio or Build Tools with **Desktop development with C++**, ATL support for rebuilding the sample, and a Windows SDK.

`global.json` selects Microsoft.Testing.Platform. Use `--solution` or `--project` rather than the older positional `dotnet test` syntax.

For a focused architecture check:

```powershell
dotnet test --project test/unit/AdaskoTheBeAsT.Interop.COM.Test/AdaskoTheBeAsT.Interop.COM.Test.csproj -f net10.0 --arch x64
dotnet test --project test/unit/AdaskoTheBeAsT.Interop.COM.Test/AdaskoTheBeAsT.Interop.COM.Test.csproj -f net10.0 --arch x86
dotnet test --project test/unit/AdaskoTheBeAsT.Interop.COM.Test/AdaskoTheBeAsT.Interop.COM.Test.csproj -f net472 -p:PlatformTarget=x64
```

Install the corresponding x86/x64 runtime before testing that architecture. Defaults are x64 for modern .NET and x86 for .NET Framework. The native lifetime fixture builds automatically into isolated intermediate directories; the existing `NativeCOM` DLLs and manifests supply the sample activation tests.

<a id="further-documentation"></a>

## 📚 Further documentation

The quick path stays here. The deeper dives have their own homes:

- 📝 [Changelog](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/CHANGELOG.md): what's new in 4.0.0 and what changed in earlier releases.
- 🚚 [Migration guide](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/MIGRATION.md): upgrade from 3.x to 4.0 or 2.x to 3.0 without guessing.
- 🧩 [Advanced usage](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/docs/advanced-usage.md): multiple manifests, DI decorators, testing, and STA scheduling.
- 🏗️ [Native builds and architecture](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/docs/native-builds.md): make the DLL, manifest, and interop assembly agree on x86/x64.
- 🧠 [Architecture decisions](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/docs/adr/README.md): why the library works this way. Historical context, not a current support matrix.
- 📖 [Microsoft: registration-free COM interop](https://learn.microsoft.com/en-us/dotnet/framework/interop/registration-free-com-interop).
- 📖 [Microsoft: activation contexts](https://learn.microsoft.com/en-us/windows/win32/sbscs/activation-contexts).

## 🤝 Contributing

Found a bug? Have a COM edge case worth a regression test? [Open an issue](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/issues) or send a pull request. Reproduction steps, target framework, process architecture, and a minimal manifest make a great starting point.

Please add tests for behavior changes and run the relevant builds and tests with analyzers enabled before submitting a pull request.

## 📄 License and credits

Created and maintained by [Adam Pluciński](https://github.com/AdaskoTheBeAsT). Licensed under the [MIT License](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/blob/main/LICENSE).

---

Built for the COM component that survived three rewrites and still has a job to do. ☕
