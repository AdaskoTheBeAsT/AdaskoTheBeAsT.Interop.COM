# Advanced usage

These examples supplement the [README](../README.md) and target version **4.0.0**. See the [migration guide](../MIGRATION.md) when upgrading.

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

`Execute` manages activation contexts only. Use `Create` for an exclusively owned RCW that needs a handle, and follow the [lifetime rules](../README.md#lifetime-and-threading-rules).

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
