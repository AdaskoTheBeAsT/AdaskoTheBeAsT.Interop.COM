# AdaskoTheBeAsT.Interop.COM

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A lightweight .NET library for **registration-free COM interop** that enables you to use COM objects without registering them in the Windows registry. Execute COM methods in a Single Threaded Apartment (STA) with proper activation context management.

## Why Use This Library?

- 🚀 **No COM Registration Required** - Use COM DLLs directly without `regsvr32` or admin rights
- 🎯 **Manifest-Based Activation** - Leverage side-by-side (SxS) assembly loading
- 🔒 **Safe Resource Management** - Automatic cleanup of activation contexts
- 🧵 **STA Message Pumping** - Proper handling of COM message loops
- 🎨 **Clean API** - Simple, fluent interface for COM execution
- ✅ **Battle-Tested** - Comprehensive analyzer suite ensuring code quality

## Features

- ✨ Create COM objects registration-free using manifest files and release them explicitly when you are done
- 🏃 Execute COM methods in Single Threaded Apartment (STA)
- 📦 Support for single or multiple concurrent activation contexts
- 🎯 Proper exception handling with detailed result reporting
- 🔧 Cross-platform support: .NET Standard 2.0, .NET 8.0, .NET 9.0

## Installation

```bash
dotnet add package AdaskoTheBeAsT.Interop.COM
```

Or via NuGet Package Manager:
```
Install-Package AdaskoTheBeAsT.Interop.COM
```

## Quick Start

### Basic Usage

```csharp
using AdaskoTheBeAsT.Interop.COM;

// Define paths to your COM DLL and manifest
var comDllPath = @"C:\MyApp\MyCom.dll";
var manifestPath = @"C:\MyApp\MyCom.manifest";

// Execute COM method in STA context
var result = Executor.Execute(comDllPath, manifestPath, () =>
{
    // Create and use your COM object
    var comObject = new MyCom.MyComClass();
    var output = comObject.DoSomething();
    Console.WriteLine(output);
});

// Check execution result
if (result.Success)
{
    Console.WriteLine("COM execution successful!");
}
else
{
    Console.WriteLine($"Error: {result.Exception?.Message}");
}
```

### Real-World Example

```csharp
using AdaskoTheBeAsT.Interop.COM;

public class ComStringProcessor
{
    public string ConcatenateStrings(string str1, string str2)
    {
        var comDllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NativeCOM.dll");
        var manifestPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NativeCOM.manifest");
        
        string? result = null;
        
        var executionResult = Executor.Execute(comDllPath, manifestPath, () =>
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

### Create a COM Object and Release It Later

Use `Executor.Create(...)` when the COM object should outlive a single callback and you want to release it explicitly.

```csharp
using AdaskoTheBeAsT.Interop.COM;

var creation = Executor.Create(
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
    var release = Executor.Free(handle);

    if (!release.Success)
    {
        throw new InvalidOperationException("COM object release failed.", release.Exception);
    }
}
```

### Multiple COM Contexts

When you need to work with multiple COM components simultaneously:

```csharp
var descriptors = new List<ComPathDescriptor>
{
    new ComPathDescriptor(@"C:\MyApp\ComLib1.dll", @"C:\MyApp\ComLib1.manifest"),
    new ComPathDescriptor(@"C:\MyApp\ComLib2.dll", @"C:\MyApp\ComLib2.manifest"),
};

var result = Executor.Execute(descriptors, () =>
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
    private readonly string _comDllPath;
    private readonly string _manifestPath;

    public TimedComStringProcessor(string comDllPath, string manifestPath)
    {
        _comDllPath = comDllPath;
        _manifestPath = manifestPath;
    }

    public Task<string> ConcatAsync(string left, string right, CancellationToken cancellationToken)
        => SingleThreadedApartmentTask.RunWithTimeoutAsync(
            TimeSpan.FromSeconds(10),
            () =>
            {
                string output = string.Empty;

                var execution = Executor.Execute(_comDllPath, _manifestPath, () =>
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
    private readonly string _comDllPath;
    private readonly string _manifestPath;
    private ComObjectHandle<NativeCOM.StringConcatenatorClass>? _handle;
    private NativeCOM.StringConcatenatorClass? _concatenator;

    public ScheduledComStringProcessor(string comDllPath, string manifestPath)
    {
        _comDllPath = comDllPath;
        _manifestPath = manifestPath;
    }

    public Task InitializeAsync(CancellationToken cancellationToken)
        => SingleThreadedApartmentTaskScheduler.RunAsync(
            () =>
            {
                var creation = Executor.Create(
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
                        var release = Executor.Free(_handle);

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

## How It Works

1. **Activation Context Creation** - Creates Windows activation context from your manifest file
2. **Context Activation** - Activates the context to enable registration-free COM
3. **COM Execution** - Runs your code with the COM object in STA
4. **Message Pumping** - Processes Windows messages for COM callbacks
5. **Cleanup** - Automatically deactivates and releases contexts

## Creating COM Manifests

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
// Quick validation test
var result = Executor.Execute(@"C:\path\to\COM.dll", @"C:\path\to\COM.manifest", () =>
{
    Console.WriteLine("Manifest loaded successfully!");
});

if (!result.Success)
{
    Console.WriteLine($"Manifest error: {result.Exception?.Message}");
}
```

## API Reference

### Executor Class

#### `Execute(string comAssemblyPath, string manifestPath, Action action)`

Executes an action with a single COM activation context.

**Parameters:**
- `comAssemblyPath` - Full path to the COM DLL
- `manifestPath` - Full path to the manifest file
- `action` - Code to execute in the activation context

**Returns:** `Result` object containing success status and any exception

#### `Execute(ICollection<ComPathDescriptor> comPathDescriptors, Action action)`

Executes an action with multiple COM activation contexts.

**Parameters:**
- `comPathDescriptors` - Collection of COM DLL and manifest path pairs
- `action` - Code to execute with all contexts activated

**Returns:** `Result` object containing success status and any exception

#### `Create<T>(string comAssemblyPath, string manifestPath, Func<T> factory)`

Creates a COM object inside a registration-free activation context and returns a `ComObjectHandle<T>`.
Release the handle later by calling `Executor.Free(...)` on the same thread.

#### `Create<T>(ICollection<ComPathDescriptor> comPathDescriptors, Func<T> factory)`

Creates a COM object inside multiple registration-free activation contexts and returns a `ComObjectHandle<T>`.
Release the handle later by calling `Executor.Free(...)` on the same thread.

#### `Free<T>(ComObjectHandle<T> comObjectHandle)`

Releases a handle created by `Executor.Create(...)`.
This method is idempotent, so calling it for an already released handle succeeds without doing additional work.

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

```csharp
public sealed class ComObjectHandle<T>
    where T : class
{
    public T? ComObject { get; internal set; }
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

## Best Practices

1. ✅ **Always check `Result.Success`** before assuming COM execution succeeded
2. ✅ **Use absolute paths** for COM DLL and manifest files
3. ✅ **Keep manifests alongside DLLs** for easy deployment
4. ✅ **Handle exceptions** - COM calls can fail in various ways
5. ✅ **Test on target platform** - COM behavior can vary by Windows version
6. ⚠️ **Avoid long-running operations** in the action delegate
7. ⚠️ **Be aware of STA threading model** requirements

## Requirements

- **OS**: Windows (uses Windows-specific APIs: kernel32.dll, user32.dll)
- **Framework**: .NET Standard 2.0+ / .NET 8.0+ / .NET 9.0+
- **Architecture**: x86, x64 (defined in your project configuration)

## Troubleshooting

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

## Building from Source

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

## Code Quality

This project maintains high code quality standards with:
- 20+ static analyzers (StyleCop, Roslynator, SonarAnalyzer, etc.)
- Nullable reference types enabled
- Comprehensive documentation
- Unit test coverage
- Strict warning-as-error policy

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Ensure all analyzers pass
4. Add unit tests for new functionality
5. Submit a pull request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Credits

Created and maintained by [Adam Pluciński](https://github.com/AdaskoTheBeAsT)

## Related Resources

- [Registration-Free COM Interop](https://learn.microsoft.com/en-us/dotnet/framework/interop/registration-free-com-interop)
- [Activation Contexts](https://learn.microsoft.com/en-us/windows/win32/sbscs/activation-contexts)
- [Side-by-Side Assemblies](https://learn.microsoft.com/en-us/windows/win32/sbscs/about-side-by-side-assemblies-)

---

**Questions or issues?** Open an issue on [GitHub](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/issues)