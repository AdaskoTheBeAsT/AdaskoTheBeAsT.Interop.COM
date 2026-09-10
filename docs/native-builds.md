# Native builds and architecture

See the [README](../README.md) for managed builds and usage, and the [migration guide](../MIGRATION.md) for the AnyCPU change introduced in 3.0.

## Match the process and native payload

The managed `AdaskoTheBeAsT.Interop.COM` library is AnyCPU. Native COM DLLs and their dependencies are not.

| Host process | Native DLL | Manifest `processorArchitecture` | Managed interop assembly |
| --- | --- | --- | --- |
| x86 | x86 | `x86` | Agnostic or x86 |
| x64 | x64 | `amd64` | Agnostic or x64 |

Set the consuming application's architecture explicitly. A 64-bit process cannot load a 32-bit in-process COM server.

The repository's test project selects native payloads from the explicit `PlatformTarget`, runtime identifier, or platform, with defaults of x86 for .NET Framework and x64 for modern .NET. Only x86 and x64 native test hosts are supported.

## Generate an architecture-neutral interop assembly

For a type library whose imported declarations work on both architectures, generate a managed interop assembly with `TlbImp.exe /machine:Agnostic`:

```powershell
# From the repository root. Adjust the installed Windows SDK tool path if needed.
$tlbimp = "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\TlbImp.exe"
& $tlbimp .\x86\Debug\NativeCOM.dll `
    /out:.\x86\Debug\Interop.NativeCOM.dll `
    /namespace:NativeCOM `
    /machine:Agnostic
```

This command regenerates the sample interop output. Do not run it over local changes you need to preserve. The repository's `generatelib.bat` generates agnostic assemblies from both native payloads; the tests reference the x86-directory copy for both process architectures.

An agnostic assembly still contains generated interop code and metadata. The flag does not prove that every type-library signature or native marshaling contract is portable. Verify the interfaces on both architectures, especially if the vendor exposes architecture-dependent types.

Inspect the flags with the Windows SDK `CorFlags` tool, without modification options:

```powershell
corflags .\x86\Debug\Interop.NativeCOM.dll
```

For the agnostic sample, expect `ILONLY: 1`, `32BITREQ: 0`, and `32BITPREF: 0`.

## Deploy from a consuming project

For a Windows-only modern .NET project, one simple layout is:

```text
MyApp.csproj
lib/
  Interop.NativeCOM.dll
  x86/
    NativeCOM.dll
    NativeCOM.manifest
  x64/
    NativeCOM.dll
    NativeCOM.manifest
```

Copy the matching repository manifest into each directory under the common name `NativeCOM.manifest`. Then configure the application:

```xml
<PropertyGroup>
  <TargetFramework>net10.0-windows</TargetFramework>
  <PlatformTarget>x64</PlatformTarget>
</PropertyGroup>

<ItemGroup>
  <Reference Include="Interop.NativeCOM">
    <HintPath>lib\Interop.NativeCOM.dll</HintPath>
    <Private>true</Private>
  </Reference>
  <None Update="lib\$(PlatformTarget)\NativeCOM.dll">
    <TargetPath>NativeCOM.dll</TargetPath>
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    <CopyToPublishDirectory>PreserveNewest</CopyToPublishDirectory>
  </None>
  <None Update="lib\$(PlatformTarget)\NativeCOM.manifest">
    <TargetPath>NativeCOM.manifest</TargetPath>
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    <CopyToPublishDirectory>PreserveNewest</CopyToPublishDirectory>
  </None>
</ItemGroup>
```

Change `PlatformTarget` to `x86` for a 32-bit host. Also deploy the component's required native dependencies. This example assumes the SDK's default `None` items are enabled and does not replace vendor-specific packaging requirements.

## When separate interop assemblies are necessary

Use separate interop outputs if the type libraries or imported signatures differ by architecture, or your distribution policy requires architecture-specific assemblies:

1. Build the native component and type library for each architecture.
2. Ensure the MIDL target environment matches the intended type library: `Win32` or `X64`.
3. Run `TlbImp` with `/machine:X86` or `/machine:X64` against the matching type library.
4. Select the corresponding interop reference along with the native DLL and manifest in the consuming project.
5. Test the actual interface marshaling on both architectures.

The repository sample has `Win32` MIDL settings in some x64 configurations and uses agnostic imports. Do not assume its embedded type library is ready for `/machine:X64`. Check both Debug and Release settings before adopting per-architecture imports.

`TlbImp` error `TI2010` can indicate that the selected machine type is incompatible with the input type library. Fix the native/type-library build configuration rather than merely relabeling the managed output.

## Rebuild and test the repository samples

The sample native project is [`src/NativeCOM/NativeCOM.vcxproj`](../src/NativeCOM/NativeCOM.vcxproj). Build it with Visual Studio/MSBuild and the C++ desktop workload, ATL, and a Windows SDK. Use `Win32` for the x86 native configuration and `x64` for the x64 configuration. Ensure the resulting payloads are placed where the [test project](../test/unit/AdaskoTheBeAsT.Interop.COM.Test/AdaskoTheBeAsT.Interop.COM.Test.csproj) expects them.

The repository tracks the sample DLLs and interop assemblies in `x86/Debug` and `x64/Debug`. Rebuilding may modify these tracked files. Review those changes separately from managed source changes; do not register the DLLs to make registration-free tests pass.

The separate [`NativeLifetimeProbe`](../test/native/NativeLifetimeProbe.vcxproj) fixture builds automatically during managed test builds, into per-framework/per-architecture intermediate directories. It checks actual native destruction and callback behavior. Architecture tests read PE headers to verify that the copied DLLs match the test process.

For the current Microsoft.Testing.Platform configuration:

```powershell
dotnet test --project test/unit/AdaskoTheBeAsT.Interop.COM.Test/AdaskoTheBeAsT.Interop.COM.Test.csproj -f net10.0 --arch x86
dotnet test --project test/unit/AdaskoTheBeAsT.Interop.COM.Test/AdaskoTheBeAsT.Interop.COM.Test.csproj -f net10.0 --arch x64
dotnet test --project test/unit/AdaskoTheBeAsT.Interop.COM.Test/AdaskoTheBeAsT.Interop.COM.Test.csproj -f net472 -p:PlatformTarget=x64
```

Install the corresponding runtimes before running. The managed library remains AnyCPU; the test project isolates its own outputs and compiler caches by architecture.

## References

- [TlbImp.exe reference](https://learn.microsoft.com/en-us/dotnet/framework/tools/tlbimp-exe-type-library-importer)
- [Registration-free COM interop](https://learn.microsoft.com/en-us/dotnet/framework/interop/registration-free-com-interop)
- [Sample x86 manifest](../manifest/NativeCOM/NativeCOM.manifest)
- [Sample x64 manifest](../manifest/NativeCOM/NativeCOM.x64.manifest)
