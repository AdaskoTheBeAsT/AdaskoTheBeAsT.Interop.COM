# ADR-0022: Preserve Explicit `IntPtr`/`UIntPtr` Win32 Signatures and Isolate Multi-Architecture Solution Builds

- Status: Accepted
- Date: 2026-10-24
- Deciders: @AdaskoTheBeAsT
- Release: v4.0.1
- Extends [ADR-0005](0005-internal-native-methods-wrapper.md) and [ADR-0006](0006-strict-analyzer-stack.md)

## Context

Two build and analyzer challenges surfaced when performing full non-incremental rebuilds (`dotnet build --no-incremental`) with the updated analyzer stack (`StyleCop.Analyzers 1.3.151-beta`, `SonarAnalyzer.CSharp 10.35.0.4138`, `AdaskoTheBeAsT.CodeCracker.CSharp 1.1.0.270`, `Meziantou.Analyzer 3.0.316`):

1. **`SA1121` on `IntPtr` and `UIntPtr`:** `StyleCop.Analyzers 1.3.151-beta` flags `IntPtr` and `UIntPtr` under `SA1121` ("Use built-in type alias"), preferring C# `nint` and `nuint`. However, `ComObjectHandle<T>` exposes `IList<IntPtr> ActivationContextHandles` and `IList<IntPtr> ActivationCookies` as part of its internal/documented contract, and Win32 interop structs (`ActCtx`, `NativeMethods.MSG`) and P/Invoke declarations map directly to Windows SDK `HANDLE`, `HWND`, `WPARAM`, `LPARAM`, and `ULONG_PTR` types where `IntPtr`/`UIntPtr` is clearer and avoids churn across multi-targeted files.
2. **Legacy FxCop `CA1060` on .NET Framework test P/Invokes:** Test fixtures that declare supplemental `[DllImport]` entry points (`PostThreadMessageW`, `NativeLifetimeProbe.dll`) must reside in classes named `NativeMethods`, `SafeNativeMethods`, or `UnsafeNativeMethods` on `net472`/`net48`/`net481` without shadowing `AdaskoTheBeAsT.Interop.COM.NativeMethods`.
3. **Parallel `GeneratePackageOnBuild` race (`NU5026`) on solution rebuilds:** In `AdaskoTheBeAsT.Interop.COM.Test.csproj`, unconditionally removing `Platform` in `GlobalPropertiesToRemove` on the `ProjectReference` caused MSBuild to evaluate `AdaskoTheBeAsT.Interop.COM.csproj` twice in parallel during `.slnx` builds: once with `Platform=AnyCPU` from the solution, and once with `Platform` unset from the test project. Under `--no-incremental`, the second evaluation's `CoreClean` could delete compiled outputs while the first evaluation's `Pack` target was reading them.

## Decision

- Keep explicit `IntPtr` and `UIntPtr` types across Win32 structs, P/Invoke declarations, activation-context helpers, and test fixtures, and bracket those interop files with `#pragma warning disable SA1121` / `#pragma warning restore SA1121` (matching `ActCtx.cs`).
- Name test-only P/Invoke container classes `SafeNativeMethods` so they satisfy `CA1060` on .NET Framework while allowing unqualified access to `AdaskoTheBeAsT.Interop.COM.NativeMethods`.
- In `AdaskoTheBeAsT.Interop.COM.Test.csproj`, remove `Platform` from `<ProjectReference>` global properties only when `$(Platform)` is non-empty and not `AnyCPU` (`Condition="'$(Platform)' != '' AND '$(Platform)' != 'AnyCPU'"`), so default `AnyCPU` solution builds share a single MSBuild evaluation of `AdaskoTheBeAsT.Interop.COM.csproj`.

## Consequences

- Positive: `dotnet build AdaskoTheBeAsT.Interop.COM.slnx -c Debug --no-incremental` and `-c Release --no-incremental` build `AdaskoTheBeAsT.Interop.COM` exactly once per configuration, eliminating the `NU5026` parallel `Pack`/`Clean` race and cutting library compile time in half during solution builds.
- Positive: Win32 interop signatures remain consistent across all six target frameworks without `nint`/`IntPtr` mixed notation, and test files no longer need `global::AdaskoTheBeAsT.Interop.COM.NativeMethods` disambiguation.
- Negative: files that use `IntPtr`/`UIntPtr` require an explicit `#pragma warning disable SA1121` header and matching `#pragma warning restore SA1121` footer.

## Alternatives Considered

- Replacing all `IntPtr`/`UIntPtr` usages with `nint`/`nuint`: rejected because it adds noise to Win32 handle/cookie APIs and XML documentation without improving runtime behavior or type safety.
- Disabling `SA1121` globally in `AdaskoTheBeAsT.ruleset` or `.editorconfig`: rejected so built-in C# aliases (`int`, `string`, `bool`, `uint`, etc.) remain enforced everywhere else.
- Disabling `GeneratePackageOnBuild` in `AdaskoTheBeAsT.Interop.COM.csproj`: rejected because automatic package generation during build is part of the repository's standard workflow and catches packaging issues early.

## References

- `src/AdaskoTheBeAsT.Interop.COM/ActCtx.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ActivationContextApi.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ComObjectHandle.cs`
- `src/AdaskoTheBeAsT.Interop.COM/Executor.cs`
- `src/AdaskoTheBeAsT.Interop.COM/NativeMethods.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/AdaskoTheBeAsT.Interop.COM.Test.csproj`
