# ADR-0016: Build the Managed Library as `AnyCPU`

- Status: Accepted
- Date: 2026-04-19
- Deciders: @AdaskoTheBeAsT
- Release: v3.0.0
- Related commit(s): v3.0.0 changes

## Context

From v1.0.0 until v3.0.0 the managed library carried `<PlatformTarget>x86</PlatformTarget>` in its `.csproj`. That was a pragmatic choice because the bundled sample COM component (`src/NativeCOM`) is built as a 32-bit DLL and the unit tests load it directly. Forcing the managed assembly to 32-bit made the in-repo test loop "just work" without requiring consumers to understand process bitness.

The side effect, however, is that the NuGet package was usable only from 32-bit host processes. Real-world COM components increasingly ship in 64-bit form (Office 2016+, many enterprise line-of-business automation APIs, SQL Server OLE DB providers, modern shell extensions, and every driver-related COM surface). A 32-bit-only managed library cannot be loaded into a 64-bit host at all, so downstream users either could not reference the package or had to fork it and strip the `PlatformTarget`.

The activation-context primitives the library wraps (`CreateActCtx`, `ActivateActCtx`, `DeactivateActCtx`, `ReleaseActCtx`, `PeekMessage`, `TranslateMessage`, `DispatchMessage`) exist on both 32-bit and 64-bit Windows with identical contracts. The one bitness-sensitive detail — the size of the `ACTCTX` struct — is already guarded by a runtime check (`IntPtr.Size == 4 ? 0x20 : 0x38`) that raises `ActCtxWrongSizeException` on mismatch (see ADR-0005).

## Decision

Remove the `<PlatformTarget>x86</PlatformTarget>` element from `src/AdaskoTheBeAsT.Interop.COM/AdaskoTheBeAsT.Interop.COM.csproj`. The managed library now builds as `AnyCPU` on every target framework, which means:

- Loaded into a 32-bit host process, the CLR JITs it as 32-bit and the existing `0x20` `cbSize` check passes.
- Loaded into a 64-bit host process, the CLR JITs it as 64-bit and the `0x38` `cbSize` check passes.

The runtime bitness is the caller's choice, not ours.

The test project (`test/unit/AdaskoTheBeAsT.Interop.COM.Test`) keeps its `<PlatformTarget>x86</PlatformTarget>` because it still references the 32-bit sample `NativeCOM.dll` directly. That is an artefact of the test harness, not a constraint on the shipped library.

## Consequences

- Positive: the NuGet package is usable from 64-bit hosts — the common deployment scenario — without any workaround.
- Positive: consumers who need 32-bit behaviour can still get it simply by running their host as 32-bit; nothing about the library forbids it.
- Positive: no code change is required; the `cbSize` guard already handles both architectures correctly.
- Negative: bugs in `CreateActCtx` marshalling that only manifest in 64-bit processes (historically rare but possible) are now reachable through the library. The existing `ActCtxWrongSizeException` will surface such cases explicitly.
- Negative: contributors testing against a 64-bit sample COM component must now build such a component themselves; the in-repo `src/NativeCOM` project remains 32-bit.

## Alternatives Considered

- Ship two NuGet packages (`.x86` and `.x64`): rejected as a needless split; a single `AnyCPU` managed assembly covers both.
- Parameterise `PlatformTarget` via an MSBuild property: rejected because it complicates the packaging story and gives consumers no advantage over `AnyCPU`.
- Keep `PlatformTarget=x86` but expose a 64-bit switch: rejected because such a switch would just be `AnyCPU` under a different name.

## References

- `src/AdaskoTheBeAsT.Interop.COM/AdaskoTheBeAsT.Interop.COM.csproj`
- `src/AdaskoTheBeAsT.Interop.COM/Executor.cs` (`PrepareContext`, `cbSize` check)
- [ADR-0005](0005-internal-native-methods-wrapper.md)
