# Changelog

Notable changes to `AdaskoTheBeAsT.Interop.COM`.

Dates shown for historical releases refer to repository tags, not NuGet publication dates.

See the [migration guide](MIGRATION.md) for required upgrade steps and the [README](README.md) for current usage.

## [4.0.0]

### Breaking changes

- Remove `netstandard2.0`, `net462`, `net47`, and `net471` assets. The supported targets are now `net8.0`, `net9.0`, `net10.0`, `net472`, `net48`, and `net481`.
- Enforce handle release on the creating managed and native thread, in reverse activation order. Wrong-thread, reentrant, and out-of-order release returns a failed `Result` before releasing the COM object.
- Return operational activation-context setup failures through `Result.Exception`. Callers that previously caught setup exceptions must also check the result. Invalid arguments still throw.
- Bound each automatic message pump to 256 messages instead of draining the queue indefinitely. Applications must not rely on an API call emptying their message queue.

### Added

- `ComExecutor(bool pumpPendingMessages)` for an immutable message-pumping policy.
- Additive `Executor.Execute` and `Executor.Create` overloads with a final `bool pumpPendingMessages` parameter, for both single and multiple manifests.
- Handles retain the policy selected at creation for `Free` and `Dispose`, including release through another executor. Existing signatures still enable pumping by default; `IComExecutor` has no new members.
- Source-built native lifetime tests for actual COM destruction, callbacks, alias invalidation, and reentrant release.
- Regression tests for activation failures, release ordering, thread ownership, Unicode messages, `WM_QUIT`, bounded pumping, and pump policies.

### Fixed

- Release already-created activation contexts when a later manifest fails to load, and clean up after partial activation failures.
- Validate null actions and null entries in descriptor collections before allocating native resources.
- Use the COM DLL's containing directory, rather than the DLL path itself, as the assembly directory for activation-context probing.
- Check native deactivation results and preserve unfinished handle cleanup state for retry. Combine work and cleanup failures in `AggregateException` when both occur.
- Keep leak diagnostics enabled after rejected or incomplete disposal instead of suppressing finalization unconditionally.
- Preserve `WM_QUIT` and its exit code for the host message loop.
- Use Unicode message retrieval and dispatch on .NET Framework as well as modern .NET.
- Select native test payloads by explicit process architecture and isolate managed outputs by architecture to avoid reusing incompatible assemblies.

### Build and documentation

- Update the SDK selected by `global.json` to `10.0.401` and select Microsoft.Testing.Platform for `dotnet test`.
- Update analyzer and test dependencies, including xUnit v3 across the remaining test targets.
- Document caller-owned COM initialization, STA setup, message processing, and exclusive RCW ownership.
- Reorganize the README and add standalone changelog, migration, advanced usage, and native build guides.
- Add [ADR-0019](docs/adr/0019-safe-activation-context-cleanup.md) and [ADR-0020](docs/adr/0020-host-owned-message-pumping.md).

## [3.0.0] - 2026-04-19

### Added

- `IComExecutor` and its `ComExecutor` implementation for dependency injection and testing. Existing static `Executor` entry points remain available.
- `IDisposable` on `ComObjectHandle<T>` and public `IsReleased` state.
- Diagnostic-only handle finalization and the internal `AdaskoTheBeAsT.Interop.COM` event source: `HandleLeaked` (ID 1) and `HandleReleaseFailed` (ID 2).
- Windows platform annotations on `Executor`, `IComExecutor`, `ComExecutor`, and `ComObjectHandle<T>` for modern .NET targets. Consumers may need to address new `CA1416` warnings.
- Dedicated .NET Framework targets: `net462`, `net47`, `net471`, `net472`, `net48`, and `net481`, alongside existing `netstandard2.0`, `net8.0`, `net9.0`, and `net10.0` assets.
- Shared GitHub Actions CI workflow, architecture decision records, and expanded managed tests.

### Changed

- Build the managed library as AnyCPU instead of x86. Native COM components and their host processes must still have matching architectures.

### Fixed

- Compatibility of native declarations and exception code across the target framework matrix.

## [2.1.0] - 2026-04-07

### Added

- `Executor.Create<T>` / `Executor.Free<T>`, `ComObjectHandle<T>`, and `ComObjectCreationResult<T>` for explicit object lifetime management.
- Automatic pending Windows message pumping around COM operations.
- Lifetime and STA scheduling examples in the README.

The tagged package targets `netstandard2.0`, `net8.0`, `net9.0`, and `net10.0`; its managed assembly is x86.

## Earlier history

Older README and ADR entries use version labels that do not line up consistently with repository tags. They are preserved here as development milestones rather than assigned unverified release versions or dates:

- **2025-08-10 to 2025-11-23:** multi-context activation through `ComPathDescriptor`, modern .NET targets, source-generated P/Invoke on modern .NET, and expanded package documentation. The repository's [`v1.0.0` tag](https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/tree/v1.0.0) points to the 2025-11-23 commit, not the initial implementation. There is no `v2.0.0` tag.
- **2023-12-09:** initial implementation with static `Executor`, `Result`, a `netstandard2.0` target, and internal Win32 activation-context wrappers.

The [architecture decision index](docs/adr/README.md) provides additional historical context.

[4.0.0]: https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/compare/v3.0.0...v4.0.0
[3.0.0]: https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/compare/v2.1.0...v3.0.0
[2.1.0]: https://github.com/AdaskoTheBeAsT/AdaskoTheBeAsT.Interop.COM/releases/tag/v2.1.0
