# Architecture Decision Records

This folder contains the Architecture Decision Records (ADRs) for `AdaskoTheBeAsT.Interop.COM`.

An ADR captures a single architecturally significant decision, the context in which it was made, the forces considered, and its consequences. We use a lightweight [MADR](https://adr.github.io/madr/)-inspired format.

## Index

| ID | Title | Status | Introduced in |
| -- | ----- | ------ | ------------- |
| [0001](0001-record-architecture-decisions.md) | Record Architecture Decisions | Accepted | v3.0.0 |
| [0002](0002-registration-free-com-via-activation-contexts.md) | Use Registration-Free COM via Activation Contexts | Accepted | v1.0.0 |
| [0003](0003-result-pattern-instead-of-exceptions.md) | Expose a Result Pattern Instead of Throwing from the Public API | Accepted | v1.0.0 |
| [0004](0004-static-executor-api.md) | Expose Functionality Through a Static `Executor` Class | Accepted | v1.0.0 |
| [0005](0005-internal-native-methods-wrapper.md) | Encapsulate Win32 P/Invoke in an Internal `NativeMethods` | Accepted | v1.0.0 |
| [0006](0006-strict-analyzer-stack.md) | Enforce a Strict Analyzer Stack with Warnings as Errors | Accepted | v1.0.0 |
| [0007](0007-multi-target-frameworks.md) | Multi-Target `netstandard2.0`, Modern .NET, and .NET Framework | Accepted | v2.0.0 / v3.0.0 |
| [0008](0008-libraryimport-on-net8-plus.md) | Use `LibraryImport` Source Generator on .NET 8+ | Accepted | v2.0.0 |
| [0009](0009-com-path-descriptor.md) | Introduce `ComPathDescriptor` for Multi-Context Activation | Accepted | v2.0.0 |
| [0010](0010-com-object-handle-lifetime-api.md) | Introduce `ComObjectHandle<T>` Lifetime API with Explicit Free | Accepted | v2.1.0 |
| [0011](0011-pump-sta-messages-after-com-calls.md) | Pump Pending STA Windows Messages Around COM Calls | Accepted | v2.1.0 |
| [0012](0012-com-object-handle-idisposable.md) | Make `ComObjectHandle<T>` Implement `IDisposable` | Accepted | v3.0.0 |
| [0013](0013-icomexecutor-abstraction.md) | Introduce `IComExecutor` Abstraction and `ComExecutor` Implementation | Accepted | v3.0.0 |
| [0014](0014-supported-os-platform-windows.md) | Annotate Public Surface with `[SupportedOSPlatform("windows")]` | Accepted | v3.0.0 |
| [0015](0015-ci-reusable-workflow.md) | Adopt a Shared Reusable GitHub Actions CI Workflow | Accepted | v3.0.0 |
| [0016](0016-remove-x86-platform-target.md) | Build the Managed Library as `AnyCPU` | Accepted | v3.0.0 |
| [0017](0017-no-dependency-injection-package.md) | Do Not Ship a Separate `DependencyInjection` Package (For Now) | Accepted | v3.0.0 |
| [0018](0018-eventsource-for-leaked-handles.md) | Diagnostic `EventSource` for Leaked `ComObjectHandle<T>` | Accepted | v3.0.0 |

## Conventions

- Each file is named `NNNN-kebab-case-title.md`, where `NNNN` is a zero-padded monotonically increasing number.
- Never rewrite an accepted ADR's decision once merged. If a later decision invalidates it, mark the original as `Superseded by ADR-NNNN` and create a new record.
- Keep each ADR short (ideally one or two pages). Link to code, commits, or issues rather than duplicating content.

## Template

See [`template.md`](template.md) for the skeleton used by all ADRs in this repository.
