# ADR-0002: Use Registration-Free COM via Activation Contexts

- Status: Accepted
- Date: 2023-12-09
- Deciders: @AdaskoTheBeAsT
- Release: v1.0.0
- Related commit(s): `fb360c4` (initial commit)

## Context

Many deployment scenarios cannot rely on registry-based COM activation: unattended installs on locked-down machines, portable applications, CI agents, and side-by-side installations that must not touch `HKCR`. Windows supports exactly this use case through activation contexts driven by SxS manifests (`CreateActCtx` / `ActivateActCtx`), but the raw Win32 API is cumbersome, error-prone, and requires careful teardown semantics.

## Decision

The library's single purpose is to expose registration-free COM activation on Windows. It wraps `CreateActCtx`, `ActivateActCtx`, `DeactivateActCtx`, and `ReleaseActCtx` behind a small C# facade so that callers supply only the COM DLL path and the manifest path. Every higher-level API (single-context `Execute`, multi-context `Execute`, `Create`, `Free`) is built on top of this primitive.

## Consequences

- Positive: callers can use COM components without admin rights or registry changes, which is the whole reason this library exists.
- Positive: the scope is bounded and well-defined, which keeps the public surface small.
- Negative: the library is inherently Windows-only; it will never run on Linux or macOS (see ADR-0014 for the platform annotation).
- Negative: manifests must be authored or generated separately; the library does not attempt to synthesize them.

## Alternatives Considered

- Wrapping `DllGetClassObject` directly: rejected because it bypasses the SxS infrastructure that keeps multiple COM versions isolated.
- Shipping a manifest generator: rejected as out of scope; third-party tools (e.g. ManifestMaker) already solve this.
- Binding via `dynamic`/`IDispatch`: rejected because it sacrifices compile-time safety and is orthogonal to the activation-context problem.

## References

- [Registration-free COM interop (Microsoft Learn)](https://learn.microsoft.com/en-us/dotnet/framework/interop/registration-free-com-interop)
- [Activation Contexts (Microsoft Learn)](https://learn.microsoft.com/en-us/windows/win32/sbscs/activation-contexts)
