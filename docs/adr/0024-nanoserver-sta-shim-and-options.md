# ADR-0024: Provide an In-Process User-Mode STA & Win32 Message-Queue Shim and `NanoServerOptions` Toggles on Windows Nano Server

- Status: Accepted
- Date: 2026-10-27
- Deciders: @AdaskoTheBeAsT
- Release: v4.2.0
- Extends [ADR-0011](0011-pump-sta-messages-after-com-calls.md), [ADR-0019](0019-safe-activation-context-cleanup.md), [ADR-0020](0020-host-owned-message-pumping.md), and [ADR-0023](0023-nanoserver-inproc-class-factory-fallback.md)

## Context

In v4.1.0 ([ADR-0023](0023-nanoserver-inproc-class-factory-fallback.md)), `AdaskoTheBeAsT.Interop.COM` introduced a manifest-driven `DllGetClassObject` + `CoRegisterClassObject` fallback that enabled registration-free COM activation on Windows Nano Server 2025 (`ltsc2025`), but only on `MTA` threads:

1. **Single-Threaded Apartment (`STA` / `COINIT_APARTMENTTHREADED`) fails with `E_NOTIMPL` (`0x80004001`) on unshimmed Nano Server:**
   - Inside `combase.dll`, `IncrementMTAUsage` succeeds, whereas `InitMainThreadWnd` and `ThreadFirstInitialize` check `IsRegisterClassWPresent()` / `ext-ms-win-rtcore-ntuser-window-ext-l1-1-0` (`win32u.dll`) and create the hidden `OleMainThreadWndClass` (`HWND_MESSAGE`) window for cross-apartment marshaling and STA synchronization.
   - Because Nano Server omits `win32u.dll` and full `user32.dll`, `CoInitializeEx(NULL, COINIT_APARTMENTTHREADED)` and `Thread.SetApartmentState(ApartmentState.STA)` fail with `0x80004001` (`E_NOTIMPL`).
2. **Legacy Win32 forwarder path resolution:**
   - Native C++/ATL COM servers typically link against `USER32.dll`, `OLE32.dll`, `ADVAPI32.dll`, and `KERNEL32.dll`, which reside under `%SystemRoot%\System32\forwarders` (or `%SystemRoot%\SysWOW64\forwarders` on 32-bit WOW64) on Nano Server.
3. **Coexistence with OS-level container shims (`nanoserver_sta_shim.dll`) and testability:**
   - Custom Nano Server images (`adaskothebeast/nanoserver`) pre-install a native C++ OS-level STA shim (`nanoserver_sta_shim.dll`) into `System32` / `SysWOW64`.
   - Callers and test suites need a deterministic way to toggle both the OS-level native shim and the library-level in-process STA / SxS shims on or off via environment variables and runtime options (`NanoServerOptions`).

## Decision

1. **In-Process User-Mode Win32 Window & Message-Queue Manager (`NanoServerWindowManager`):**
   - Implement an in-memory window class table, `HWND` table, per-thread FIFO posted-message queue (`Win32Event` wake signaling), synchronous `SendMessageW` dispatch to registered `lpfnWndProc` callbacks, timer table, and window-property table (`RegisterClassExW`, `CreateWindowExW`, `DestroyWindow`, `DefWindowProcW`, `PeekMessageW`, `GetMessageW`, `TranslateMessage`, `DispatchMessageW`, `PostMessageW`, `PostThreadMessageW`, `SendMessageW`, `SendMessageTimeoutW`, `MsgWaitForMultipleObjectsEx`, `SetTimer`, `KillTimer`, `SetPropW`, `GetPropW`, `RemovePropW`).
2. **Import Address Table (IAT) Redirection & Coordination (`NanoServerStaShim`):**
   - On Nano Server (`win32u.dll` absent in `%SystemRoot%\System32`), append `%SystemRoot%\System32\forwarders` (or `SysWOW64\forwarders` in 32-bit processes) to `PATH` and preload `combase.dll`, `user32.dll`, `ole32.dll`, `advapi32.dll`, and `oleaut32.dll`.
   - Patch the API-set presence and window/message-queue IAT slots in `combase.dll`, `ole32.dll`, and `user32.dll` (`VirtualProtect` + atomic pointer write).
   - Coordinate with `nanoserver_sta_shim.dll` (`NanoServerStaShim_Refresh` and `NanoServerStaShim_RestoreOriginal`) so `NanoServerStaShim` captures clean unhooked thunks and defers to the native OS-level shim by default when present, while seamlessly taking over when the OS-level shim is absent or disabled (`NANOSERVER_DISABLE_STA_SHIM=1` or `NanoServerOptions.ForceInProcessStaShim = true`).
   - Start a dedicated background `NanoServerMainStaHost` thread outside the CLR module-initializer lock (`EnsureInitialized(waitForHostThread: true)` on first entry point after `[ModuleInitializer]`) so `combase.dll`'s `g_hwndOleMainThread` remains alive across user STA threads.
3. **Public Configuration API (`NanoServerOptions`):**
   - Expose `NanoServerOptions.EnableSxsFallback`, `NanoServerOptions.EnableStaShim`, `NanoServerOptions.ForceInProcessStaShim`, `NanoServerOptions.IsInProcessStaShimActive`, `NanoServerOptions.RefreshStaState()`, and `NanoServerOptions.ResetToEnvironmentDefaults()`, backed by `ADASKO_NANOSERVER_DISABLE_SXS_FALLBACK`, `ADASKO_NANOSERVER_DISABLE_STA_SHIM`, and `NANOSERVER_DISABLE_STA_SHIM`.

## Consequences

- Positive: Full `ApartmentState.STA` (`COINIT_APARTMENTTHREADED`) and bounded message pumping (`pumpPendingMessages: true`) work on Windows Nano Server 2025 (`x86` and `x64`) both on custom `adaskothebeast/nanoserver` images and on stock Microsoft Nano Server images.
- Positive: Zero impact on full Windows Desktop or Windows Server Core (`win32u.dll` present), where `NanoServerStaShim` immediately no-ops.
- Positive: Both OS-level and library-level shims can be independently toggled for testing and diagnostics.

## References

- `src/AdaskoTheBeAsT.Interop.COM/NanoServerOptions.cs`
- `src/AdaskoTheBeAsT.Interop.COM/NanoServerStaShim.cs`
- `src/AdaskoTheBeAsT.Interop.COM/NanoServerWindowManager.cs`
- `src/AdaskoTheBeAsT.Interop.COM/ActivationContextApi.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/NanoServerFallbackTest.cs`
