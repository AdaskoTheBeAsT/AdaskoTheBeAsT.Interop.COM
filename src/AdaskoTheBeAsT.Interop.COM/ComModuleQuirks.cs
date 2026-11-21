using System;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1121 // Use built-in type alias
/// <summary>
/// Explicit, opt-in per-module compatibility helpers for legacy native COM binaries that exhibit
/// non-conforming Win32 COM lifetime behavior inside shared managed worker apartments.
/// </summary>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public static class ComModuleQuirks
{
    /// <summary>
    /// Redirects <c>CoUninitialize</c> and <c>OleUninitialize</c> Import Address Table (IAT) and delay-import
    /// entries inside the single native PE module identified by <paramref name="moduleHandle"/> to a static
    /// native W^X no-op stub (<c>ret</c>).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Certain legacy C++/VB6 COM DLLs call an extra, unbalanced <c>CoUninitialize()</c> or <c>OleUninitialize()</c>
    /// during internal error or teardown paths, prematurely tearing down the hosting STA/MTA worker apartment while
    /// managed Runtime Callable Wrappers (RCWs) are still active on the thread.
    /// </para>
    /// <para>
    /// This method patches only <paramref name="moduleHandle"/> (inspecting <c>ole32.dll</c>, <c>combase.dll</c>,
    /// and <c>api-ms-win-core-com-l1-*</c> import descriptors) using temporary <c>PAGE_READWRITE</c> protection
    /// and restoring the original page protection afterward. Protected OS and CLR runtime modules are rejected.
    /// </para>
    /// </remarks>
    /// <param name="moduleHandle">The loaded <c>HMODULE</c> base address of the isolated worker COM DLL to patch.</param>
    /// <returns>The number of IAT/delay-import slots redirected in <paramref name="moduleHandle"/>.</returns>
    /// <exception cref="ArgumentException">Thrown when <paramref name="moduleHandle"/> is zero or not a loaded PE module base address.</exception>
    /// <exception cref="InvalidOperationException">Thrown when <paramref name="moduleHandle"/> refers to a protected OS or CLR module.</exception>
    public static int SuppressUninitializeImports(IntPtr moduleHandle)
        => NanoServerOle2Shim.SuppressUninitializeImports(moduleHandle);

    /// <summary>
    /// Restores any <c>CoUninitialize</c> and <c>OleUninitialize</c> IAT/delay-import slots previously redirected
    /// in <paramref name="moduleHandle"/> by <see cref="SuppressUninitializeImports"/>.
    /// </summary>
    /// <param name="moduleHandle">The loaded <c>HMODULE</c> base address whose suppressed imports should be restored.</param>
    /// <returns>The number of IAT/delay-import slots restored, or <c>0</c> if the module had no suppressed slots.</returns>
    /// <exception cref="ArgumentException">Thrown when <paramref name="moduleHandle"/> is zero.</exception>
    public static int RestoreUninitializeImports(IntPtr moduleHandle)
        => NanoServerOle2Shim.RestoreUninitializeImports(moduleHandle);
}
#pragma warning restore SA1121
