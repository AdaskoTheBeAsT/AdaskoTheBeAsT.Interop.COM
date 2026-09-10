using System;

namespace AdaskoTheBeAsT.Interop.COM;

// An operation-scoped native boundary permits deterministic failure tests without global hooks.
internal class ActivationContextApi
{
    internal static ActivationContextApi Instance { get; } = new();

    internal virtual IntPtr CreateActCtx(ref ActCtx context) => NativeMethods.CreateActCtx(ref context);

    internal virtual bool ActivateActCtx(IntPtr handle, out IntPtr cookie)
        => NativeMethods.ActivateActCtx(handle, out cookie);

    internal virtual bool DeactivateActCtx(IntPtr cookie) => NativeMethods.DeactivateActCtx(0, cookie);

    internal virtual bool GetCurrentActCtx(out IntPtr handle) => NativeMethods.GetCurrentActCtx(out handle);

    internal virtual void ReleaseActCtx(IntPtr handle) => NativeMethods.ReleaseActCtx(handle);
}
