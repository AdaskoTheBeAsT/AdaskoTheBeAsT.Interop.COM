using System;
using System.Runtime.InteropServices;
using System.Security;

namespace AdaskoTheBeAsT.Interop.COM;

[SuppressUnmanagedCodeSecurity]
#if NETSTANDARD2_0
internal static class NativeMethods
#endif
#if NET8_0_OR_GREATER
internal static partial class NativeMethods
#endif
{
    internal const int ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID = 0x004;

#if NETSTANDARD2_0
    [DllImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool ActivateActCtx(IntPtr hActCtx, out IntPtr lpCookie);
#endif

#if NET8_0_OR_GREATER
    [LibraryImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool ActivateActCtx(IntPtr hActCtx, out IntPtr lpCookie);
#endif

    [DllImport("Kernel32.dll", SetLastError = true)]
    internal static extern IntPtr CreateActCtxW(ref ActCtx pActCtx);

#if NETSTANDARD2_0
    [DllImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool DeactivateActCtx(int dwFlags, IntPtr lpCookie);
#endif

#if NET8_0_OR_GREATER
    [LibraryImport("Kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool DeactivateActCtx(int dwFlags, IntPtr lpCookie);
#endif

#if NETSTANDARD2_0
    [DllImport("Kernel32.dll", SetLastError = true)]
    internal static extern void ReleaseActCtx(IntPtr hActCtx);
#endif

#if NET8_0_OR_GREATER
    [LibraryImport("Kernel32.dll", SetLastError = true)]
    internal static partial void ReleaseActCtx(IntPtr hActCtx);
#endif
}
