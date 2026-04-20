using System;
using System.Runtime.CompilerServices;

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Internal argument-validation helper. Uses <c>ArgumentNullException.ThrowIfNull</c> when available
/// (net8.0+) and falls back to the classic <c>if-null/throw</c> pattern on older TFMs (netstandard2.0,
/// .NET Framework).
/// </summary>
internal static class ThrowHelper
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
#pragma warning disable S3236 // paramName is forwarded from the real caller, not the local 'argument' expression
    public static void ThrowIfNull<T>(T? argument, string paramName)
#pragma warning restore S3236
        where T : class
    {
#if NET8_0_OR_GREATER
#pragma warning disable S3236
        ArgumentNullException.ThrowIfNull(argument, paramName);
#pragma warning restore S3236
#else
#pragma warning disable CA1510, RCS1256
        if (argument is null)
        {
            throw new ArgumentNullException(paramName);
        }
#pragma warning restore CA1510, RCS1256
#endif
    }
}
