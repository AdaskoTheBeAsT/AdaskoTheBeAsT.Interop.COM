using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Thread-affine disposable scope that temporarily sets both the Win32 thread locale (<c>SetThreadLocale</c>,
/// consumed by <c>oleaut32.dll</c> <c>VARIANT</c> coercion APIs such as <c>VarR8FromStr</c> and <c>IDispatch::Invoke</c>
/// when <c>LOCALE_USER_DEFAULT</c> is passed) and managed <see cref="CultureInfo.CurrentCulture"/> /
/// <see cref="CultureInfo.CurrentUICulture"/>, restoring the prior values upon <see cref="Dispose"/>.
/// </summary>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public sealed class ComAutomationLocaleScope : IDisposable
{
    /// <summary>
    /// Win32 invariant locale identifier (<c>LOCALE_INVARIANT</c>, <c>0x007F</c>).
    /// </summary>
    public const uint InvariantLocaleId = 0x007F;

    /// <summary>
    /// Win32 US English locale identifier (<c>en-US</c>, <c>0x0409</c>).
    /// </summary>
    public const uint UsEnglishLocaleId = 0x0409;

    /// <summary>
    /// Win32 UK English locale identifier (<c>en-GB</c>, <c>0x0809</c>).
    /// </summary>
    public const uint UkEnglishLocaleId = 0x0809;

    private readonly CultureInfo _previousCulture;
    private readonly CultureInfo _previousUiCulture;
    private bool _disposed;

    private ComAutomationLocaleScope(
        int ownerThreadId,
        uint previousThreadLocaleId,
        CultureInfo previousCulture,
        CultureInfo previousUiCulture)
    {
        OwnerThreadId = ownerThreadId;
        PreviousThreadLocaleId = previousThreadLocaleId;
        _previousCulture = previousCulture;
        _previousUiCulture = previousUiCulture;
    }

    /// <summary>
    /// Gets the managed thread identifier (<see cref="Environment.CurrentManagedThreadId"/>) of the thread that entered this scope.
    /// </summary>
    public int OwnerThreadId { get; }

    /// <summary>
    /// Gets the Win32 thread locale identifier (<c>LCID</c>) that was active on <see cref="OwnerThreadId"/> before this scope was entered.
    /// </summary>
    public uint PreviousThreadLocaleId { get; }

    /// <summary>
    /// Enters a thread-affine locale scope on the current thread, applying <paramref name="threadLocaleId"/> via Win32
    /// <c>SetThreadLocale</c> and setting <see cref="CultureInfo.CurrentCulture"/> and <see cref="CultureInfo.CurrentUICulture"/>
    /// to <paramref name="managedCulture"/> (or <see cref="CultureInfo.InvariantCulture"/> when <see langword="null"/>).
    /// </summary>
    /// <param name="threadLocaleId">The Win32 <c>LCID</c> to apply via <c>SetThreadLocale</c>. Defaults to <see cref="UsEnglishLocaleId"/> (<c>0x0409</c>).</param>
    /// <param name="managedCulture">The managed <see cref="CultureInfo"/> to apply, or <see langword="null"/> to use <see cref="CultureInfo.InvariantCulture"/>.</param>
    /// <returns>A disposable <see cref="ComAutomationLocaleScope"/> that restores the prior Win32 <c>LCID</c> and managed cultures on <see cref="Dispose"/>.</returns>
    /// <exception cref="Win32Exception">Thrown when Win32 <c>SetThreadLocale</c> rejects <paramref name="threadLocaleId"/>.</exception>
    public static ComAutomationLocaleScope Enter(
        uint threadLocaleId = UsEnglishLocaleId,
        CultureInfo? managedCulture = null)
    {
        var ownerThreadId = Environment.CurrentManagedThreadId;
        var previousLcid = NativeMethods.GetThreadLocale();
        var previousCulture = CultureInfo.CurrentCulture;
        var previousUiCulture = CultureInfo.CurrentUICulture;
        var targetCulture = managedCulture ?? CultureInfo.InvariantCulture;

        try
        {
            // Apply managed cultures BEFORE Win32 SetThreadLocale: on .NET Framework (clr.dll),
            // assigning CurrentCulture internally calls SetThreadLocale(targetCulture.LCID),
            // which would otherwise overwrite threadLocaleId (e.g. resetting 0x0409 to 0x007F).
            CultureInfo.CurrentCulture = targetCulture;
            CultureInfo.CurrentUICulture = targetCulture;

            if (!NativeMethods.SetThreadLocale(threadLocaleId))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
        }
        catch
        {
            CultureInfo.CurrentCulture = previousCulture;
            CultureInfo.CurrentUICulture = previousUiCulture;
            _ = NativeMethods.SetThreadLocale(previousLcid);
            throw;
        }

        return new ComAutomationLocaleScope(
            ownerThreadId,
            previousLcid,
            previousCulture,
            previousUiCulture);
    }

    /// <summary>
    /// Enters a thread-affine locale scope configured with <see cref="InvariantLocaleId"/> (<c>0x007F</c>)
    /// and <see cref="CultureInfo.InvariantCulture"/>.
    /// </summary>
    /// <returns>A disposable <see cref="ComAutomationLocaleScope"/> that restores the prior Win32 <c>LCID</c> and managed cultures on <see cref="Dispose"/>.</returns>
    public static ComAutomationLocaleScope EnterInvariant()
        => Enter(InvariantLocaleId, CultureInfo.InvariantCulture);

    /// <summary>
    /// Restores the original Win32 thread locale and managed cultures on <see cref="OwnerThreadId"/>.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when disposed from a thread other than <see cref="OwnerThreadId"/>.</exception>
    public void Dispose()
    {
        var currentThreadId = Environment.CurrentManagedThreadId;
        if (currentThreadId != OwnerThreadId)
        {
#pragma warning disable CA1065, S3877 // Intentional fail-closed thread-affinity guard so wrong-thread disposal never mutates another thread's LCID.
            throw new InvalidOperationException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "ComAutomationLocaleScope must be disposed on its owner thread (expected thread {0}, actual thread {1}).",
                    OwnerThreadId,
                    currentThreadId));
#pragma warning restore CA1065, S3877
        }

        if (_disposed)
        {
            return;
        }

        _disposed = true;
        try
        {
            CultureInfo.CurrentCulture = _previousCulture;
            CultureInfo.CurrentUICulture = _previousUiCulture;
        }
        finally
        {
            _ = NativeMethods.SetThreadLocale(PreviousThreadLocaleId);
        }
    }
}
