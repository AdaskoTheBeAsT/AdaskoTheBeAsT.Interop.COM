using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1121 // Use built-in type alias
/// <summary>
/// Safe marshaling utilities for 1-based Visual Basic 6 <c>SAFEARRAY</c> values, negative-lower-bound arrays,
/// <c>BSTR</c>/<c>VARIANT</c> <c>SAFEARRAY</c> pointers, and untyped or borrowed Borland Delphi <c>SAFEARRAY*</c> descriptors.
/// </summary>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public static class SafeArrayMarshal
{
    private const ushort FadfRecord = 0x0020;
    private const ushort FadfHaveVartype = 0x0080;
    private const ushort FadfBstr = 0x0100;
    private const ushort FadfUnknown = 0x0200;
    private const ushort FadfDispatch = 0x0400;
    private const ushort FadfVariant = 0x0800;
    private const ushort FadfReferenceMask = FadfRecord | FadfBstr | FadfUnknown | FadfDispatch | FadfVariant;

    private const ushort VtEmpty = 0;
    private const ushort VtI2 = 2;
    private const ushort VtI4 = 3;
    private const ushort VtR4 = 4;
    private const ushort VtR8 = 5;
    private const ushort VtCy = 6;
    private const ushort VtDate = 7;
    private const ushort VtBstr = 8;
    private const ushort VtDispatch = 9;
    private const ushort VtError = 10;
    private const ushort VtVariant = 12;
    private const ushort VtUnknown = 13;
    private const ushort VtDecimal = 14;
    private const ushort VtI1 = 16;
    private const ushort VtUi1 = 17;
    private const ushort VtUi2 = 18;
    private const ushort VtUi4 = 19;
    private const ushort VtI8 = 20;
    private const ushort VtUi8 = 21;
    private const ushort VtInt = 22;
    private const ushort VtUint = 23;
    private const ushort VtRecord = 36;

    // Matches Array.MaxLength on .NET without exceeding int.MaxValue on .NET Framework.
    private const int MaxSafeArrayElements = 0x7FEFFFFF;

    /// <summary>
    /// Copies a single-dimensional unmanaged <c>SAFEARRAY*</c> without destroying or modifying the borrowed native descriptor.
    /// Supports untyped Delphi <c>SAFEARRAY</c> descriptors that omit <c>FADF_HAVEVARTYPE</c>.
    /// </summary>
    /// <typeparam name="T">The unmanaged blittable element type.</typeparam>
    /// <param name="safeArrayPtr">Non-zero pointer to a borrowed 1D <c>SAFEARRAY</c>.</param>
    /// <returns>A zero-based managed copy of the array elements.</returns>
    public static T[] CopyFromBorrowedSafeArray<T>(IntPtr safeArrayPtr)
        where T : unmanaged
        => CopyUnmanagedCore<T>(safeArrayPtr, destroyAfterRead: false);

    /// <summary>
    /// Copies a single-dimensional unmanaged <c>SAFEARRAY*</c> and destroys the caller-owned native <c>SAFEARRAY</c>
    /// after unlocking its data buffer.
    /// </summary>
    /// <typeparam name="T">The unmanaged blittable element type.</typeparam>
    /// <param name="safeArrayPtr">Non-zero pointer to an owned 1D <c>SAFEARRAY</c>.</param>
    /// <returns>A zero-based managed copy of the array elements.</returns>
    public static T[] ConsumeOwnedSafeArray<T>(IntPtr safeArrayPtr)
        where T : unmanaged
        => CopyUnmanagedCore<T>(safeArrayPtr, destroyAfterRead: true);

    /// <summary>
    /// Copies a single-dimensional unmanaged <c>SAFEARRAY*</c>, optionally destroying the native descriptor when
    /// <paramref name="destroyAfterRead"/> is <see langword="true"/>.
    /// Borrowed arrays (<paramref name="destroyAfterRead"/> = <see langword="false"/>) are never destroyed or modified.
    /// </summary>
    /// <typeparam name="T">The unmanaged blittable element type.</typeparam>
    /// <param name="safeArrayPtr">Non-zero pointer to a 1D <c>SAFEARRAY</c>.</param>
    /// <param name="destroyAfterRead">Whether the caller transfers ownership for <c>SafeArrayDestroy</c>.</param>
    /// <returns>A zero-based managed copy of the array elements.</returns>
    public static T[] CopyFromSafeArrayPointer<T>(IntPtr safeArrayPtr, bool destroyAfterRead = false)
        where T : unmanaged
        => CopyUnmanagedCore<T>(safeArrayPtr, destroyAfterRead);

    /// <summary>
    /// Copies a borrowed single-dimensional <c>BSTR</c> <c>SAFEARRAY*</c> into a managed <see cref="string"/> array
    /// without freeing the borrowed native <c>BSTR</c> elements or destroying the descriptor.
    /// </summary>
    /// <param name="safeArrayPtr">Non-zero pointer to a borrowed 1D <c>SAFEARRAY</c> of <c>BSTR</c>.</param>
    /// <returns>A zero-based managed array of strings.</returns>
    public static string?[] CopyBorrowedBstrArray(IntPtr safeArrayPtr)
        => CopyBstrCore(safeArrayPtr, destroyAfterRead: false);

    /// <summary>
    /// Copies an owned single-dimensional <c>BSTR</c> <c>SAFEARRAY*</c> into a managed <see cref="string"/> array
    /// and destroys the native <c>SAFEARRAY</c> (and its owned <c>BSTR</c> elements) afterwards.
    /// </summary>
    /// <param name="safeArrayPtr">Non-zero pointer to an owned 1D <c>SAFEARRAY</c> of <c>BSTR</c>.</param>
    /// <returns>A zero-based managed array of strings.</returns>
    public static string?[] ConsumeOwnedBstrArray(IntPtr safeArrayPtr)
        => CopyBstrCore(safeArrayPtr, destroyAfterRead: true);

    /// <summary>
    /// Copies a borrowed single-dimensional <c>VARIANT</c> <c>SAFEARRAY*</c> into a managed <see cref="object"/> array
    /// without clearing the borrowed <c>VARIANT</c> elements or destroying the descriptor.
    /// </summary>
    /// <param name="safeArrayPtr">Non-zero pointer to a borrowed 1D <c>SAFEARRAY</c> of <c>VARIANT</c>.</param>
    /// <returns>A zero-based managed array of converted variant values.</returns>
    public static object?[] CopyBorrowedVariantArray(IntPtr safeArrayPtr)
        => CopyVariantCore(safeArrayPtr, destroyAfterRead: false);

    /// <summary>
    /// Copies an owned single-dimensional <c>VARIANT</c> <c>SAFEARRAY*</c> into a managed <see cref="object"/> array
    /// and destroys the native <c>SAFEARRAY</c> (clearing its owned <c>VARIANT</c> elements) afterwards.
    /// </summary>
    /// <param name="safeArrayPtr">Non-zero pointer to an owned 1D <c>SAFEARRAY</c> of <c>VARIANT</c>.</param>
    /// <returns>A zero-based managed array of converted variant values.</returns>
    public static object?[] ConsumeOwnedVariantArray(IntPtr safeArrayPtr)
        => CopyVariantCore(safeArrayPtr, destroyAfterRead: true);

    /// <summary>
    /// Normalizes a COM-returned <c>VARIANT</c> array (zero-based <c>T[]</c>, 1-based VB6 <c>Array</c>,
    /// negative-lower-bound 1D <c>Array</c>, or <c>object[]</c> of boxed <c>VARIANT</c> values) into a zero-based <c>T[]</c>.
    /// </summary>
    /// <typeparam name="T">Target managed element type.</typeparam>
    /// <param name="comArray">The COM array value, or <see langword="null"/>/<see cref="DBNull"/>.</param>
    /// <returns>A zero-based <c>T[]</c> containing the normalized elements.</returns>
    public static T[] NormalizeVariantArray<T>(object? comArray)
    {
        if (comArray is null or DBNull)
        {
            return [];
        }

        if (comArray is T[] fastArray)
        {
            return fastArray;
        }

        if (comArray is not Array array)
        {
            throw new ArgumentException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Expected a single-dimensional COM Array compatible with {0}[], but received {1}.",
                    typeof(T).Name,
                    comArray.GetType().FullName),
                nameof(comArray));
        }

        if (array.Rank != 1)
        {
            throw new RankException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Only single-dimensional COM arrays are supported, but received rank {0}.",
                    array.Rank));
        }

        var length = array.Length;
        if (length == 0)
        {
            return [];
        }

        var lowerBound = array.GetLowerBound(0);
        var result = new T[length];
        if (array.GetType().GetElementType() == typeof(T))
        {
            Array.Copy(array, lowerBound, result, 0, length);
            return result;
        }

        for (var i = 0; i < length; i++)
        {
            var rawValue = array.GetValue(checked(lowerBound + i));
            result[i] = ConvertVariantElement<T>(rawValue);
        }

        return result;
    }

    /// <summary>
    /// Creates a single-dimensional <see cref="Array"/> with lower bound <c>1</c> (<c>1..N</c>) suitable for
    /// Visual Basic 6 <c>Option Base 1</c> and <c>Dim arr(1 To N)</c> COM parameters.
    /// </summary>
    /// <typeparam name="T">Element type.</typeparam>
    /// <param name="items">Items to copy into the 1-based array.</param>
    /// <returns>A 1-based single-dimensional <see cref="Array"/>.</returns>
    public static Array CreateOneBasedArray<T>(IReadOnlyList<T> items)
    {
        ThrowHelper.ThrowIfNull(items, nameof(items));

        var count = items.Count;
        var array = Array.CreateInstance(typeof(T), [count], [1]);
        for (var i = 0; i < count; i++)
        {
            array.SetValue(items[i], i + 1);
        }

        return array;
    }

#if NET8_0_OR_GREATER
    /// <summary>
    /// Creates a single-dimensional <see cref="Array"/> with lower bound <c>1</c> (<c>1..N</c>) from <paramref name="items"/>.
    /// </summary>
    /// <typeparam name="T">Element type.</typeparam>
    /// <param name="items">Span of items to copy into the 1-based array.</param>
    /// <returns>A 1-based single-dimensional <see cref="Array"/>.</returns>
    public static Array CreateOneBasedArray<T>(ReadOnlySpan<T> items)
    {
        var count = items.Length;
        var array = Array.CreateInstance(typeof(T), [count], [1]);
        for (var i = 0; i < count; i++)
        {
            array.SetValue(items[i], i + 1);
        }

        return array;
    }
#endif

#pragma warning disable S6640 // Blittable element copy requires unsafe sizeof(T) and Buffer.MemoryCopy.
    private static unsafe T[] CopyUnmanagedCore<T>(IntPtr safeArrayPtr, bool destroyAfterRead)
        where T : unmanaged
    {
        if (safeArrayPtr == IntPtr.Zero)
        {
            throw new ArgumentNullException(nameof(safeArrayPtr));
        }

        try
        {
            var count = ValidateDescriptorAndCount(safeArrayPtr, (uint)sizeof(T), out var features);
            ValidateUnmanagedTypeCompatibility<T>(safeArrayPtr, features);

            var byteCount = checked((long)count * sizeof(T));
            var accessHr = NativeMethods.SafeArrayAccessData(safeArrayPtr, out var dataPtr);
            Marshal.ThrowExceptionForHR(accessHr);
            try
            {
                if (count == 0)
                {
                    return [];
                }

                if (dataPtr == IntPtr.Zero)
                {
                    throw new InvalidOperationException("SafeArrayAccessData returned a null data pointer for a non-empty SAFEARRAY.");
                }

                var result = new T[count];
                fixed (T* destination = result)
                {
                    Buffer.MemoryCopy((void*)dataPtr, destination, byteCount, byteCount);
                }

                return result;
            }
            finally
            {
                var unaccessHr = NativeMethods.SafeArrayUnaccessData(safeArrayPtr);
                Marshal.ThrowExceptionForHR(unaccessHr);
            }
        }
        finally
        {
            if (destroyAfterRead)
            {
                var destroyHr = NativeMethods.SafeArrayDestroy(safeArrayPtr);
                Marshal.ThrowExceptionForHR(destroyHr);
            }
        }
    }
#pragma warning restore S6640

    private static string?[] CopyBstrCore(IntPtr safeArrayPtr, bool destroyAfterRead)
    {
        if (safeArrayPtr == IntPtr.Zero)
        {
            throw new ArgumentNullException(nameof(safeArrayPtr));
        }

        try
        {
            var count = ValidateDescriptorAndCount(safeArrayPtr, (uint)IntPtr.Size, out var features);
            ValidateBstrCompatibility(safeArrayPtr, features);

            var accessHr = NativeMethods.SafeArrayAccessData(safeArrayPtr, out var dataPtr);
            Marshal.ThrowExceptionForHR(accessHr);
            try
            {
                if (count == 0)
                {
                    return [];
                }

                if (dataPtr == IntPtr.Zero)
                {
                    throw new InvalidOperationException("SafeArrayAccessData returned a null data pointer for a non-empty BSTR SAFEARRAY.");
                }

                var result = new string?[count];
                for (var i = 0; i < count; i++)
                {
                    var bstrPtr = Marshal.ReadIntPtr(dataPtr, checked(i * IntPtr.Size));
                    result[i] = bstrPtr == IntPtr.Zero ? null : Marshal.PtrToStringBSTR(bstrPtr);
                }

                return result;
            }
            finally
            {
                var unaccessHr = NativeMethods.SafeArrayUnaccessData(safeArrayPtr);
                Marshal.ThrowExceptionForHR(unaccessHr);
            }
        }
        finally
        {
            if (destroyAfterRead)
            {
                var destroyHr = NativeMethods.SafeArrayDestroy(safeArrayPtr);
                Marshal.ThrowExceptionForHR(destroyHr);
            }
        }
    }

    private static object?[] CopyVariantCore(IntPtr safeArrayPtr, bool destroyAfterRead)
    {
        if (safeArrayPtr == IntPtr.Zero)
        {
            throw new ArgumentNullException(nameof(safeArrayPtr));
        }

        var variantSize = IntPtr.Size == 8 ? 24 : 16;
        try
        {
            var count = ValidateDescriptorAndCount(safeArrayPtr, (uint)variantSize, out var features);
            ValidateVariantCompatibility(safeArrayPtr, features);

            var accessHr = NativeMethods.SafeArrayAccessData(safeArrayPtr, out var dataPtr);
            Marshal.ThrowExceptionForHR(accessHr);
            try
            {
                if (count == 0)
                {
                    return [];
                }

                if (dataPtr == IntPtr.Zero)
                {
                    throw new InvalidOperationException("SafeArrayAccessData returned a null data pointer for a non-empty VARIANT SAFEARRAY.");
                }

                var result = new object?[count];
                var baseAddress = dataPtr.ToInt64();
                for (var i = 0; i < count; i++)
                {
                    var elemPtr = new IntPtr(checked(baseAddress + ((long)i * variantSize)));
                    result[i] = Marshal.GetObjectForNativeVariant(elemPtr);
                }

                return result;
            }
            finally
            {
                var unaccessHr = NativeMethods.SafeArrayUnaccessData(safeArrayPtr);
                Marshal.ThrowExceptionForHR(unaccessHr);
            }
        }
        finally
        {
            if (destroyAfterRead)
            {
                var destroyHr = NativeMethods.SafeArrayDestroy(safeArrayPtr);
                Marshal.ThrowExceptionForHR(destroyHr);
            }
        }
    }

    private static int ValidateDescriptorAndCount(IntPtr safeArrayPtr, uint expectedElemSize, out ushort features)
    {
        var rank = NativeMethods.SafeArrayGetDim(safeArrayPtr);
        if (rank != 1)
        {
            throw new RankException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Expected a single-dimensional SAFEARRAY (rank 1), but received rank {0}.",
                    rank));
        }

        features = unchecked((ushort)Marshal.ReadInt16(safeArrayPtr, 2));
        var elemSize = NativeMethods.SafeArrayGetElemsize(safeArrayPtr);
        if (elemSize != expectedElemSize)
        {
            throw new SafeArrayTypeMismatchException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "SAFEARRAY element size mismatch: expected {0} bytes, but descriptor specifies {1} bytes.",
                    expectedElemSize,
                    elemSize));
        }

        var lboundHr = NativeMethods.SafeArrayGetLBound(safeArrayPtr, 1, out var lowerBound);
        Marshal.ThrowExceptionForHR(lboundHr);

        var uboundHr = NativeMethods.SafeArrayGetUBound(safeArrayPtr, 1, out var upperBound);
        Marshal.ThrowExceptionForHR(uboundHr);

        var countLong = (long)upperBound - lowerBound + 1L;
        if (countLong == 0)
        {
            return 0;
        }

        if (countLong < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(safeArrayPtr),
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Invalid SAFEARRAY bounds [{0}..{1}].",
                    lowerBound,
                    upperBound));
        }

        if (countLong > MaxSafeArrayElements)
        {
            throw new OverflowException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "SAFEARRAY element count ({0}) exceeds maximum managed array length ({1}).",
                    countLong,
                    MaxSafeArrayElements));
        }

        return (int)countLong;
    }

    private static void ValidateUnmanagedTypeCompatibility<T>(IntPtr safeArrayPtr, ushort features)
        where T : unmanaged
    {
        if ((features & FadfReferenceMask) != 0)
        {
            throw new SafeArrayTypeMismatchException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Cannot copy reference/variant SAFEARRAY (fFeatures=0x{0:X4}) as unmanaged {1}[]; use CopyBorrowedBstrArray or CopyBorrowedVariantArray.",
                    features,
                    typeof(T).Name));
        }

        if ((features & FadfHaveVartype) == 0)
        {
            // Untyped Delphi SAFEARRAY descriptors omit FADF_HAVEVARTYPE; element size was already verified.
            return;
        }

        if (NativeMethods.SafeArrayGetVartype(safeArrayPtr, out var vt) < 0 || vt == VtEmpty)
        {
            return;
        }

        if (vt is VtBstr or VtVariant or VtUnknown or VtDispatch or VtRecord)
        {
            throw new SafeArrayTypeMismatchException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "SAFEARRAY VARTYPE {0} requires ownership-aware conversion and cannot be copied as unmanaged {1}[].",
                    vt,
                    typeof(T).Name));
        }

        if (!IsCompatibleUnmanagedVartype(typeof(T), vt))
        {
            throw new SafeArrayTypeMismatchException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "SAFEARRAY VARTYPE {0} does not match target unmanaged element type {1}.",
                    vt,
                    typeof(T).Name));
        }
    }

    private static bool IsCompatibleUnmanagedVartype(Type elementType, ushort vt)
    {
        if (elementType == typeof(sbyte))
        {
            return vt == VtI1;
        }

        if (elementType == typeof(byte))
        {
            return vt == VtUi1;
        }

        if (elementType == typeof(short))
        {
            return vt == VtI2;
        }

        if (elementType == typeof(ushort))
        {
            return vt == VtUi2;
        }

        if (elementType == typeof(int))
        {
            return vt is VtI4 or VtInt or VtError;
        }

        if (elementType == typeof(uint))
        {
            return vt is VtUi4 or VtUint;
        }

        if (elementType == typeof(long))
        {
            return vt is VtI8 or VtCy;
        }

        if (elementType == typeof(ulong))
        {
            return vt == VtUi8;
        }

        if (elementType == typeof(float))
        {
            return vt == VtR4;
        }

        if (elementType == typeof(double))
        {
            return vt is VtR8 or VtDate;
        }

        return elementType != typeof(decimal) || vt == VtDecimal;
    }

    private static void ValidateBstrCompatibility(IntPtr safeArrayPtr, ushort features)
    {
        if ((features & (FadfVariant | FadfUnknown | FadfDispatch | FadfRecord)) != 0)
        {
            throw new SafeArrayTypeMismatchException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "SAFEARRAY feature flags 0x{0:X4} are incompatible with BSTR array conversion.",
                    features));
        }

        if ((features & FadfHaveVartype) != 0
            && NativeMethods.SafeArrayGetVartype(safeArrayPtr, out var vt) >= 0
            && vt != VtEmpty
            && vt != VtBstr)
        {
            throw new SafeArrayTypeMismatchException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Expected VT_BSTR (8) SAFEARRAY, but descriptor reports VARTYPE {0}.",
                    vt));
        }
    }

    private static void ValidateVariantCompatibility(IntPtr safeArrayPtr, ushort features)
    {
        if ((features & (FadfBstr | FadfUnknown | FadfDispatch | FadfRecord)) != 0)
        {
            throw new SafeArrayTypeMismatchException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "SAFEARRAY feature flags 0x{0:X4} are incompatible with VARIANT array conversion.",
                    features));
        }

        if ((features & FadfHaveVartype) != 0
            && NativeMethods.SafeArrayGetVartype(safeArrayPtr, out var vt) >= 0
            && vt != VtEmpty
            && vt != VtVariant)
        {
            throw new SafeArrayTypeMismatchException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Expected VT_VARIANT (12) SAFEARRAY, but descriptor reports VARTYPE {0}.",
                    vt));
        }
    }

    private static T ConvertVariantElement<T>(object? rawValue)
    {
        if (rawValue is T direct)
        {
            return direct;
        }

        if (rawValue is null or DBNull)
        {
            if (default(T) is null)
            {
                return default!;
            }

            throw new InvalidCastException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Cannot convert null or DBNull VARIANT element to non-nullable {0}.",
                    typeof(T).Name));
        }

        return (T)Convert.ChangeType(rawValue, Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T), CultureInfo.InvariantCulture);
    }
}
#pragma warning restore SA1121
