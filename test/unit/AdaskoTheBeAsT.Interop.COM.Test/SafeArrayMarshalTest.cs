using System.Runtime.InteropServices;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable MA0051 // Comprehensive SafeArray marshaling scenarios.
public class SafeArrayMarshalTest
{
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(-5)]
    public void CopyFromBorrowedSafeArrayShouldReadUntypedDelphiSafeArrayWithoutMutatingOrDestroyingIt(int lowerBound)
    {
        var input = new[] { 10, -20, 30, 42 };
        Marshal.ThrowExceptionForHR(
            ProbeNativeInterop.SafeNativeMethods.CreateDelphiBorrowedInt32SafeArray(input, (uint)input.Length, lowerBound, out var psa));

        var copied = SafeArrayMarshal.CopyFromBorrowedSafeArray<int>(psa);
        copied.Should().Equal(input);

        // Borrowed SAFEARRAY must retain fFeatures == 0 and cLocks == 0 (paired Access/Unaccess, not destroyed).
        Marshal.ThrowExceptionForHR(
            ProbeNativeInterop.SafeNativeMethods.InspectAndFreeDelphiSafeArray(psa, out var features, out var locks));
        features.Should().Be(0);
        locks.Should().Be(0);
    }

    [Fact]
    public void CopyFromBorrowedSafeArrayShouldSupportEmptyDelphiSafeArray()
    {
        Marshal.ThrowExceptionForHR(
            ProbeNativeInterop.SafeNativeMethods.CreateDelphiBorrowedInt32SafeArray(values: null, 0, 1, out var psa));

        var copied = SafeArrayMarshal.CopyFromBorrowedSafeArray<int>(psa);
        copied.Should().BeEmpty();

        Marshal.ThrowExceptionForHR(
            ProbeNativeInterop.SafeNativeMethods.InspectAndFreeDelphiSafeArray(psa, out var features, out var locks));
        features.Should().Be(0);
        locks.Should().Be(0);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(-3)]
    public void ConsumeOwnedSafeArrayShouldCopyValuesAndDestroyNativeArray(int lowerBound)
    {
        var expected = new[] { 1.5, -2.25, 3.75 };
        Marshal.ThrowExceptionForHR(
            ProbeNativeInterop.SafeNativeMethods.CreateOwnedDoubleSafeArray(expected, (uint)expected.Length, lowerBound, out var psa));

        var copied = SafeArrayMarshal.ConsumeOwnedSafeArray<double>(psa);
        copied.Should().Equal(expected);
    }

    [Fact]
    public void BorrowedAndOwnedBstrArraysShouldPreserveBorrowedAndDestroyOwned()
    {
        var s0 = Marshal.StringToHGlobalUni("Alpha");
        var s2 = Marshal.StringToHGlobalUni("Omega");
        try
        {
            var ptrs = new[] { s0, IntPtr.Zero, s2 };
            Marshal.ThrowExceptionForHR(
                ProbeNativeInterop.SafeNativeMethods.CreateBstrSafeArray(ptrs, (uint)ptrs.Length, lowerBound: 1, out var psa));

            // First read as borrowed: native SAFEARRAY remains intact for subsequent consumption!
            var borrowedRead = SafeArrayMarshal.CopyBorrowedBstrArray(psa);
            borrowedRead.Should().Equal("Alpha", null, "Omega");

            // Second read as owned: reads and destroys the native SAFEARRAY + BSTRs.
            var ownedRead = SafeArrayMarshal.ConsumeOwnedBstrArray(psa);
            ownedRead.Should().Equal("Alpha", null, "Omega");
        }
        finally
        {
            Marshal.FreeHGlobal(s2);
            Marshal.FreeHGlobal(s0);
        }
    }

    [Fact]
    public void BorrowedAndOwnedVariantArraysShouldConvertMixedVariantElementsAcrossNegativeLowerBounds()
    {
        Marshal.ThrowExceptionForHR(
            ProbeNativeInterop.SafeNativeMethods.CreateMixedVariantSafeArray(42, 6.5, "DelphiVB6", lowerBound: -2, out var psa));

        var borrowed = SafeArrayMarshal.CopyBorrowedVariantArray(psa);
        borrowed.Should().HaveCount(4);
        borrowed.Should().HaveElementAt(0, 42);
        borrowed.Should().HaveElementAt(1, 6.5);
        borrowed.Should().HaveElementAt(2, "DelphiVB6");
        borrowed[3].Should().BeNull();

        var owned = SafeArrayMarshal.ConsumeOwnedVariantArray(psa);
        owned.Should().HaveElementAt(0, 42);
        owned.Should().HaveElementAt(1, 6.5);
        owned.Should().HaveElementAt(2, "DelphiVB6");
        owned[3].Should().BeNull();
    }

    [Fact]
    public void NormalizeVariantArrayAndCreateOneBasedArrayShouldHandleZeroBasedOneBasedNegativeAndVariantArrays()
    {
        SafeArrayMarshal.NormalizeVariantArray<double>(comArray: null).Should().BeEmpty();

        var szArray = new[] { 1.0, 2.0, 3.0 };
        SafeArrayMarshal.NormalizeVariantArray<double>(szArray).Should().BeSameAs(szArray);

        IReadOnlyList<float> items = new[] { 10.5f, 20.5f, 30.5f };
        var oneBased = SafeArrayMarshal.CreateOneBasedArray(items);
        oneBased.GetLowerBound(0).Should().Be(1);
        oneBased.GetUpperBound(0).Should().Be(3);
        SafeArrayMarshal.NormalizeVariantArray<float>(oneBased).Should().Equal(10.5f, 20.5f, 30.5f);

#if NET8_0_OR_GREATER
        ReadOnlySpan<int> spanItems = [7, 8, 9];
        var spanOneBased = SafeArrayMarshal.CreateOneBasedArray(spanItems);
        spanOneBased.GetLowerBound(0).Should().Be(1);
        SafeArrayMarshal.NormalizeVariantArray<int>(spanOneBased).Should().Equal(7, 8, 9);
#endif

        var negativeArray = Array.CreateInstance(typeof(int), [3], [-2]);
        negativeArray.SetValue(100, -2);
        negativeArray.SetValue(200, -1);
        negativeArray.SetValue(300, 0);
        SafeArrayMarshal.NormalizeVariantArray<int>(negativeArray).Should().Equal(100, 200, 300);

        var variantObjects = Array.CreateInstance(typeof(object), [3], [1]);
        variantObjects.SetValue((short)5, 1);
        variantObjects.SetValue(12, 2);
        variantObjects.SetValue(19.0, 3);
        SafeArrayMarshal.NormalizeVariantArray<double>(variantObjects).Should().Equal(5.0, 12.0, 19.0);

        var multiDim = new int[2, 2];
        Action badRank = () => SafeArrayMarshal.NormalizeVariantArray<int>(multiDim);
        badRank.Should().Throw<RankException>();
    }

    [Fact]
    public void SafeArrayFailurePathsShouldRejectNullWrongRankElementSizeMismatchAndReferenceFlags()
    {
        Action nullBorrowed = () => SafeArrayMarshal.CopyFromBorrowedSafeArray<int>(IntPtr.Zero);
        nullBorrowed.Should().Throw<ArgumentNullException>();

        Action nullBstr = () => SafeArrayMarshal.CopyBorrowedBstrArray(IntPtr.Zero);
        nullBstr.Should().Throw<ArgumentNullException>();

        Action nullVariant = () => SafeArrayMarshal.CopyBorrowedVariantArray(IntPtr.Zero);
        nullVariant.Should().Throw<ArgumentNullException>();

        // Rank 2 SAFEARRAY rejected cleanly without leaking locks.
        Marshal.ThrowExceptionForHR(ProbeNativeInterop.SafeNativeMethods.CreateRank2Int32SafeArray(out var rank2));
        try
        {
            Action readRank2 = () => SafeArrayMarshal.CopyFromBorrowedSafeArray<int>(rank2);
            readRank2.Should().Throw<RankException>();
        }
        finally
        {
            Marshal.ThrowExceptionForHR(ProbeNativeInterop.SafeNativeMethods.DestroyNativeSafeArray(rank2));
        }

        // Element size mismatch: reading 4-byte Int32 array as 8-byte double.
        Marshal.ThrowExceptionForHR(
            ProbeNativeInterop.SafeNativeMethods.CreateDelphiBorrowedInt32SafeArray([1, 2], 2, 0, out var intPsa));
        try
        {
            Action wrongSize = () => SafeArrayMarshal.CopyFromBorrowedSafeArray<double>(intPsa);
            wrongSize.Should().Throw<SafeArrayTypeMismatchException>();
        }
        finally
        {
            Marshal.ThrowExceptionForHR(ProbeNativeInterop.SafeNativeMethods.InspectAndFreeDelphiSafeArray(intPsa, out _, out var locks));
            locks.Should().Be(0);
        }

        // Reference-flag mismatch: reading FADF_BSTR array through unmanaged CopyFromBorrowedSafeArray<IntPtr>.
        var strPtr = Marshal.StringToHGlobalUni("SampleText");
        try
        {
            Marshal.ThrowExceptionForHR(ProbeNativeInterop.SafeNativeMethods.CreateBstrSafeArray([strPtr], 1, 0, out var bstrPsa));
            try
            {
                Action unmanagedOnBstr = () => SafeArrayMarshal.CopyFromBorrowedSafeArray<IntPtr>(bstrPsa);
                unmanagedOnBstr.Should().Throw<SafeArrayTypeMismatchException>();
            }
            finally
            {
                Marshal.ThrowExceptionForHR(ProbeNativeInterop.SafeNativeMethods.DestroyNativeSafeArray(bstrPsa));
            }
        }
        finally
        {
            Marshal.FreeHGlobal(strPtr);
        }
    }
}
#pragma warning restore MA0051
#pragma warning restore SA1121
