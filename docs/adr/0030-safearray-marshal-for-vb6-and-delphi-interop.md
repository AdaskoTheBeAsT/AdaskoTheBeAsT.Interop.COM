# ADR-0030: Ownership-Aware `SafeArrayMarshal` Utilities for VB6 1-Based and Delphi Untyped `SAFEARRAY*` Payloads

- Status: Accepted
- Date: 2026-11-15
- Deciders: @AdaskoTheBeAsT
- Release: v5.0.0

## Context

Interoperating with legacy Visual Basic 6 and Borland Delphi 6/7 COM calculation servers on .NET triggers two recurring `SAFEARRAY` marshaling failures:

1. **Delphi borrowed/untyped `SAFEARRAY*` pointers:** Several Delphi COM methods (`Variants.DynArrayToVariant` or raw `out IntPtr` `SAFEARRAY*` parameters) omit `FADF_HAVEVARTYPE` (`0x0080`) or leave caller-owned lifetime on the returned pointer. The CLR's built-in `[MarshalAs(UnmanagedType.SafeArray)]` marshaler rejects untyped descriptors with `SafeArrayTypeMismatchException` (`0x80131533`) or destroys borrowed caller memory if marshaled as owned.
2. **VB6 non-zero lower-bound arrays (`1..N` or negative lower bounds):** VB6 `Option Base 1` and `Dim values(1 To N)` produce non-SZArray `System.Array` instances (`double[*]`, `string[*]`, `object[*]`) when marshaled via `VARIANT` (`object`), causing `(double[])result` casts in C# to throw `InvalidCastException`. Similarly, `BSTR` and `VARIANT` arrays require element-wise conversion rather than raw `Buffer.MemoryCopy`.

## Decision

Provide a dedicated public `SafeArrayMarshal` utility class with explicit ownership and bounds validation:

1. **Separate borrowed vs. owned unmanaged `SAFEARRAY*` helpers:**
   - `CopyFromBorrowedSafeArray<T>(IntPtr)` never destroys or mutates the caller's `SAFEARRAY`.
   - `ConsumeOwnedSafeArray<T>(IntPtr)` copies the payload and deterministically invokes `oleaut32!SafeArrayDestroy` in `finally` (paired `CopyFromSafeArrayPointer<T>(IntPtr, bool destroyAfterRead)` also available).
   - Validate rank (`SafeArrayGetDim == 1`), element size (`SafeArrayGetElemsize == sizeof(T)`), incompatible feature flags (`FADF_BSTR`, `FADF_UNKNOWN`, `FADF_DISPATCH`, `FADF_VARIANT` rejected on the unmanaged numeric path), `SafeArrayGetVartype` compatibility when `FADF_HAVEVARTYPE` is set, and checked 64-bit `upperBound - lowerBound + 1` element-count arithmetic; always pair `SafeArrayAccessData` with `SafeArrayUnaccessData` in `finally`.
2. **Ownership-aware `BSTR` and `VARIANT` `SAFEARRAY*` helpers:**
   - Provide `CopyBorrowedBstrArray` / `ConsumeOwnedBstrArray` (`Marshal.PtrToStringBSTR` per slot) and `CopyBorrowedVariantArray` / `ConsumeOwnedVariantArray` (`Marshal.GetObjectsForNativeVariants` over contiguous 16-byte/24-byte `VARIANT` strides).
3. **Managed VB6 array normalization and creation:**
   - Provide `NormalizeVariantArray<T>(object?)` (zero-copy fast path for `T[]`, `Array.Copy` normalization for 1D non-zero lower-bound `T[*]`, and `Convert.ChangeType` / invariant conversion for `object[]` / `object[*]`) and `CreateOneBasedArray<T>` (`ReadOnlySpan<T>` on `.NET 8+` and `IReadOnlyList<T>` on all TFMs) returning a `1`-based `System.Array` (`1..N`, or `[0]` with lower bound `1` when empty).

## Consequences

- Positive: Consumers can safely marshal `0`-based, `1`-based, negative-lower-bound, borrowed, and owned `SAFEARRAY` payloads across `byte`, `short`, `int`, `long`, `float`, `double`, `decimal`, `bool`, `BSTR`, and `VARIANT` without `SafeArrayTypeMismatchException`, `InvalidCastException`, double-frees, or unmanaged leaks.
- Positive: Zero external NuGet dependencies across all supported TFMs.
- Negative: Multi-dimensional (`rank > 1`) raw `SAFEARRAY*` pointers intentionally throw `RankException`; callers with matrix `SAFEARRAY` instances continue to use standard multi-dimensional `System.Array` indexing.

## Alternatives Considered

- Exposing a single `CopySafeArray<T>(IntPtr)` method that always destroys the input array: rejected because calling `SafeArrayDestroy` on a Delphi borrowed/field-backed `SAFEARRAY*` corrupts the native heap.
- Copying `BSTR` or `VARIANT` arrays through the `unmanaged` byte-copy path: rejected because `BSTR` and `VARIANT` elements hold indirect native heap allocations requiring COM-aware conversion.

## References

- `src/AdaskoTheBeAsT.Interop.COM/SafeArrayMarshal.cs`
- `src/AdaskoTheBeAsT.Interop.COM/NativeMethods.cs`
- `test/unit/AdaskoTheBeAsT.Interop.COM.Test/SafeArrayMarshalTest.cs`
