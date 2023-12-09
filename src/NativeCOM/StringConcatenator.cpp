// StringConcatenator.cpp : Implementation of CStringConcatenator

#include "pch.h"
#include "StringConcatenator.h"


// CStringConcatenator

STDMETHODIMP CStringConcatenator::ConcatStrings(BSTR string1, BSTR string2, BSTR* result)
{
    if (string1 == nullptr || string2 == nullptr || result == nullptr)
        return E_INVALIDARG;

    // Calculate the length of the two strings
    const UINT len1 = SysStringLen(string1);
    const UINT len2 = SysStringLen(string2);
    const UINT totalLen = len1 + len2;

    // Allocate a new BSTR for the result
    *result = SysAllocStringLen(nullptr, totalLen);
    if (*result == nullptr)
        return E_OUTOFMEMORY;

    // Copy the first string into result
    wcscpy_s(*result, len1 + 1, string1);

    // Concatenate the second string onto result
    wcscat_s(*result, totalLen + 1, string2);

    return S_OK;
}