// StringConcatenator.h : Declaration of the CStringConcatenator

#pragma once
#include "resource.h"       // main symbols



#include "NativeCOM_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;


// CStringConcatenator

class ATL_NO_VTABLE CStringConcatenator :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CStringConcatenator, &CLSID_StringConcatenator>,
	public IDispatchImpl<IStringConcatenator, &IID_IStringConcatenator, &LIBID_NativeCOMLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CStringConcatenator()
	{
	}

DECLARE_REGISTRY_RESOURCEID(106)


BEGIN_COM_MAP(CStringConcatenator)
	COM_INTERFACE_ENTRY(IStringConcatenator)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:
    // Implement the IStringConcatenator interface
    STDMETHOD(ConcatStrings)(BSTR string1, BSTR string2, BSTR* result);
};

OBJECT_ENTRY_AUTO(__uuidof(StringConcatenator), CStringConcatenator)
