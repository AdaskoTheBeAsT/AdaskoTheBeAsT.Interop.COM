// dllmain.h : Declaration of module class.

class CNativeCOMModule : public ATL::CAtlDllModuleT< CNativeCOMModule >
{
public :
	DECLARE_LIBID(LIBID_NativeCOMLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_NATIVECOM, "{4d04c8d8-6773-491b-be1a-08f83268186d}")
};

extern class CNativeCOMModule _AtlModule;
