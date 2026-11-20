#include <Windows.h>
#include <Unknwn.h>
#include <objidl.h>
#include <ole2.h>
#include <oleauto.h>
#include <comsvcs.h>
#include <new>
#include <cstring>

extern "C" IMAGE_DOS_HEADER __ImageBase;

struct __declspec(uuid("a80d257f-4121-49f2-84ca-0b4b40487d94")) IReleaseCallback : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE OnReleased(DWORD threadId) = 0;
};

struct __declspec(uuid("c64bcaf7-3ee8-421d-a8ee-46cb0c9238f4")) ILifetimeProbe : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Ping() = 0;
    virtual HRESULT STDMETHODCALLTYPE SetReleaseCallback(IReleaseCallback* callback) = 0;
};

// Coclass CLSIDs and custom interfaces for COM-01..COM-05 multi-coclass and module-isolation tests:
// CLSID_ProbeAlpha       = {3a930e11-62b4-41d5-8c0a-510d92a70101}
// IID_IProbeAlpha        = {3a930e11-62b4-41d5-8c0a-510d92a70102}
// IID_IProbeSecondary    = {3a930e11-62b4-41d5-8c0a-510d92a70103}
// CLSID_ProbeBeta        = {3a930e11-62b4-41d5-8c0a-510d92a70201}
// IID_IProbeBeta         = {3a930e11-62b4-41d5-8c0a-510d92a70202}
// CLSID_ProbeProxyStub   = {3a930e11-62b4-41d5-8c0a-510d92a70301}
// CLSID_ProbeUnavailable = {3a930e11-62b4-41d5-8c0a-510d92a70401}

static const GUID CLSID_ProbeAlpha =
    { 0x3a930e11, 0x62b4, 0x41d5, { 0x8c, 0x0a, 0x51, 0x0d, 0x92, 0xa7, 0x01, 0x01 } };

static const GUID CLSID_ProbeBeta =
    { 0x3a930e11, 0x62b4, 0x41d5, { 0x8c, 0x0a, 0x51, 0x0d, 0x92, 0xa7, 0x02, 0x01 } };

static const GUID CLSID_ProbeProxyStub =
    { 0x3a930e11, 0x62b4, 0x41d5, { 0x8c, 0x0a, 0x51, 0x0d, 0x92, 0xa7, 0x03, 0x01 } };

static const GUID CLSID_ProbeUnavailable =
    { 0x3a930e11, 0x62b4, 0x41d5, { 0x8c, 0x0a, 0x51, 0x0d, 0x92, 0xa7, 0x04, 0x01 } };

static const GUID CLSID_PersistProbe =
    { 0x3a930e11, 0x62b4, 0x41d5, { 0x8c, 0x0a, 0x51, 0x0d, 0x92, 0xa7, 0x05, 0x01 } };

struct __declspec(uuid("3a930e11-62b4-41d5-8c0a-510d92a70102")) IProbeAlpha : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE ComputeAlpha(LONG input, LONG* output) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetModuleBaseAddress(void** moduleBase) = 0;
};

struct __declspec(uuid("3a930e11-62b4-41d5-8c0a-510d92a70103")) IProbeSecondary : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetSignature(LONG* signature) = 0;
};

struct __declspec(uuid("3a930e11-62b4-41d5-8c0a-510d92a70202")) IProbeBeta : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE CombineWithAlpha(IProbeAlpha* alpha, LONG seed, LONG* output) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetBetaTag(LONG* tag) = 0;
};

static volatile LONG destroyedCount = 0;
static volatile LONG lastReleaseThreadId = 0;
static volatile LONG activeComObjectCount = 0;
static volatile LONG serverLockCount = 0;
static volatile LONG forceBusyUnload = 0;
static volatile LONG moduleAlphaCounter = 0;

class LifetimeProbe final : public ILifetimeProbe
{
    volatile LONG references = 1;
    IReleaseCallback* releaseCallback = nullptr;

    ~LifetimeProbe()
    {
        const DWORD threadId = GetCurrentThreadId();
        InterlockedExchange(&lastReleaseThreadId, static_cast<LONG>(threadId));
        InterlockedIncrement(&destroyedCount);
        if (releaseCallback != nullptr)
        {
            releaseCallback->OnReleased(threadId);
            releaseCallback->Release();
        }
    }

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override
    {
        if (result == nullptr)
            return E_POINTER;

        *result = nullptr;
        if (iid != IID_IUnknown && iid != __uuidof(ILifetimeProbe))
            return E_NOINTERFACE;

        *result = static_cast<ILifetimeProbe*>(this);
        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return InterlockedIncrement(&references);
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const LONG remaining = InterlockedDecrement(&references);
        if (remaining == 0)
            delete this;
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE Ping() override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetReleaseCallback(IReleaseCallback* callback) override
    {
        if (callback != nullptr)
            callback->AddRef();
        if (releaseCallback != nullptr)
            releaseCallback->Release();
        releaseCallback = callback;
        return S_OK;
    }
};

class ProbeAlphaObject final : public IProbeAlpha, public IProbeSecondary
{
    volatile LONG references = 1;

    ~ProbeAlphaObject()
    {
        InterlockedDecrement(&activeComObjectCount);
    }

public:
    ProbeAlphaObject()
    {
        InterlockedIncrement(&activeComObjectCount);
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override
    {
        if (result == nullptr)
            return E_POINTER;

        *result = nullptr;
        if (iid == IID_IUnknown || iid == __uuidof(IProbeAlpha))
        {
            *result = static_cast<IProbeAlpha*>(this);
            AddRef();
            return S_OK;
        }

        if (iid == __uuidof(IProbeSecondary))
        {
            *result = static_cast<IProbeSecondary*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return InterlockedIncrement(&references);
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const LONG remaining = InterlockedDecrement(&references);
        if (remaining == 0)
            delete this;
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE ComputeAlpha(LONG input, LONG* output) override
    {
        if (output == nullptr)
            return E_POINTER;

        const LONG seq = InterlockedIncrement(&moduleAlphaCounter);
        *output = (input * 10) + seq;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetModuleBaseAddress(void** moduleBase) override
    {
        if (moduleBase == nullptr)
            return E_POINTER;

        *moduleBase = reinterpret_cast<void*>(&__ImageBase);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetSignature(LONG* signature) override
    {
        if (signature == nullptr)
            return E_POINTER;

        *signature = 0xA1FA;
        return S_OK;
    }
};

class ProbeBetaObject final : public IProbeBeta
{
    volatile LONG references = 1;

    ~ProbeBetaObject()
    {
        InterlockedDecrement(&activeComObjectCount);
    }

public:
    ProbeBetaObject()
    {
        InterlockedIncrement(&activeComObjectCount);
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override
    {
        if (result == nullptr)
            return E_POINTER;

        *result = nullptr;
        if (iid == IID_IUnknown || iid == __uuidof(IProbeBeta))
        {
            *result = static_cast<IProbeBeta*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return InterlockedIncrement(&references);
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const LONG remaining = InterlockedDecrement(&references);
        if (remaining == 0)
            delete this;
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE CombineWithAlpha(IProbeAlpha* alpha, LONG seed, LONG* output) override
    {
        if (alpha == nullptr || output == nullptr)
            return E_POINTER;

        LONG alphaValue = 0;
        const HRESULT hr = alpha->ComputeAlpha(seed, &alphaValue);
        if (FAILED(hr))
            return hr;

        *output = alphaValue + 7;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetBetaTag(LONG* tag) override
    {
        if (tag == nullptr)
            return E_POINTER;

        *tag = 0xBE7A;
        return S_OK;
    }
};

enum class ProbeCoclassKind
{
    Alpha = 1,
    Beta = 2,
    Persist = 3,
};

class ProbeClassFactory final : public IClassFactory
{
    volatile LONG references = 1;
    ProbeCoclassKind kind;

    ~ProbeClassFactory() = default;

public:
    explicit ProbeClassFactory(ProbeCoclassKind coclassKind)
        : kind(coclassKind)
    {
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override
    {
        if (result == nullptr)
            return E_POINTER;

        *result = nullptr;
        if (iid == IID_IUnknown || iid == IID_IClassFactory)
        {
            *result = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return InterlockedIncrement(&references);
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const LONG remaining = InterlockedDecrement(&references);
        if (remaining == 0)
            delete this;
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer, REFIID iid, void** result) override;

    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override
    {
        if (lock)
            InterlockedIncrement(&serverLockCount);
        else
            InterlockedDecrement(&serverLockCount);
        return S_OK;
    }
};

// Native IPersistStream + IRunnableObject + IProbeAlpha object for NanoServerOle2Shim tests.
class PersistRunnableProbe final : public IPersistStream, public IRunnableObject, public IProbeAlpha
{
    volatile LONG references = 1;
    LONG payloadValue;
    BOOL running = FALSE;

    ~PersistRunnableProbe()
    {
        InterlockedDecrement(&activeComObjectCount);
    }

public:
    explicit PersistRunnableProbe(LONG initialValue)
        : payloadValue(initialValue)
    {
        InterlockedIncrement(&activeComObjectCount);
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override
    {
        if (result == nullptr)
            return E_POINTER;

        *result = nullptr;
        if (iid == IID_IUnknown || iid == IID_IPersist || iid == IID_IPersistStream)
        {
            *result = static_cast<IPersistStream*>(this);
            AddRef();
            return S_OK;
        }

        if (iid == IID_IRunnableObject)
        {
            *result = static_cast<IRunnableObject*>(this);
            AddRef();
            return S_OK;
        }

        if (iid == __uuidof(IProbeAlpha))
        {
            *result = static_cast<IProbeAlpha*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return InterlockedIncrement(&references);
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const LONG remaining = InterlockedDecrement(&references);
        if (remaining == 0)
            delete this;
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE GetClassID(CLSID* pClassID) override
    {
        if (pClassID == nullptr)
            return E_POINTER;

        *pClassID = CLSID_PersistProbe;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE IsDirty() override
    {
        return S_FALSE;
    }

    HRESULT STDMETHODCALLTYPE Load(IStream* pStm) override
    {
        if (pStm == nullptr)
            return E_POINTER;

        ULONG cbRead = 0;
        LONG value = 0;
        const HRESULT hr = pStm->Read(&value, sizeof(LONG), &cbRead);
        if (FAILED(hr))
            return hr;
        if (cbRead != sizeof(LONG))
            return STG_E_READFAULT;

        payloadValue = value;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Save(IStream* pStm, BOOL) override
    {
        if (pStm == nullptr)
            return E_POINTER;

        ULONG cbWritten = 0;
        const HRESULT hr = pStm->Write(&payloadValue, sizeof(LONG), &cbWritten);
        if (FAILED(hr))
            return hr;
        if (cbWritten != sizeof(LONG))
            return STG_E_WRITEFAULT;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER* pcbSize) override
    {
        if (pcbSize == nullptr)
            return E_POINTER;

        pcbSize->QuadPart = sizeof(LONG);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetRunningClass(LPCLSID lpClsid) override
    {
        if (lpClsid == nullptr)
            return E_POINTER;

        *lpClsid = CLSID_ProbeAlpha;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Run(LPBINDCTX) override
    {
        running = TRUE;
        payloadValue += 1000;
        return S_OK;
    }

    BOOL STDMETHODCALLTYPE IsRunning() override
    {
        return running;
    }

    HRESULT STDMETHODCALLTYPE LockRunning(BOOL, BOOL) override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetContainedObject(BOOL) override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ComputeAlpha(LONG input, LONG* output) override
    {
        if (output == nullptr)
            return E_POINTER;

        *output = payloadValue + input;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetModuleBaseAddress(void** moduleBase) override
    {
        if (moduleBase == nullptr)
            return E_POINTER;

        *moduleBase = reinterpret_cast<void*>(&__ImageBase);
        return S_OK;
    }
};

HRESULT STDMETHODCALLTYPE ProbeClassFactory::CreateInstance(IUnknown* outer, REFIID iid, void** result)
{
    if (result == nullptr)
        return E_POINTER;

    *result = nullptr;
    if (outer != nullptr)
        return CLASS_E_NOAGGREGATION;

    if (kind == ProbeCoclassKind::Alpha)
    {
        auto* obj = new (std::nothrow) ProbeAlphaObject();
        if (obj == nullptr)
            return E_OUTOFMEMORY;

        const HRESULT hr = obj->QueryInterface(iid, result);
        obj->Release();
        return hr;
    }

    if (kind == ProbeCoclassKind::Persist)
    {
        auto* persist = new (std::nothrow) PersistRunnableProbe(0);
        if (persist == nullptr)
            return E_OUTOFMEMORY;

        const HRESULT hr = persist->QueryInterface(iid, result);
        persist->Release();
        return hr;
    }

    auto* beta = new (std::nothrow) ProbeBetaObject();
    if (beta == nullptr)
        return E_OUTOFMEMORY;

    const HRESULT hr = beta->QueryInterface(iid, result);
    beta->Release();
    return hr;
}

class NativeMemoryStream final : public IStream
{
    volatile LONG references = 1;
    BYTE buffer[256] {};
    ULONG length = 0;
    ULONG position = 0;

    ~NativeMemoryStream() = default;

public:
    void Rewind()
    {
        position = 0;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override
    {
        if (result == nullptr)
            return E_POINTER;

        *result = nullptr;
        if (iid == IID_IUnknown || iid == IID_ISequentialStream || iid == IID_IStream)
        {
            *result = static_cast<IStream*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return InterlockedIncrement(&references);
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const LONG remaining = InterlockedDecrement(&references);
        if (remaining == 0)
            delete this;
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE Read(void* pv, ULONG cb, ULONG* pcbRead) override
    {
        if (pv == nullptr)
            return STG_E_INVALIDPOINTER;

        const ULONG available = position < length ? (length - position) : 0;
        const ULONG toCopy = cb < available ? cb : available;
        if (toCopy > 0)
        {
            std::memcpy(pv, buffer + position, toCopy);
            position += toCopy;
        }

        if (pcbRead != nullptr)
            *pcbRead = toCopy;
        return toCopy == cb ? S_OK : S_FALSE;
    }

    HRESULT STDMETHODCALLTYPE Write(const void* pv, ULONG cb, ULONG* pcbWritten) override
    {
        if (pv == nullptr)
            return STG_E_INVALIDPOINTER;
        if (position + cb > sizeof(buffer))
            return STG_E_MEDIUMFULL;

        if (cb > 0)
        {
            std::memcpy(buffer + position, pv, cb);
            position += cb;
            if (position > length)
                length = position;
        }

        if (pcbWritten != nullptr)
            *pcbWritten = cb;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition) override
    {
        LONGLONG target = 0;
        if (dwOrigin == STREAM_SEEK_SET)
            target = dlibMove.QuadPart;
        else if (dwOrigin == STREAM_SEEK_CUR)
            target = static_cast<LONGLONG>(position) + dlibMove.QuadPart;
        else if (dwOrigin == STREAM_SEEK_END)
            target = static_cast<LONGLONG>(length) + dlibMove.QuadPart;
        else
            return STG_E_INVALIDFUNCTION;

        if (target < 0 || target > static_cast<LONGLONG>(sizeof(buffer)))
            return STG_E_INVALIDFUNCTION;

        position = static_cast<ULONG>(target);
        if (plibNewPosition != nullptr)
            plibNewPosition->QuadPart = position;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetSize(ULARGE_INTEGER) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE CopyTo(IStream*, ULARGE_INTEGER, ULARGE_INTEGER*, ULARGE_INTEGER*) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE Commit(DWORD) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE Revert() override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE LockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE UnlockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE Stat(STATSTG*, DWORD) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE Clone(IStream**) override { return E_NOTIMPL; }
};

class NativeTestDispenserDriver final : public IDispenserDriver
{
    volatile LONG references = 1;

public:
    LONG createdCount = 0;
    LONG destroyedCountLocal = 0;
    LONG destroyedStringCount = 0;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override
    {
        if (result == nullptr)
            return E_POINTER;

        *result = nullptr;
        if (iid == IID_IUnknown || iid == IID_IDispenserDriver)
        {
            *result = static_cast<IDispenserDriver*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return InterlockedIncrement(&references);
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const LONG remaining = InterlockedDecrement(&references);
        if (remaining == 0)
            delete this;
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE CreateResource(const RESTYPID resTypId, RESID* pResId, TIMEINSECS* pSecsFreeBeforeDestroy) override
    {
        if (pResId == nullptr || pSecsFreeBeforeDestroy == nullptr)
            return E_POINTER;

        const LONG index = InterlockedIncrement(&createdCount);
        *pResId = static_cast<RESID>(resTypId + static_cast<RESTYPID>(index * 100));
        *pSecsFreeBeforeDestroy = 60;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE RateResource(const RESTYPID, const RESID, const BOOL, RESOURCERATING* pRating) override
    {
        if (pRating == nullptr)
            return E_POINTER;

        *pRating = 100;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE EnlistResource(const RESID, const TRANSID) override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ResetResource(const RESID) override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DestroyResource(const RESID) override
    {
        InterlockedIncrement(&destroyedCountLocal);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DestroyResourceS(constSRESID) override
    {
        InterlockedIncrement(&destroyedStringCount);
        return S_OK;
    }
};

struct RawDelphiSafeArrayAllocation
{
    SAFEARRAY descriptor;
    LONG payload[1];
};

extern "C" HRESULT __stdcall CreateLifetimeProbe(ILifetimeProbe** result)
{
    if (result == nullptr)
        return E_POINTER;
    *result = new (std::nothrow) LifetimeProbe();
    return *result == nullptr ? E_OUTOFMEMORY : S_OK;
}

extern "C" LONG __stdcall GetDestroyedCount()
{
    return InterlockedCompareExchange(&destroyedCount, 0, 0);
}

extern "C" DWORD __stdcall GetLastReleaseThreadId()
{
    return static_cast<DWORD>(InterlockedCompareExchange(&lastReleaseThreadId, 0, 0));
}

extern "C" LONG __stdcall GetActiveComObjectCount()
{
    return InterlockedCompareExchange(&activeComObjectCount, 0, 0);
}

extern "C" void __stdcall SetForceBusyUnload(BOOL busy)
{
    InterlockedExchange(&forceBusyUnload, busy ? 1 : 0);
}

extern "C" void __stdcall ResetProbeModuleState()
{
    InterlockedExchange(&moduleAlphaCounter, 0);
    InterlockedExchange(&forceBusyUnload, 0);
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    if (ppv == nullptr)
        return E_POINTER;

    *ppv = nullptr;
    if (rclsid == CLSID_ProbeProxyStub || rclsid == CLSID_ProbeUnavailable)
        return CLASS_E_CLASSNOTAVAILABLE;

    ProbeCoclassKind kind;
    if (rclsid == CLSID_ProbeAlpha)
        kind = ProbeCoclassKind::Alpha;
    else if (rclsid == CLSID_ProbeBeta)
        kind = ProbeCoclassKind::Beta;
    else if (rclsid == CLSID_PersistProbe)
        kind = ProbeCoclassKind::Persist;
    else
        return CLASS_E_CLASSNOTAVAILABLE;

    auto* factory = new (std::nothrow) ProbeClassFactory(kind);
    if (factory == nullptr)
        return E_OUTOFMEMORY;

    const HRESULT hr = factory->QueryInterface(riid, ppv);
    factory->Release();
    return hr;
}

extern "C" HRESULT __stdcall DllCanUnloadNow()
{
    if (InterlockedCompareExchange(&forceBusyUnload, 0, 0) != 0)
        return S_FALSE;

    const LONG objects = InterlockedCompareExchange(&activeComObjectCount, 0, 0);
    const LONG locks = InterlockedCompareExchange(&serverLockCount, 0, 0);
    return (objects == 0 && locks == 0) ? S_OK : S_FALSE;
}

extern "C" HRESULT __stdcall CreateDelphiBorrowedInt32SafeArray(
    const LONG* values,
    ULONG count,
    LONG lowerBound,
    SAFEARRAY** result)
{
    if (result == nullptr)
        return E_POINTER;

    *result = nullptr;
    const size_t dataBytes = count == 0 ? sizeof(LONG) : static_cast<size_t>(count) * sizeof(LONG);
    auto* rawBytes = static_cast<BYTE*>(CoTaskMemAlloc(sizeof(SAFEARRAY) + dataBytes));
    if (rawBytes == nullptr)
        return E_OUTOFMEMORY;

    std::memset(rawBytes, 0, sizeof(SAFEARRAY) + dataBytes);
    auto* psa = reinterpret_cast<SAFEARRAY*>(rawBytes);
    auto* dataPtr = reinterpret_cast<LONG*>(rawBytes + sizeof(SAFEARRAY));
    if (count > 0 && values != nullptr)
        std::memcpy(dataPtr, values, static_cast<size_t>(count) * sizeof(LONG));

    // Deliberately omit FADF_HAVEVARTYPE (0x0080) to replicate Borland Delphi untyped borrowed SAFEARRAY*.
    psa->cDims = 1;
    psa->fFeatures = 0;
    psa->cbElements = sizeof(LONG);
    psa->cLocks = 0;
    psa->pvData = dataPtr;
    psa->rgsabound[0].cElements = count;
    psa->rgsabound[0].lLbound = lowerBound;

    *result = psa;
    return S_OK;
}

extern "C" HRESULT __stdcall InspectAndFreeDelphiSafeArray(
    SAFEARRAY* psa,
    USHORT* outFeatures,
    ULONG* outLocks)
{
    if (psa == nullptr || outFeatures == nullptr || outLocks == nullptr)
        return E_POINTER;

    *outFeatures = psa->fFeatures;
    *outLocks = psa->cLocks;
    CoTaskMemFree(psa);
    return S_OK;
}

extern "C" HRESULT __stdcall CreateOwnedDoubleSafeArray(
    const double* values,
    ULONG count,
    LONG lowerBound,
    SAFEARRAY** result)
{
    if (result == nullptr)
        return E_POINTER;

    *result = nullptr;
    SAFEARRAYBOUND bound {};
    bound.cElements = count;
    bound.lLbound = lowerBound;

    SAFEARRAY* psa = SafeArrayCreate(VT_R8, 1, &bound);
    if (psa == nullptr)
        return E_OUTOFMEMORY;

    if (count > 0 && values != nullptr)
    {
        void* data = nullptr;
        const HRESULT hr = SafeArrayAccessData(psa, &data);
        if (FAILED(hr))
        {
            SafeArrayDestroy(psa);
            return hr;
        }

        std::memcpy(data, values, static_cast<size_t>(count) * sizeof(double));
        SafeArrayUnaccessData(psa);
    }

    *result = psa;
    return S_OK;
}

extern "C" HRESULT __stdcall CreateBstrSafeArray(
    const LPCWSTR* values,
    ULONG count,
    LONG lowerBound,
    SAFEARRAY** result)
{
    if (result == nullptr)
        return E_POINTER;

    *result = nullptr;
    SAFEARRAYBOUND bound {};
    bound.cElements = count;
    bound.lLbound = lowerBound;

    SAFEARRAY* psa = SafeArrayCreate(VT_BSTR, 1, &bound);
    if (psa == nullptr)
        return E_OUTOFMEMORY;

    for (ULONG i = 0; i < count; ++i)
    {
        LONG index = lowerBound + static_cast<LONG>(i);
        BSTR item = (values != nullptr && values[i] != nullptr) ? SysAllocString(values[i]) : nullptr;
        if (values != nullptr && values[i] != nullptr && item == nullptr)
        {
            SafeArrayDestroy(psa);
            return E_OUTOFMEMORY;
        }

        const HRESULT hr = SafeArrayPutElement(psa, &index, item);
        SysFreeString(item);
        if (FAILED(hr))
        {
            SafeArrayDestroy(psa);
            return hr;
        }
    }

    *result = psa;
    return S_OK;
}

extern "C" HRESULT __stdcall CreateMixedVariantSafeArray(
    LONG intValue,
    double doubleValue,
    LPCWSTR stringValue,
    LONG lowerBound,
    SAFEARRAY** result)
{
    if (result == nullptr)
        return E_POINTER;

    *result = nullptr;
    SAFEARRAYBOUND bound {};
    bound.cElements = 4;
    bound.lLbound = lowerBound;

    SAFEARRAY* psa = SafeArrayCreate(VT_VARIANT, 1, &bound);
    if (psa == nullptr)
        return E_OUTOFMEMORY;

    VARIANT v0;
    VariantInit(&v0);
    v0.vt = VT_I4;
    v0.lVal = intValue;
    LONG idx0 = lowerBound;
    HRESULT hr = SafeArrayPutElement(psa, &idx0, &v0);
    if (FAILED(hr))
    {
        SafeArrayDestroy(psa);
        return hr;
    }

    VARIANT v1;
    VariantInit(&v1);
    v1.vt = VT_R8;
    v1.dblVal = doubleValue;
    LONG idx1 = lowerBound + 1;
    hr = SafeArrayPutElement(psa, &idx1, &v1);
    if (FAILED(hr))
    {
        SafeArrayDestroy(psa);
        return hr;
    }

    VARIANT v2;
    VariantInit(&v2);
    v2.vt = VT_BSTR;
    v2.bstrVal = stringValue != nullptr ? SysAllocString(stringValue) : nullptr;
    LONG idx2 = lowerBound + 2;
    hr = SafeArrayPutElement(psa, &idx2, &v2);
    VariantClear(&v2);
    if (FAILED(hr))
    {
        SafeArrayDestroy(psa);
        return hr;
    }

    VARIANT v3;
    VariantInit(&v3);
    LONG idx3 = lowerBound + 3;
    hr = SafeArrayPutElement(psa, &idx3, &v3);
    if (FAILED(hr))
    {
        SafeArrayDestroy(psa);
        return hr;
    }

    *result = psa;
    return S_OK;
}

extern "C" HRESULT __stdcall CreateRank2Int32SafeArray(SAFEARRAY** result)
{
    if (result == nullptr)
        return E_POINTER;

    SAFEARRAYBOUND bounds[2] {};
    bounds[0].cElements = 2;
    bounds[0].lLbound = 0;
    bounds[1].cElements = 2;
    bounds[1].lLbound = 0;
    *result = SafeArrayCreate(VT_I4, 2, bounds);
    return *result == nullptr ? E_OUTOFMEMORY : S_OK;
}

extern "C" HRESULT __stdcall DestroyNativeSafeArray(SAFEARRAY* psa)
{
    if (psa == nullptr)
        return S_OK;
    return SafeArrayDestroy(psa);
}

extern "C" HRESULT __stdcall CreateNativePersistRunnableProbe(LONG initialValue, IUnknown** result)
{
    if (result == nullptr)
        return E_POINTER;

    *result = nullptr;
    auto* obj = new (std::nothrow) PersistRunnableProbe(initialValue);
    if (obj == nullptr)
        return E_OUTOFMEMORY;

    const HRESULT hr = obj->QueryInterface(IID_IUnknown, reinterpret_cast<void**>(result));
    obj->Release();
    return hr;
}

extern "C" HRESULT __stdcall CreateNativeMemoryStream(IStream** result)
{
    if (result == nullptr)
        return E_POINTER;

    *result = new (std::nothrow) NativeMemoryStream();
    return *result == nullptr ? E_OUTOFMEMORY : S_OK;
}

extern "C" HRESULT __stdcall RewindNativeMemoryStream(IStream* stream)
{
    if (stream == nullptr)
        return E_POINTER;

    LARGE_INTEGER zero {};
    return stream->Seek(zero, STREAM_SEEK_SET, nullptr);
}

extern "C" HRESULT __stdcall ExerciseNativeDispenserFactory(
    IClassFactory* classFactory,
    LONG* outCreatedCount,
    LONG* outDestroyedCount,
    LONG* outDestroyedStringCount)
{
    if (classFactory == nullptr || outCreatedCount == nullptr || outDestroyedCount == nullptr || outDestroyedStringCount == nullptr)
        return E_POINTER;

    *outCreatedCount = 0;
    *outDestroyedCount = 0;
    *outDestroyedStringCount = 0;

    IDispenserManager* manager = nullptr;
    HRESULT hr = classFactory->CreateInstance(nullptr, IID_IDispenserManager, reinterpret_cast<void**>(&manager));
    if (FAILED(hr) || manager == nullptr)
        return FAILED(hr) ? hr : E_FAIL;

    INSTID instId = 99;
    TRANSID transId = 99;
    hr = manager->GetContext(&instId, &transId);
    if (FAILED(hr) || instId != 0 || transId != 0)
    {
        manager->Release();
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    auto* driver = new (std::nothrow) NativeTestDispenserDriver();
    if (driver == nullptr)
    {
        manager->Release();
        return E_OUTOFMEMORY;
    }

    IHolder* holder = nullptr;
    hr = manager->RegisterDispenser(driver, L"NativeProbeDispenser", &holder);
    if (FAILED(hr) || holder == nullptr)
    {
        driver->Release();
        manager->Release();
        return FAILED(hr) ? hr : E_FAIL;
    }

    RESID res1 = 0;
    hr = holder->AllocResource(static_cast<RESTYPID>(7), &res1);
    if (FAILED(hr) || res1 != static_cast<RESID>(107))
    {
        holder->Close();
        holder->Release();
        driver->Release();
        manager->Release();
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    hr = holder->TrackResource(res1);
    if (FAILED(hr))
    {
        holder->Close();
        holder->Release();
        driver->Release();
        manager->Release();
        return hr;
    }

    // Untrack with fDestroy == FALSE must NOT invoke DestroyResource.
    hr = holder->UntrackResource(res1, FALSE);
    if (FAILED(hr) || driver->destroyedCountLocal != 0)
    {
        holder->Close();
        holder->Release();
        driver->Release();
        manager->Release();
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    // Untrack with fDestroy == TRUE MUST invoke DestroyResource.
    hr = holder->UntrackResource(res1, TRUE);
    if (FAILED(hr) || driver->destroyedCountLocal != 1)
    {
        holder->Close();
        holder->Release();
        driver->Release();
        manager->Release();
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    // FreeResource MUST invoke DestroyResource.
    RESID res2 = 0;
    hr = holder->AllocResource(static_cast<RESTYPID>(9), &res2);
    if (FAILED(hr) || res2 != static_cast<RESID>(209))
    {
        holder->Close();
        holder->Release();
        driver->Release();
        manager->Release();
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    hr = holder->FreeResource(res2);
    if (FAILED(hr) || driver->destroyedCountLocal != 2)
    {
        holder->Close();
        holder->Release();
        driver->Release();
        manager->Release();
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    // UntrackResourceS with fDestroy == FALSE then TRUE.
    hr = holder->TrackResourceS(L"ConnA");
    if (FAILED(hr))
    {
        holder->Close();
        holder->Release();
        driver->Release();
        manager->Release();
        return hr;
    }

    hr = holder->UntrackResourceS(L"ConnA", FALSE);
    if (FAILED(hr) || driver->destroyedStringCount != 0)
    {
        holder->Close();
        holder->Release();
        driver->Release();
        manager->Release();
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    hr = holder->UntrackResourceS(L"ConnA", TRUE);
    if (FAILED(hr) || driver->destroyedStringCount != 1)
    {
        holder->Close();
        holder->Release();
        driver->Release();
        manager->Release();
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    hr = holder->Close();
    if (FAILED(hr))
    {
        holder->Release();
        driver->Release();
        manager->Release();
        return hr;
    }

    *outCreatedCount = driver->createdCount;
    *outDestroyedCount = driver->destroyedCountLocal;
    *outDestroyedStringCount = driver->destroyedStringCount;

    holder->Release();
    driver->Release();
    manager->Release();
    return S_OK;
}

extern "C" HRESULT __stdcall InvokeImportedOleRun(IUnknown* unknown)
{
    return ::OleRun(unknown);
}

extern "C" HRESULT __stdcall PopulateProbeComErrorInfo(
    const GUID* interfaceId,
    LPCWSTR source,
    LPCWSTR description,
    LPCWSTR helpFile,
    DWORD helpContext,
    HRESULT returnHr)
{
    ICreateErrorInfo* createErrorInfo = nullptr;
    HRESULT hr = ::CreateErrorInfo(&createErrorInfo);
    if (FAILED(hr) || createErrorInfo == nullptr)
        return FAILED(hr) ? hr : E_FAIL;

    if (interfaceId != nullptr)
        createErrorInfo->SetGUID(*interfaceId);

    if (source != nullptr)
    {
        BSTR bstrSource = ::SysAllocString(source);
        if (bstrSource != nullptr)
        {
            createErrorInfo->SetSource(bstrSource);
            ::SysFreeString(bstrSource);
        }
    }

    if (description != nullptr)
    {
        BSTR bstrDescription = ::SysAllocString(description);
        if (bstrDescription != nullptr)
        {
            createErrorInfo->SetDescription(bstrDescription);
            ::SysFreeString(bstrDescription);
        }
    }

    if (helpFile != nullptr)
    {
        BSTR bstrHelpFile = ::SysAllocString(helpFile);
        if (bstrHelpFile != nullptr)
        {
            createErrorInfo->SetHelpFile(bstrHelpFile);
            ::SysFreeString(bstrHelpFile);
        }
    }

    createErrorInfo->SetHelpContext(helpContext);

    IErrorInfo* errorInfo = nullptr;
    hr = createErrorInfo->QueryInterface(IID_IErrorInfo, reinterpret_cast<void**>(&errorInfo));
    createErrorInfo->Release();
    if (FAILED(hr) || errorInfo == nullptr)
        return FAILED(hr) ? hr : E_FAIL;

    hr = ::SetErrorInfo(0, errorInfo);
    errorInfo->Release();
    if (FAILED(hr))
        return hr;

    return returnHr;
}

extern "C" HRESULT __stdcall ParseDoubleWithThreadLocale(
    LPCWSTR input,
    DWORD dwFlags,
    double* outValue,
    LCID* outObservedLcid)
{
    if (outValue == nullptr)
        return E_POINTER;

    *outValue = 0.0;
    const LCID lcid = ::GetThreadLocale();
    if (outObservedLcid != nullptr)
        *outObservedLcid = lcid;

    return ::VarR8FromStr(input, lcid, dwFlags, outValue);
}

extern "C" HRESULT __stdcall FormatDoubleWithThreadLocale(
    double input,
    DWORD dwFlags,
    BSTR* outBstr,
    LCID* outObservedLcid)
{
    if (outBstr == nullptr)
        return E_POINTER;

    *outBstr = nullptr;
    const LCID lcid = ::GetThreadLocale();
    if (outObservedLcid != nullptr)
        *outObservedLcid = lcid;

    return ::VarBstrFromR8(input, lcid, dwFlags, outBstr);
}

extern "C" void __stdcall InvokeImportedCoUninitializeTwice()
{
    ::CoUninitialize();
    ::CoUninitialize();
}

extern "C" void __stdcall InvokeImportedOleUninitializeTwice()
{
    ::OleUninitialize();
    ::OleUninitialize();
}
