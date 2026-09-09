#include <Windows.h>
#include <Unknwn.h>
#include <new>

struct __declspec(uuid("a80d257f-4121-49f2-84ca-0b4b40487d94")) IReleaseCallback : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE OnReleased(DWORD threadId) = 0;
};

struct __declspec(uuid("c64bcaf7-3ee8-421d-a8ee-46cb0c9238f4")) ILifetimeProbe : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Ping() = 0;
    virtual HRESULT STDMETHODCALLTYPE SetReleaseCallback(IReleaseCallback* callback) = 0;
};

static volatile LONG destroyedCount = 0;
static volatile LONG lastReleaseThreadId = 0;

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
