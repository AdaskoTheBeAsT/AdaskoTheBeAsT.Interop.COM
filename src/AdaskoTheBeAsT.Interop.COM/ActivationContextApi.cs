using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif
using System.Threading;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable MA0158 // Use System.Threading.Lock
#pragma warning disable SA1121 // Use built-in type alias
// An operation-scoped native boundary permits deterministic failure tests without global hooks
// and provides the manifest-driven in-process class-factory fallback on Windows Nano Server.
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
internal class ActivationContextApi
{
#pragma warning disable MA0176 // Optimize guid creation
    private static readonly Guid IidIClassFactory = new("00000001-0000-0000-C000-000000000046");
    private static readonly Guid IidIUnknown = new("00000000-0000-0000-C000-000000000046");
#pragma warning restore MA0176

#if NET10_0_OR_GREATER
    private static readonly Lock ModuleLock = new();
#endif
#if NET472_OR_GREATER
    private static readonly object ModuleLock = new();
#endif
    private static readonly Dictionary<string, LoadedModuleEntry> LoadedModules = new(StringComparer.OrdinalIgnoreCase);

    [ThreadStatic]
    private static PendingLeaseContext? _pendingLeaseContext;

    [ThreadStatic]
    private static PendingCreateContext? _pendingCreateContext;

    private static long _nextFallbackToken = -1;

    private readonly ConcurrentDictionary<IntPtr, FallbackContext> _fallbackContexts = new();
    private readonly ConcurrentDictionary<IntPtr, FallbackActivation> _fallbackActivations = new();
    private readonly ConcurrentDictionary<uint, List<FallbackActivation>> _threadFallbackStacks = new();
    private readonly ConcurrentDictionary<uint, int> _lastFallbackErrors = new();
    private readonly Dictionary<ApartmentRegistrationKey, SharedClassRegistrationLease> _apartmentRegistrations = [];
#if NET10_0_OR_GREATER
    private readonly Lock _registrationLock = new();
#endif
#if NET472_OR_GREATER
    private readonly object _registrationLock = new();
#endif

    internal static ActivationContextApi Instance { get; } = new();

    internal static bool IsFallbackHandle(IntPtr handle)
        => handle != (IntPtr)(-1) && handle.ToInt64() < 0;

    internal static int CollectDeferredUnloads(ActivationContextApi? contextApi = null)
    {
        var api = contextApi ?? Instance;
        api.FreeUnusedLibrariesCore();
        var candidates = new List<LoadedModuleEntry>();
        lock (ModuleLock)
        {
            foreach (var entry in LoadedModules.Values)
            {
                if (entry.ActiveLeaseCount == 0 && !entry.RetainPermanently && entry.PendingDeferredUnload)
                {
                    candidates.Add(entry);
                }
            }

            var unloaded = 0;
            for (var i = 0; i < candidates.Count; i++)
            {
                if (TryUnloadModuleEntryUnderLock(candidates[i], api))
                {
                    unloaded++;
                }
            }

            return unloaded;
        }
    }

    internal static IReadOnlyList<IntPtr> GetLoadedHandlesForPaths(IReadOnlyList<string> dllPaths)
    {
        if (dllPaths.Count == 0)
        {
            return [];
        }

        var result = new List<IntPtr>(dllPaths.Count);
        var seenHandles = new HashSet<IntPtr>();
        for (var i = 0; i < dllPaths.Count; i++)
        {
            var path = dllPaths[i];
            var handle = IntPtr.Zero;
            lock (ModuleLock)
            {
                if (LoadedModules.TryGetValue(path, out var entry) && entry.Handle != IntPtr.Zero)
                {
                    handle = entry.Handle;
                }
            }

            if (handle == IntPtr.Zero
                && NativeMethods.GetModuleHandleExByName(
                    NativeMethods.GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                    path,
                    out var queried))
            {
                handle = queried;
            }

            if (handle != IntPtr.Zero && seenHandles.Add(handle))
            {
                result.Add(handle);
            }
        }

        return result;
    }

    internal virtual IntPtr CreateActCtx(ref ActCtx context) => NativeMethods.CreateActCtx(ref context);

    internal virtual IntPtr CreateActCtx(ref ActCtx context, string? defaultAssemblyPath, out int errorCode)
    {
        NanoServerOptions.EnsureInitialized();
        var pending = _pendingCreateContext;
        var options = pending?.Options ?? ComActivationOptions.Default;

        // Normalize and validate manifest structure, GUIDs, duplicate CLSIDs, and path safety before touching OS state.
        if (!ComManifestParser.TryParse(
            context.lpSource,
            context.lpAssemblyDirectory,
            defaultAssemblyPath,
            out var parsedManifest,
            out errorCode))
        {
            return (IntPtr)(-1);
        }

        pending?.SetParsedManifest(parsedManifest);

        if (options.ForceInProcessFallback)
        {
            if (!NanoServerOptions.EnableSxsFallback)
            {
                throw new InvalidOperationException(
                    "ComActivationOptions.ForceInProcessFallback cannot be used when NanoServerOptions.EnableSxsFallback is disabled.");
            }

            return CreateFallbackContext(parsedManifest!, options, out errorCode);
        }

        var handle = CreateActCtx(ref context);
        if (handle != (IntPtr)(-1))
        {
            errorCode = 0;
            return handle;
        }

        var lastError = GetLastWin32Error();
        if (!NanoServerOptions.EnableSxsFallback || !IsFallbackTriggerError(lastError))
        {
            errorCode = lastError;
            return (IntPtr)(-1);
        }

        return CreateFallbackContext(parsedManifest!, options, out errorCode);
    }

    internal IntPtr CreateActCtxWithOptions(
        ref ActCtx context,
        string? defaultAssemblyPath,
        ComActivationOptions options,
        out ComManifestParser.ParsedManifest? parsedManifest,
        out int errorCode)
    {
        var pending = new PendingCreateContext(options);
        var previous = SwapPendingCreateContext(pending);
        try
        {
            var handle = CreateActCtx(ref context, defaultAssemblyPath, out errorCode);
            parsedManifest = pending.ParsedManifest;
            return handle;
        }
        finally
        {
            SwapPendingCreateContext(previous);
        }
    }

    internal virtual bool ActivateActCtx(IntPtr handle, out IntPtr cookie)
    {
        if (_fallbackContexts.TryGetValue(handle, out var fallbackContext))
        {
            return ActivateFallbackContext(handle, fallbackContext, out cookie);
        }

        return NativeMethods.ActivateActCtx(handle, out cookie);
    }

    internal virtual bool DeactivateActCtx(IntPtr cookie)
    {
        if (_fallbackActivations.TryGetValue(cookie, out var activation))
        {
            return DeactivateFallbackContext(activation);
        }

        return NativeMethods.DeactivateActCtx(0, cookie);
    }

    internal virtual bool GetCurrentActCtx(out IntPtr handle)
    {
        var threadId = NativeMethods.GetCurrentThreadId();
        if (_threadFallbackStacks.TryGetValue(threadId, out var stack) && stack.Count > 0)
        {
            if (NativeMethods.GetCurrentActCtx(out var nativeCurrent) && nativeCurrent != IntPtr.Zero)
            {
                handle = nativeCurrent;
                return true;
            }

            var topHandle = stack[stack.Count - 1].ContextHandle;
            if (_fallbackContexts.TryGetValue(topHandle, out var context))
            {
                context.AddReference();
            }

            handle = topHandle;
            return true;
        }

        return NativeMethods.GetCurrentActCtx(out handle);
    }

    internal virtual void ReleaseActCtx(IntPtr handle)
    {
        if (_fallbackContexts.TryGetValue(handle, out var context))
        {
            if (context.ReleaseReference() <= 0)
            {
                _fallbackContexts.TryRemove(handle, out _);
            }

            return;
        }

        NativeMethods.ReleaseActCtx(handle);
    }

    internal virtual int GetLastWin32Error() => Marshal.GetLastWin32Error();

    internal virtual IntPtr LoadLibrary(string dllPath)
    {
        var fullPath = Path.GetFullPath(dllPath);
        var pending = _pendingLeaseContext;
        var unloadOnDeactivate = pending?.UnloadOnDeactivate ?? false;

        lock (ModuleLock)
        {
            if (!LoadedModules.TryGetValue(fullPath, out var entry) || entry.Handle == IntPtr.Zero)
            {
                var handle = LoadLibraryCore(fullPath);
                if (handle == IntPtr.Zero)
                {
                    return IntPtr.Zero;
                }

                entry = new LoadedModuleEntry(fullPath, handle);
                LoadedModules[fullPath] = entry;
            }

            if (!unloadOnDeactivate)
            {
                entry.RetainPermanently = true;
            }

            if (pending is not null)
            {
                entry.ActiveLeaseCount++;
                entry.PendingDeferredUnload = false;
                pending.AcquiredLease = new ModuleLease(entry, unloadOnDeactivate);
            }

            return entry.Handle;
        }
    }

    internal virtual IntPtr LoadLibraryCore(string fullPath)
    {
        var handle = NativeMethods.LoadLibraryEx(fullPath, IntPtr.Zero, NativeMethods.LOAD_LIBRARY_SEARCH_DLL_AND_DEFAULT_DIRS);
        if (handle == IntPtr.Zero)
        {
            handle = NativeMethods.LoadLibraryEx(fullPath, IntPtr.Zero, NativeMethods.LOAD_WITH_ALTERED_SEARCH_PATH);
        }

        if (handle == IntPtr.Zero)
        {
            handle = NativeMethods.LoadLibraryEx(fullPath, IntPtr.Zero, 0);
        }

        if (handle != IntPtr.Zero)
        {
            NanoServerOle2Shim.OnModuleLoaded(handle);
        }

        return handle;
    }

    internal virtual bool FreeLibraryCore(IntPtr moduleHandle)
        => NativeMethods.FreeLibrary(moduleHandle);

    internal virtual void FreeUnusedLibrariesCore()
        => NativeMethods.CoFreeUnusedLibrariesEx(0, 0);

    internal virtual int QueryDllCanUnloadNow(IntPtr moduleHandle)
    {
        var proc = GetProcAddress(moduleHandle, "DllCanUnloadNow");
        if (proc == IntPtr.Zero)
        {
            return NativeMethods.S_OK;
        }

        var callback = Marshal.GetDelegateForFunctionPointer<NativeMethods.DllCanUnloadNowDelegate>(proc);
        return callback();
    }

    internal virtual IntPtr GetProcAddress(IntPtr moduleHandle, string procName)
        => NativeMethods.GetProcAddress(moduleHandle, procName);

    internal virtual int GetClassObject(IntPtr procAddress, ref Guid clsid, ref Guid iid, out IntPtr classFactory)
    {
        var callback = Marshal.GetDelegateForFunctionPointer<NativeMethods.DllGetClassObjectDelegate>(procAddress);
        return callback(ref clsid, ref iid, out classFactory);
    }

    internal virtual int RegisterClassObject(ref Guid clsid, IntPtr classFactory, out uint registrationToken)
        => NativeMethods.CoRegisterClassObject(
            ref clsid,
            classFactory,
            NativeMethods.CLSCTX_INPROC_SERVER,
            NativeMethods.REGCLS_MULTIPLEUSE,
            out registrationToken);

    internal virtual int InitializeMultithreadedApartment()
        => NativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_MULTITHREADED);

    internal virtual int InitializeSingleThreadedApartment()
        => NativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_APARTMENTTHREADED);

    internal virtual int InitializeApartment(uint dwCoInit)
        => dwCoInit == NativeMethods.COINIT_APARTMENTTHREADED
            ? InitializeSingleThreadedApartment()
            : InitializeMultithreadedApartment();

    internal virtual void UninitializeApartment()
        => NativeMethods.CoUninitialize();

    internal virtual int LoadTypeLibrary(string typeLibPath, out IntPtr typeLib)
        => NativeMethods.LoadTypeLibEx(typeLibPath, NativeMethods.REGKIND_NONE, out typeLib);

    internal virtual void ReleaseTypeLibrary(IntPtr typeLib)
        => Marshal.Release(typeLib);

    internal virtual int RevokeClassObject(uint registrationToken)
        => NativeMethods.CoRevokeClassObject(registrationToken);

    internal virtual void ReleaseClassObject(IntPtr classFactory)
        => Marshal.Release(classFactory);

    internal ModuleLease? AcquireModuleLease(string dllPath, bool unloadOnDeactivate)
    {
        var context = new PendingLeaseContext(unloadOnDeactivate);
        var previous = SwapPendingLeaseContext(context);
        try
        {
            var handle = LoadLibrary(dllPath);
            if (handle == IntPtr.Zero)
            {
                context.AcquiredLease?.TryMarkReleased();
                return null;
            }

            if (context.AcquiredLease is not null)
            {
                return context.AcquiredLease;
            }

            // Custom LoadLibrary override returned a handle without invoking base.LoadLibrary.
            return new ModuleLease(new LoadedModuleEntry(Path.GetFullPath(dllPath), handle) { RetainPermanently = true }, unloadOnDeactivate: false);
        }
        finally
        {
            _ = SwapPendingLeaseContext(previous);
        }
    }

    internal void ReleaseModuleLeases(List<ModuleLease> leases, bool forceFreeUnusedLibraries = false)
    {
        var anyUnloadRequested = forceFreeUnusedLibraries || leases.Exists(static lease => lease.UnloadOnDeactivate);

        if (anyUnloadRequested)
        {
            FreeUnusedLibrariesCore();
        }

        var unmappedAny = false;
        lock (ModuleLock)
        {
            for (var i = leases.Count - 1; i >= 0; i--)
            {
                var lease = leases[i];
                if (!lease.TryMarkReleased())
                {
                    continue;
                }

                var entry = lease.Entry;
                if (entry.ActiveLeaseCount > 0)
                {
                    entry.ActiveLeaseCount--;
                }

                if (entry.ActiveLeaseCount != 0 || entry.RetainPermanently)
                {
                    continue;
                }

                var unmapped = TryUnloadModuleEntryUnderLock(entry, this);
                if (!unmapped)
                {
                    entry.PendingDeferredUnload = true;
                }

                unmappedAny |= unmapped;
            }
        }

        leases.Clear();
        if (anyUnloadRequested && unmappedAny)
        {
            FreeUnusedLibrariesCore();
        }
    }

    internal IReadOnlyList<IntPtr> GetFallbackLoadedHandles(IntPtr cookie)
    {
        if (!_fallbackActivations.TryGetValue(cookie, out var activation))
        {
            return [];
        }

        var handles = new List<IntPtr>(activation.ModuleLeases.Count);
        var seen = new HashSet<IntPtr>();
        for (var i = 0; i < activation.ModuleLeases.Count; i++)
        {
            var h = activation.ModuleLeases[i].Handle;
            if (h != IntPtr.Zero && seen.Add(h))
            {
                handles.Add(h);
            }
        }

        return handles;
    }

    internal int ConsumeLastError()
    {
        var threadId = NativeMethods.GetCurrentThreadId();
        return _lastFallbackErrors.TryRemove(threadId, out var error) ? error : 0;
    }

    private static PendingLeaseContext? SwapPendingLeaseContext(PendingLeaseContext? next)
    {
        var previous = _pendingLeaseContext;
        _pendingLeaseContext = next;
        return previous;
    }

    private static PendingCreateContext? SwapPendingCreateContext(PendingCreateContext? next)
    {
        var previous = _pendingCreateContext;
        _pendingCreateContext = next;
        return previous;
    }

    private static bool TryUnloadModuleEntryUnderLock(LoadedModuleEntry entry, ActivationContextApi api)
    {
        if (entry.Handle == IntPtr.Zero || entry.RetainPermanently || entry.ActiveLeaseCount > 0)
        {
            return false;
        }

        var canUnloadHr = api.QueryDllCanUnloadNow(entry.Handle);
        if (canUnloadHr != NativeMethods.S_OK)
        {
            return false;
        }

        LoadedModules.Remove(entry.FullPath);
        entry.PendingDeferredUnload = false;
        var handle = entry.Handle;
        entry.Handle = IntPtr.Zero;
        _ = api.FreeLibraryCore(handle);
        NanoServerOle2Shim.OnModuleUnloaded(handle);
        return true;
    }

    private static bool IsFallbackTriggerError(int errorCode)
        => errorCode is NativeMethods.ERROR_INVALID_FUNCTION
            or NativeMethods.ERROR_NOT_SUPPORTED
            or NativeMethods.ERROR_CALL_NOT_IMPLEMENTED;

    private static IntPtr AllocateFallbackToken()
        => new(Interlocked.Decrement(ref _nextFallbackToken));

    private static int NormalizeWin32Error(int errorCode, int defaultError)
        => errorCode != 0 ? errorCode : defaultError;

    private static uint SelectDesiredApartment(ComManifestParser.ComThreadingModel threadingModel)
    {
        if (threadingModel is ComManifestParser.ComThreadingModel.Apartment
            or ComManifestParser.ComThreadingModel.Single)
        {
            return NativeMethods.COINIT_APARTMENTTHREADED;
        }

        return Thread.CurrentThread.GetApartmentState() == ApartmentState.STA
            ? NativeMethods.COINIT_APARTMENTTHREADED
            : NativeMethods.COINIT_MULTITHREADED;
    }

    private static ApartmentRegistrationKey ResolveApartmentRegistrationKey(
        ComManifestParser.ComThreadingModel threadingModel,
        uint currentThreadId,
        out bool currentThreadNeedsComInit)
    {
        if (NativeMethods.CoGetApartmentType(out var aptType, out var aptQualifier) >= 0)
        {
            if (aptType is NativeMethods.APTTYPE_STA or NativeMethods.APTTYPE_MAINSTA
                || (aptType == NativeMethods.APTTYPE_NA
                    && aptQualifier is NativeMethods.APTTYPEQUALIFIER_NA_ON_STA or NativeMethods.APTTYPEQUALIFIER_NA_ON_MAINSTA))
            {
                currentThreadNeedsComInit = false;
                return new ApartmentRegistrationKey(isMta: false, currentThreadId, Guid.Empty);
            }

            if ((aptType == NativeMethods.APTTYPE_MTA && aptQualifier != NativeMethods.APTTYPEQUALIFIER_IMPLICIT_MTA)
                || (aptType == NativeMethods.APTTYPE_NA && aptQualifier == NativeMethods.APTTYPEQUALIFIER_NA_ON_MTA))
            {
                currentThreadNeedsComInit = false;
                return new ApartmentRegistrationKey(isMta: true, 0u, Guid.Empty);
            }
        }

        currentThreadNeedsComInit = true;
        var desiredCoInit = SelectDesiredApartment(threadingModel);
        var isMta = desiredCoInit != NativeMethods.COINIT_APARTMENTTHREADED;
        return new ApartmentRegistrationKey(isMta, isMta ? 0u : currentThreadId, Guid.Empty);
    }

    private void SetLastFallbackError(int errorCode)
    {
        var threadId = NativeMethods.GetCurrentThreadId();
        _lastFallbackErrors[threadId] = errorCode;
    }

    private IntPtr CreateFallbackContext(
        ComManifestParser.ParsedManifest parsedManifest,
        ComActivationOptions options,
        out int errorCode)
    {
        var handle = AllocateFallbackToken();
        _fallbackContexts[handle] = new FallbackContext(parsedManifest, options.Clone());
        errorCode = 0;
        return handle;
    }

    private bool ActivateFallbackContext(IntPtr handle, FallbackContext fallbackContext, out IntPtr cookie)
    {
        var options = fallbackContext.Options;
        var manifest = fallbackContext.Manifest;
        var activeRegistrations = new List<SharedClassRegistrationLease>(manifest.Classes.Count);
        var moduleLeases = new List<ModuleLease>(manifest.ReferencedDllPaths.Count);
        var loadedTypeLibs = new List<IntPtr>(manifest.TypeLibs.Count);
        var acquiredComInitCount = 0;

        if (!AcquireManifestModuleLeases(manifest, options.UnloadModulesOnDeactivate, moduleLeases, out var moduleMap))
        {
            ReleaseModuleLeases(moduleLeases);
            cookie = IntPtr.Zero;
            return false;
        }

        if (!TryRegisterAllManifestClasses(manifest, moduleMap, options.IgnoreUnavailableProxyStubClasses, activeRegistrations, ref acquiredComInitCount)
            || (options.LoadTypeLibsOnFallback && !TryLoadManifestTypeLibs(manifest, loadedTypeLibs))
            || !TryAcquireOptionalFallbackRegistryLease(manifest, options, out var registryLease))
        {
            UnwindFallbackSetup(activeRegistrations, loadedTypeLibs, registryLease: null, moduleLeases, ref acquiredComInitCount);
            cookie = IntPtr.Zero;
            return false;
        }

        var threadId = NativeMethods.GetCurrentThreadId();
        cookie = AllocateFallbackToken();
        var activation = new FallbackActivation(
            cookie,
            handle,
            threadId,
            activeRegistrations,
            moduleLeases,
            loadedTypeLibs,
            registryLease,
            acquiredComInitCount);
        _fallbackActivations[cookie] = activation;
        var stack = _threadFallbackStacks.GetOrAdd(threadId, _ => []);
        stack.Add(activation);
        return true;
    }

    private bool TryRegisterAllManifestClasses(
        ComManifestParser.ParsedManifest manifest,
        Dictionary<string, IntPtr> moduleMap,
        bool ignoreUnavailableProxyStub,
        List<SharedClassRegistrationLease> activeRegistrations,
        ref int acquiredComInitCount)
    {
        var currentThreadId = NativeMethods.GetCurrentThreadId();
        for (var i = 0; i < manifest.Classes.Count; i++)
        {
            var registration = manifest.Classes[i];
            if (!TryRegisterSingleClass(
                registration,
                moduleMap[registration.DllPath],
                currentThreadId,
                ignoreUnavailableProxyStub,
                ref acquiredComInitCount,
                out var activeRegistration))
            {
                return false;
            }

            if (activeRegistration is not null)
            {
                activeRegistrations.Add(activeRegistration);
            }
        }

        return true;
    }

    private bool TryAcquireOptionalFallbackRegistryLease(
        ComManifestParser.ParsedManifest manifest,
        ComActivationOptions options,
        out CurrentUserRegistryLeaseManager.RegistryScopeLease? registryLease)
    {
        if (CurrentUserRegistryLeaseManager.TryAcquireFallbackRegistryLease(
            manifest,
            options.RegisterCurrentUserProgIdsOnFallback,
            options.RegisterCurrentUserTypeLibsOnFallback,
            options.SkipCurrentUserRegistrationWhenPresentInClassesRoot,
            out registryLease))
        {
            return true;
        }

        SetLastFallbackError(NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX);
        return false;
    }

    private bool AcquireManifestModuleLeases(
        ComManifestParser.ParsedManifest manifest,
        bool unloadOnDeactivate,
        List<ModuleLease> moduleLeases,
        out Dictionary<string, IntPtr> moduleMap)
    {
        moduleMap = new Dictionary<string, IntPtr>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < manifest.Classes.Count; i++)
        {
            var dllPath = manifest.Classes[i].DllPath;
            if (moduleMap.ContainsKey(dllPath))
            {
                continue;
            }

            var lease = AcquireModuleLease(dllPath, unloadOnDeactivate);
            if (lease is null || lease.Handle == IntPtr.Zero)
            {
                SetLastFallbackError(NormalizeWin32Error(GetLastWin32Error(), NativeMethods.ERROR_MOD_NOT_FOUND));
                return false;
            }

            moduleLeases.Add(lease);
            moduleMap[dllPath] = lease.Handle;
        }

        return true;
    }

    private bool TryLoadManifestTypeLibs(
        ComManifestParser.ParsedManifest manifest,
        List<IntPtr> loadedTypeLibs)
    {
        for (var i = 0; i < manifest.TypeLibs.Count; i++)
        {
            var tlb = manifest.TypeLibs[i];
            var loadPath = string.IsNullOrEmpty(tlb.ResourceId)
                ? tlb.DllPath
                : tlb.DllPath + Path.DirectorySeparatorChar + tlb.ResourceId;
            var hr = LoadTypeLibrary(loadPath, out var typeLib);
            if (hr < 0 && !string.IsNullOrEmpty(tlb.ResourceId))
            {
                hr = LoadTypeLibrary(tlb.DllPath, out typeLib);
            }

            if (hr >= 0 && typeLib != IntPtr.Zero)
            {
                loadedTypeLibs.Add(typeLib);
            }
            else if (!File.Exists(tlb.DllPath))
            {
                SetLastFallbackError(NativeMethods.ERROR_FILE_NOT_FOUND);
                return false;
            }
        }

        return true;
    }

    private bool TryRegisterSingleClass(
        ComManifestParser.ManifestComClassEntry registration,
        IntPtr moduleHandle,
        uint currentThreadId,
        bool ignoreUnavailableProxyStub,
        ref int acquiredComInitCount,
        out SharedClassRegistrationLease? activeRegistration)
    {
        activeRegistration = null;
        var normalizedDllPath = Path.GetFullPath(registration.DllPath);

        lock (_registrationLock)
        {
            var baseApt = ResolveApartmentRegistrationKey(registration.ThreadingModel, currentThreadId, out var needsComInit);
            var lookupKey = new ApartmentRegistrationKey(baseApt.IsMta, baseApt.StaThreadId, registration.Clsid);
            if (_apartmentRegistrations.TryGetValue(lookupKey, out var existingLease))
            {
                return TryReuseExistingApartmentLeaseUnderLock(
                    existingLease,
                    normalizedDllPath,
                    registration.ThreadingModel,
                    needsComInit,
                    ref acquiredComInitCount,
                    out activeRegistration);
            }

            var procAddress = GetProcAddress(moduleHandle, "DllGetClassObject");
            if (procAddress == IntPtr.Zero)
            {
                SetLastFallbackError(NormalizeWin32Error(GetLastWin32Error(), NativeMethods.ERROR_PROC_NOT_FOUND));
                return false;
            }

            if (!AcquireClassFactory(
                procAddress,
                registration.Clsid,
                registration.IsKnownProxyStub && ignoreUnavailableProxyStub,
                out var classFactory,
                out var skippedUnavailableProxyStub))
            {
                return false;
            }

            if (skippedUnavailableProxyStub)
            {
                return true;
            }

            if (!RegisterFactoryWithRetry(
                registration.Clsid,
                registration.ThreadingModel,
                classFactory,
                ref acquiredComInitCount,
                out var registrationToken))
            {
                ReleaseClassObject(classFactory);
                return false;
            }

            var finalApt = ResolveApartmentRegistrationKey(registration.ThreadingModel, currentThreadId, out _);
            var finalKey = new ApartmentRegistrationKey(finalApt.IsMta, finalApt.StaThreadId, registration.Clsid);
            var createdLease = new SharedClassRegistrationLease(finalKey, normalizedDllPath, classFactory, registrationToken);
            _apartmentRegistrations[finalKey] = createdLease;
            activeRegistration = createdLease;
            return true;
        }
    }

    private bool TryReuseExistingApartmentLeaseUnderLock(
        SharedClassRegistrationLease existingLease,
        string normalizedDllPath,
        ComManifestParser.ComThreadingModel threadingModel,
        bool needsComInit,
        ref int acquiredComInitCount,
        out SharedClassRegistrationLease? activeRegistration)
    {
        activeRegistration = null;
        if (!string.Equals(existingLease.NormalizedDllPath, normalizedDllPath, StringComparison.OrdinalIgnoreCase))
        {
            SetLastFallbackError(NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX);
            return false;
        }

        if (needsComInit)
        {
            var desiredCoInit = SelectDesiredApartment(threadingModel);
            var initHr = InitializeApartment(desiredCoInit);
            if (initHr < 0)
            {
                SetLastFallbackError(initHr);
                return false;
            }

            acquiredComInitCount++;
        }

        existingLease.LeaseCount++;
        activeRegistration = existingLease;
        return true;
    }

    private bool AcquireClassFactory(
        IntPtr procAddress,
        Guid clsid,
        bool allowSkipUnavailableProxyStub,
        out IntPtr classFactory,
        out bool skippedUnavailableProxyStub)
    {
        skippedUnavailableProxyStub = false;
        var iid = IidIClassFactory;
        var hr = GetClassObject(procAddress, ref clsid, ref iid, out classFactory);
        if (hr == NativeMethods.E_NOINTERFACE)
        {
            iid = IidIUnknown;
            hr = GetClassObject(procAddress, ref clsid, ref iid, out classFactory);
        }

        if (hr == NativeMethods.CLASS_E_CLASSNOTAVAILABLE && allowSkipUnavailableProxyStub)
        {
            classFactory = IntPtr.Zero;
            skippedUnavailableProxyStub = true;
            return true;
        }

        if (hr < 0 || classFactory == IntPtr.Zero)
        {
            SetLastFallbackError(hr < 0 ? hr : NativeMethods.CLASS_E_CLASSNOTAVAILABLE);
            classFactory = IntPtr.Zero;
            return false;
        }

        return true;
    }

    private bool RegisterFactoryWithRetry(
        Guid clsid,
        ComManifestParser.ComThreadingModel threadingModel,
        IntPtr classFactory,
        ref int acquiredComInitCount,
        out uint registrationToken)
    {
        var hr = RegisterClassObject(ref clsid, classFactory, out registrationToken);
        if (hr == NativeMethods.CO_E_NOTINITIALIZED)
        {
            var desiredCoInit = SelectDesiredApartment(threadingModel);
            var initHr = InitializeApartment(desiredCoInit);
            if (initHr < 0)
            {
                // Propagate RPC_E_CHANGED_MODE or other initialization failure without altering existing apartment state.
                SetLastFallbackError(initHr);
                return false;
            }

            acquiredComInitCount++;
            hr = RegisterClassObject(ref clsid, classFactory, out registrationToken);
        }

        if (hr < 0)
        {
            SetLastFallbackError(hr);
            return false;
        }

        return true;
    }

    private void UnwindFallbackSetup(
        List<SharedClassRegistrationLease> activeRegistrations,
        List<IntPtr> loadedTypeLibs,
        CurrentUserRegistryLeaseManager.RegistryScopeLease? registryLease,
        List<ModuleLease> moduleLeases,
        ref int acquiredComInitCount)
    {
        registryLease?.Rollback();

        for (var i = loadedTypeLibs.Count - 1; i >= 0; i--)
        {
            ReleaseTypeLibrary(loadedTypeLibs[i]);
        }

        loadedTypeLibs.Clear();

        lock (_registrationLock)
        {
            for (var i = activeRegistrations.Count - 1; i >= 0; i--)
            {
                var entry = activeRegistrations[i];
                if (entry.LeaseCount > 1)
                {
                    entry.LeaseCount--;
                    continue;
                }

                entry.LeaseCount = 0;
                _apartmentRegistrations.Remove(entry.Key);
                _ = RevokeClassObject(entry.RegistrationToken);
                ReleaseClassObject(entry.ClassFactory);
            }
        }

        activeRegistrations.Clear();
        ReleaseModuleLeases(moduleLeases);

        while (acquiredComInitCount > 0)
        {
            UninitializeApartment();
            acquiredComInitCount--;
        }
    }

    private bool DeactivateFallbackContext(FallbackActivation activation)
    {
        var threadId = NativeMethods.GetCurrentThreadId();
        if (threadId != activation.OwnerThreadId
            || !_threadFallbackStacks.TryGetValue(threadId, out var stack)
            || stack.Count == 0
            || stack[stack.Count - 1].Cookie != activation.Cookie)
        {
            SetLastFallbackError(NativeMethods.ERROR_INVALID_PARAMETER);
            return false;
        }

        if (!TryRevokeFallbackRegistrations(activation.ActiveRegistrations))
        {
            return false;
        }

        activation.RegistryLease?.Rollback();

        for (var i = activation.LoadedTypeLibs.Count - 1; i >= 0; i--)
        {
            ReleaseTypeLibrary(activation.LoadedTypeLibs[i]);
            activation.LoadedTypeLibs.RemoveAt(i);
        }

        ReleaseModuleLeases(activation.ModuleLeases);

        while (activation.AcquiredComInitCount > 0)
        {
            UninitializeApartment();
            activation.AcquiredComInitCount--;
        }

        stack.RemoveAt(stack.Count - 1);
        if (stack.Count == 0)
        {
            _threadFallbackStacks.TryRemove(threadId, out _);
        }

        _fallbackActivations.TryRemove(activation.Cookie, out _);
        return true;
    }

    private bool TryRevokeFallbackRegistrations(List<SharedClassRegistrationLease> activeRegistrations)
    {
        lock (_registrationLock)
        {
            for (var i = activeRegistrations.Count - 1; i >= 0; i--)
            {
                var entry = activeRegistrations[i];
                if (entry.LeaseCount > 1)
                {
                    entry.LeaseCount--;
                    activeRegistrations.RemoveAt(i);
                    continue;
                }

                var hr = RevokeClassObject(entry.RegistrationToken);
                if (hr < 0)
                {
                    SetLastFallbackError(hr);
                    return false;
                }

                entry.LeaseCount = 0;
                _apartmentRegistrations.Remove(entry.Key);
                ReleaseClassObject(entry.ClassFactory);
                activeRegistrations.RemoveAt(i);
            }
        }

        return true;
    }

    internal readonly struct ApartmentRegistrationKey : IEquatable<ApartmentRegistrationKey>
    {
        internal ApartmentRegistrationKey(bool isMta, uint staThreadId, Guid clsid)
        {
            IsMta = isMta;
            StaThreadId = isMta ? 0u : staThreadId;
            Clsid = clsid;
        }

        internal bool IsMta { get; }

        internal uint StaThreadId { get; }

        internal Guid Clsid { get; }

        public bool Equals(ApartmentRegistrationKey other)
            => IsMta == other.IsMta && StaThreadId == other.StaThreadId && Clsid == other.Clsid;

        public override bool Equals(object? obj)
            => obj is ApartmentRegistrationKey other && Equals(other);

        public override int GetHashCode()
        {
            unchecked
            {
                var hash = IsMta ? 1 : 0;
                hash = (hash * 397) ^ (int)StaThreadId;
                hash = (hash * 397) ^ Clsid.GetHashCode();
                return hash;
            }
        }
    }

    internal sealed class ModuleLease
    {
        private int _released;

        internal ModuleLease(LoadedModuleEntry entry, bool unloadOnDeactivate)
        {
            Entry = entry;
            UnloadOnDeactivate = unloadOnDeactivate;
        }

        internal LoadedModuleEntry Entry { get; }

        internal IntPtr Handle => Entry.Handle;

        internal bool UnloadOnDeactivate { get; }

        internal bool TryMarkReleased()
            => Interlocked.CompareExchange(ref _released, 1, 0) == 0;
    }

    internal sealed class LoadedModuleEntry
    {
        internal LoadedModuleEntry(string fullPath, IntPtr handle)
        {
            FullPath = fullPath;
            Handle = handle;
        }

        internal string FullPath { get; }

        internal IntPtr Handle { get; set; }

        internal int ActiveLeaseCount { get; set; }

        internal bool RetainPermanently { get; set; }

        internal bool PendingDeferredUnload { get; set; }
    }

    private sealed class PendingLeaseContext
    {
        internal PendingLeaseContext(bool unloadOnDeactivate)
        {
            UnloadOnDeactivate = unloadOnDeactivate;
        }

        internal bool UnloadOnDeactivate { get; }

        internal ModuleLease? AcquiredLease { get; set; }
    }

    private sealed class SharedClassRegistrationLease
    {
        internal SharedClassRegistrationLease(
            ApartmentRegistrationKey key,
            string normalizedDllPath,
            IntPtr classFactory,
            uint registrationToken)
        {
            Key = key;
            NormalizedDllPath = normalizedDllPath;
            ClassFactory = classFactory;
            RegistrationToken = registrationToken;
            LeaseCount = 1;
        }

        internal ApartmentRegistrationKey Key { get; }

        internal string NormalizedDllPath { get; }

        internal IntPtr ClassFactory { get; }

        internal uint RegistrationToken { get; }

        internal int LeaseCount { get; set; }
    }

    private sealed class FallbackContext
    {
        private int _referenceCount = 1;

        internal FallbackContext(ComManifestParser.ParsedManifest manifest, ComActivationOptions options)
        {
            Manifest = manifest;
            Options = options;
        }

        internal ComManifestParser.ParsedManifest Manifest { get; }

        internal ComActivationOptions Options { get; }

        internal void AddReference() => Interlocked.Increment(ref _referenceCount);

        internal int ReleaseReference() => Interlocked.Decrement(ref _referenceCount);
    }

    private sealed class PendingCreateContext
    {
        internal PendingCreateContext(ComActivationOptions options)
        {
            Options = options;
        }

        internal ComActivationOptions Options { get; }

        internal ComManifestParser.ParsedManifest? ParsedManifest { get; private set; }

        internal void SetParsedManifest(ComManifestParser.ParsedManifest? parsedManifest)
            => ParsedManifest = parsedManifest;
    }

#pragma warning disable S107 // Internal activation record aggregates lifecycle state cleanly.
    private sealed class FallbackActivation
    {
        internal FallbackActivation(
            IntPtr cookie,
            IntPtr contextHandle,
            uint ownerThreadId,
            List<SharedClassRegistrationLease> activeRegistrations,
            List<ModuleLease> moduleLeases,
            List<IntPtr> loadedTypeLibs,
            CurrentUserRegistryLeaseManager.RegistryScopeLease? registryLease,
            int acquiredComInitCount)
        {
            Cookie = cookie;
            ContextHandle = contextHandle;
            OwnerThreadId = ownerThreadId;
            ActiveRegistrations = activeRegistrations;
            ModuleLeases = moduleLeases;
            LoadedTypeLibs = loadedTypeLibs;
            RegistryLease = registryLease;
            AcquiredComInitCount = acquiredComInitCount;
        }

        internal IntPtr Cookie { get; }

        internal IntPtr ContextHandle { get; }

        internal uint OwnerThreadId { get; }

        internal List<SharedClassRegistrationLease> ActiveRegistrations { get; }

        internal List<ModuleLease> ModuleLeases { get; }

        internal List<IntPtr> LoadedTypeLibs { get; }

        internal CurrentUserRegistryLeaseManager.RegistryScopeLease? RegistryLease { get; }

        internal int AcquiredComInitCount { get; set; }
    }
#pragma warning restore S107
}
#pragma warning restore SA1121
#pragma warning restore MA0158
