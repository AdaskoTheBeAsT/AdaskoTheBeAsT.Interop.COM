using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1121 // Use built-in type alias
// An operation-scoped native boundary permits deterministic failure tests without global hooks
// and provides the manifest-driven in-process class-factory fallback on Windows Nano Server.
internal class ActivationContextApi
{
#pragma warning disable MA0176 // Optimize guid creation
    private static readonly Guid IidIClassFactory = new("00000001-0000-0000-C000-000000000046");
    private static readonly Guid IidIUnknown = new("00000000-0000-0000-C000-000000000046");
#pragma warning restore MA0176
    private static readonly ConcurrentDictionary<string, IntPtr> LoadedModules = new(StringComparer.OrdinalIgnoreCase);

    private static long _nextFallbackToken = -1;

    private readonly ConcurrentDictionary<IntPtr, FallbackContext> _fallbackContexts = new();
    private readonly ConcurrentDictionary<IntPtr, FallbackActivation> _fallbackActivations = new();
    private readonly ConcurrentDictionary<uint, List<FallbackActivation>> _threadFallbackStacks = new();
    private readonly ConcurrentDictionary<uint, int> _lastFallbackErrors = new();

    internal static ActivationContextApi Instance { get; } = new();

    internal virtual IntPtr CreateActCtx(ref ActCtx context) => NativeMethods.CreateActCtx(ref context);

    internal IntPtr CreateActCtx(ref ActCtx context, string? defaultAssemblyPath, out int errorCode)
    {
        var handle = CreateActCtx(ref context);
        if (handle != (IntPtr)(-1))
        {
            errorCode = 0;
            return handle;
        }

        var lastError = GetLastWin32Error();
        if (!IsFallbackTriggerError(lastError))
        {
            errorCode = lastError;
            return (IntPtr)(-1);
        }

        return CreateFallbackContext(ref context, defaultAssemblyPath, out errorCode);
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
        if (LoadedModules.TryGetValue(fullPath, out var cached) && cached != IntPtr.Zero)
        {
            return cached;
        }

        var handle = NativeMethods.LoadLibraryEx(fullPath, IntPtr.Zero, NativeMethods.LOAD_WITH_ALTERED_SEARCH_PATH);
        if (handle == IntPtr.Zero)
        {
            handle = NativeMethods.LoadLibraryEx(fullPath, IntPtr.Zero, 0);
        }

        if (handle != IntPtr.Zero)
        {
            LoadedModules.TryAdd(fullPath, handle);
        }

        return handle;
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

    internal virtual int RevokeClassObject(uint registrationToken)
        => NativeMethods.CoRevokeClassObject(registrationToken);

    internal virtual void ReleaseClassObject(IntPtr classFactory)
        => Marshal.Release(classFactory);

    internal int ConsumeLastError()
    {
        var threadId = NativeMethods.GetCurrentThreadId();
        return _lastFallbackErrors.TryRemove(threadId, out var error) ? error : 0;
    }

    private static bool IsFallbackTriggerError(int errorCode)
        => errorCode is NativeMethods.ERROR_INVALID_FUNCTION
            or NativeMethods.ERROR_NOT_SUPPORTED
            or NativeMethods.ERROR_CALL_NOT_IMPLEMENTED;

    private static IntPtr AllocateFallbackToken()
        => new(Interlocked.Decrement(ref _nextFallbackToken));

    private static int NormalizeWin32Error(int errorCode, int defaultError)
        => errorCode != 0 ? errorCode : defaultError;

#pragma warning disable SCS0018, SEC0116 // Manifest and assembly paths are caller-supplied configuration by design.
    private static bool TryParseManifest(
        string manifestPath,
        string? assemblyDirectory,
        string? defaultAssemblyPath,
        out List<ComClassRegistration> registrations,
        out int errorCode)
    {
        registrations = [];
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            return ParseManifestCore(
                Path.GetFullPath(manifestPath),
                assemblyDirectory,
                defaultAssemblyPath,
                visited,
                registrations,
                out errorCode);
        }
        catch (Exception ex) when (ex is XmlException or IOException or UnauthorizedAccessException or ArgumentException)
        {
            errorCode = ex is FileNotFoundException or DirectoryNotFoundException
                ? NativeMethods.ERROR_FILE_NOT_FOUND
                : NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }
    }

    private static bool ParseManifestCore(
        string manifestPath,
        string? assemblyDirectory,
        string? defaultAssemblyPath,
        HashSet<string> visitedManifests,
        List<ComClassRegistration> registrations,
        out int errorCode)
    {
        errorCode = 0;
        if (!visitedManifests.Add(manifestPath))
        {
            return true;
        }

        if (!File.Exists(manifestPath))
        {
            errorCode = NativeMethods.ERROR_FILE_NOT_FOUND;
            return false;
        }

        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        };

        using var stream = File.OpenRead(manifestPath);
        using var reader = XmlReader.Create(stream, settings);
        return ReadManifestNodes(reader, manifestPath, assemblyDirectory, defaultAssemblyPath, visitedManifests, registrations, out errorCode);
    }

    private static bool ReadManifestNodes(
        XmlReader reader,
        string manifestPath,
        string? assemblyDirectory,
        string? defaultAssemblyPath,
        HashSet<string> visitedManifests,
        List<ComClassRegistration> registrations,
        out int errorCode)
    {
        errorCode = 0;
        var hasAssemblyRoot = false;
        string? currentFileName = null;
        var fileDepth = -1;
        var dependentDepth = -1;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                ResetScopeOnEndElement(reader, ref currentFileName, ref fileDepth, ref dependentDepth);
                continue;
            }

            if (reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            if (!hasAssemblyRoot)
            {
                if (!string.Equals(reader.LocalName, "assembly", StringComparison.Ordinal))
                {
                    errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
                    return false;
                }

                hasAssemblyRoot = true;
                continue;
            }

            if (!ProcessElement(
                reader,
                manifestPath,
                assemblyDirectory,
                defaultAssemblyPath,
                ref currentFileName,
                ref fileDepth,
                ref dependentDepth,
                visitedManifests,
                registrations,
                out errorCode))
            {
                return false;
            }
        }

        if (!hasAssemblyRoot)
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        return true;
    }

    private static void ResetScopeOnEndElement(
        XmlReader reader,
        ref string? currentFileName,
        ref int fileDepth,
        ref int dependentDepth)
    {
        if (reader.Depth == fileDepth && string.Equals(reader.LocalName, "file", StringComparison.Ordinal))
        {
            currentFileName = null;
            fileDepth = -1;
        }
        else if (reader.Depth == dependentDepth && string.Equals(reader.LocalName, "dependentAssembly", StringComparison.Ordinal))
        {
            dependentDepth = -1;
        }
    }

#pragma warning disable S107 // Methods should not have too many parameters
    private static bool ProcessElement(
        XmlReader reader,
        string manifestPath,
        string? assemblyDirectory,
        string? defaultAssemblyPath,
        ref string? currentFileName,
        ref int fileDepth,
        ref int dependentDepth,
        HashSet<string> visitedManifests,
        List<ComClassRegistration> registrations,
        out int errorCode)
#pragma warning restore S107
    {
        errorCode = 0;
        if (string.Equals(reader.LocalName, "file", StringComparison.Ordinal))
        {
            var name = reader.GetAttribute("name");
            if (!reader.IsEmptyElement)
            {
                currentFileName = name;
                fileDepth = reader.Depth;
            }

            return true;
        }

        if (string.Equals(reader.LocalName, "dependentAssembly", StringComparison.Ordinal))
        {
            if (!reader.IsEmptyElement)
            {
                dependentDepth = reader.Depth;
            }

            return true;
        }

        if (string.Equals(reader.LocalName, "comClass", StringComparison.Ordinal))
        {
            return TryAddComClassRegistration(reader, currentFileName, assemblyDirectory, defaultAssemblyPath, registrations, out errorCode);
        }

        if (dependentDepth >= 0 && string.Equals(reader.LocalName, "assemblyIdentity", StringComparison.Ordinal))
        {
            return TryParseDependentAssembly(reader, manifestPath, assemblyDirectory, defaultAssemblyPath, visitedManifests, registrations, out errorCode);
        }

        return true;
    }

    private static bool TryAddComClassRegistration(
        XmlReader reader,
        string? currentFileName,
        string? assemblyDirectory,
        string? defaultAssemblyPath,
        List<ComClassRegistration> registrations,
        out int errorCode)
    {
        var clsidAttribute = reader.GetAttribute("clsid");
        if (string.IsNullOrWhiteSpace(clsidAttribute) || !Guid.TryParse(clsidAttribute, out var clsid))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        var dllPath = ResolveComDllPath(currentFileName, assemblyDirectory, defaultAssemblyPath);
        if (dllPath is null || string.IsNullOrWhiteSpace(dllPath))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        registrations.Add(new ComClassRegistration(dllPath, clsid));
        errorCode = 0;
        return true;
    }

    private static string? ResolveComDllPath(
        string? currentFileName,
        string? assemblyDirectory,
        string? defaultAssemblyPath)
    {
        if (currentFileName is not null && !string.IsNullOrWhiteSpace(currentFileName))
        {
            string? candidate = null;
            if (Path.IsPathRooted(currentFileName))
            {
                candidate = Path.GetFullPath(currentFileName);
            }
            else if (assemblyDirectory is not null && !string.IsNullOrWhiteSpace(assemblyDirectory))
            {
                candidate = Path.GetFullPath(Path.Combine(assemblyDirectory, currentFileName));
            }

            if (candidate is not null
                && !string.IsNullOrWhiteSpace(candidate)
                && (File.Exists(candidate) || string.IsNullOrWhiteSpace(defaultAssemblyPath)))
            {
                return candidate;
            }
        }

        return defaultAssemblyPath is null || string.IsNullOrWhiteSpace(defaultAssemblyPath)
            ? null
            : Path.GetFullPath(defaultAssemblyPath);
    }

    private static bool TryParseDependentAssembly(
        XmlReader reader,
        string currentManifestPath,
        string? assemblyDirectory,
        string? defaultAssemblyPath,
        HashSet<string> visitedManifests,
        List<ComClassRegistration> registrations,
        out int errorCode)
    {
        var dependencyName = reader.GetAttribute("name");
        if (dependencyName is null || string.IsNullOrWhiteSpace(dependencyName))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        var dependentManifestPath = ResolveDependentManifestPath(dependencyName, currentManifestPath, assemblyDirectory);
        if (dependentManifestPath is null)
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        return ParseManifestCore(
            dependentManifestPath,
            assemblyDirectory,
            defaultAssemblyPath,
            visitedManifests,
            registrations,
            out errorCode);
    }

    private static string? ResolveDependentManifestPath(
        string assemblyName,
        string currentManifestPath,
        string? assemblyDirectory)
    {
        var manifestFileName = assemblyName + ".manifest";
        if (assemblyDirectory is not null && !string.IsNullOrWhiteSpace(assemblyDirectory))
        {
            var directCandidate = Path.GetFullPath(Path.Combine(assemblyDirectory, manifestFileName));
            if (File.Exists(directCandidate))
            {
                return directCandidate;
            }

            var subFolderCandidate = Path.GetFullPath(Path.Combine(assemblyDirectory, assemblyName, manifestFileName));
            if (File.Exists(subFolderCandidate))
            {
                return subFolderCandidate;
            }
        }

        var manifestDirectory = Path.GetDirectoryName(currentManifestPath);
        if (manifestDirectory is not null && !string.IsNullOrWhiteSpace(manifestDirectory))
        {
            var siblingCandidate = Path.GetFullPath(Path.Combine(manifestDirectory, manifestFileName));
            if (File.Exists(siblingCandidate))
            {
                return siblingCandidate;
            }
        }

        return null;
    }
#pragma warning restore SCS0018, SEC0116

    private void SetLastFallbackError(int errorCode)
    {
        var threadId = NativeMethods.GetCurrentThreadId();
        _lastFallbackErrors[threadId] = errorCode;
    }

    private IntPtr CreateFallbackContext(ref ActCtx context, string? defaultAssemblyPath, out int errorCode)
    {
        if (string.IsNullOrWhiteSpace(context.lpSource) || !File.Exists(context.lpSource))
        {
            errorCode = NativeMethods.ERROR_FILE_NOT_FOUND;
            return (IntPtr)(-1);
        }

        if (!TryParseManifest(context.lpSource, context.lpAssemblyDirectory, defaultAssemblyPath, out var registrations, out errorCode))
        {
            return (IntPtr)(-1);
        }

        var handle = AllocateFallbackToken();
        _fallbackContexts[handle] = new FallbackContext(registrations);
        errorCode = 0;
        return handle;
    }

    private bool ActivateFallbackContext(IntPtr handle, FallbackContext fallbackContext, out IntPtr cookie)
    {
        var activeRegistrations = new List<ActiveClassRegistration>(fallbackContext.Registrations.Count);
        foreach (var registration in fallbackContext.Registrations)
        {
            var activeRegistration = TryRegisterSingleClass(registration);
            if (activeRegistration is null)
            {
                UnwindActiveRegistrations(activeRegistrations);
                cookie = IntPtr.Zero;
                return false;
            }

            activeRegistrations.Add(activeRegistration);
        }

        cookie = AllocateFallbackToken();
        var activation = new FallbackActivation(cookie, handle, activeRegistrations);
        _fallbackActivations[cookie] = activation;
        var threadId = NativeMethods.GetCurrentThreadId();
        var stack = _threadFallbackStacks.GetOrAdd(threadId, _ => []);
        stack.Add(activation);
        return true;
    }

    private ActiveClassRegistration? TryRegisterSingleClass(ComClassRegistration registration)
    {
        var moduleHandle = LoadLibrary(registration.DllPath);
        if (moduleHandle == IntPtr.Zero)
        {
            SetLastFallbackError(NormalizeWin32Error(GetLastWin32Error(), NativeMethods.ERROR_MOD_NOT_FOUND));
            return null;
        }

        var procAddress = GetProcAddress(moduleHandle, "DllGetClassObject");
        if (procAddress == IntPtr.Zero)
        {
            SetLastFallbackError(NormalizeWin32Error(GetLastWin32Error(), NativeMethods.ERROR_PROC_NOT_FOUND));
            return null;
        }

        if (!AcquireClassFactory(procAddress, registration.Clsid, out var classFactory))
        {
            return null;
        }

        if (!RegisterFactoryWithRetry(registration.Clsid, classFactory, out var registrationToken))
        {
            ReleaseClassObject(classFactory);
            return null;
        }

        return new ActiveClassRegistration(classFactory, registrationToken);
    }

    private bool AcquireClassFactory(IntPtr procAddress, Guid clsid, out IntPtr classFactory)
    {
        var iid = IidIClassFactory;
        var hr = GetClassObject(procAddress, ref clsid, ref iid, out classFactory);
        if (hr == NativeMethods.E_NOINTERFACE)
        {
            iid = IidIUnknown;
            hr = GetClassObject(procAddress, ref clsid, ref iid, out classFactory);
        }

        if (hr < 0 || classFactory == IntPtr.Zero)
        {
            SetLastFallbackError(hr < 0 ? hr : NativeMethods.CLASS_E_CLASSNOTAVAILABLE);
            classFactory = IntPtr.Zero;
            return false;
        }

        return true;
    }

    private bool RegisterFactoryWithRetry(Guid clsid, IntPtr classFactory, out uint registrationToken)
    {
        var hr = RegisterClassObject(ref clsid, classFactory, out registrationToken);
        if (hr == NativeMethods.CO_E_NOTINITIALIZED)
        {
            var initHr = InitializeMultithreadedApartment();
            if (initHr >= 0)
            {
                hr = RegisterClassObject(ref clsid, classFactory, out registrationToken);
            }
        }

        if (hr < 0)
        {
            SetLastFallbackError(hr);
            return false;
        }

        return true;
    }

    private void UnwindActiveRegistrations(List<ActiveClassRegistration> activeRegistrations)
    {
        for (var i = activeRegistrations.Count - 1; i >= 0; i--)
        {
            var entry = activeRegistrations[i];
            _ = RevokeClassObject(entry.RegistrationToken);
            ReleaseClassObject(entry.ClassFactory);
        }

        activeRegistrations.Clear();
    }

    private bool DeactivateFallbackContext(FallbackActivation activation)
    {
        var threadId = NativeMethods.GetCurrentThreadId();
        if (!_threadFallbackStacks.TryGetValue(threadId, out var stack)
            || stack.Count == 0
            || stack[stack.Count - 1].Cookie != activation.Cookie)
        {
            SetLastFallbackError(NativeMethods.ERROR_INVALID_PARAMETER);
            return false;
        }

        for (var i = activation.ActiveRegistrations.Count - 1; i >= 0; i--)
        {
            var entry = activation.ActiveRegistrations[i];
            var hr = RevokeClassObject(entry.RegistrationToken);
            if (hr < 0)
            {
                SetLastFallbackError(hr);
                return false;
            }

            ReleaseClassObject(entry.ClassFactory);
            activation.ActiveRegistrations.RemoveAt(i);
        }

        stack.RemoveAt(stack.Count - 1);
        if (stack.Count == 0)
        {
            _threadFallbackStacks.TryRemove(threadId, out _);
        }

        _fallbackActivations.TryRemove(activation.Cookie, out _);
        return true;
    }

    private sealed class ComClassRegistration
    {
        internal ComClassRegistration(string dllPath, Guid clsid)
        {
            DllPath = dllPath;
            Clsid = clsid;
        }

        internal string DllPath { get; }

        internal Guid Clsid { get; }
    }

    private sealed class ActiveClassRegistration
    {
        internal ActiveClassRegistration(IntPtr classFactory, uint registrationToken)
        {
            ClassFactory = classFactory;
            RegistrationToken = registrationToken;
        }

        internal IntPtr ClassFactory { get; }

        internal uint RegistrationToken { get; }
    }

    private sealed class FallbackContext
    {
        private int _referenceCount = 1;

        internal FallbackContext(List<ComClassRegistration> registrations)
        {
            Registrations = registrations;
        }

        internal List<ComClassRegistration> Registrations { get; }

        internal void AddReference() => Interlocked.Increment(ref _referenceCount);

        internal int ReleaseReference() => Interlocked.Decrement(ref _referenceCount);
    }

    private sealed class FallbackActivation
    {
        internal FallbackActivation(IntPtr cookie, IntPtr contextHandle, List<ActiveClassRegistration> activeRegistrations)
        {
            Cookie = cookie;
            ContextHandle = contextHandle;
            ActiveRegistrations = activeRegistrations;
        }

        internal IntPtr Cookie { get; }

        internal IntPtr ContextHandle { get; }

        internal List<ActiveClassRegistration> ActiveRegistrations { get; }
    }
}
#pragma warning restore SA1121 // Use built-in type alias
