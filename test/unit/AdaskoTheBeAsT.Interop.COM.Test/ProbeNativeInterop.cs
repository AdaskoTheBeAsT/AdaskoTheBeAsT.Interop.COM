using System.Runtime.InteropServices;

namespace AdaskoTheBeAsT.Interop.COM.Test;

#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable SYSLIB1054 // P/Invoke in test assembly targets net472..net11.0.
#pragma warning disable SYSLIB1096 // Runtime COM wrappers intentionally test RCW behavior.
#pragma warning disable SCS0018, SEC0116 // Controlled test temp directory paths.
internal static class ProbeNativeInterop
{
    internal const string AlphaClsidString = "{3a930e11-62b4-41d5-8c0a-510d92a70101}";
    internal const string AlphaIidString = "3a930e11-62b4-41d5-8c0a-510d92a70102";
    internal const string SecondaryIidString = "3a930e11-62b4-41d5-8c0a-510d92a70103";
    internal const string BetaClsidString = "{3a930e11-62b4-41d5-8c0a-510d92a70201}";
    internal const string BetaIidString = "3a930e11-62b4-41d5-8c0a-510d92a70202";
    internal const string ProxyStubClsidString = "{3a930e11-62b4-41d5-8c0a-510d92a70301}";
    internal const string UnavailableClsidString = "{3a930e11-62b4-41d5-8c0a-510d92a70401}";
    internal const string PersistProbeClsidString = "{3a930e11-62b4-41d5-8c0a-510d92a70501}";

#pragma warning disable MA0176 // Optimize guid creation
    internal static readonly Guid AlphaClsid = new("3a930e11-62b4-41d5-8c0a-510d92a70101");
    internal static readonly Guid AlphaIid = new(AlphaIidString);
    internal static readonly Guid SecondaryIid = new(SecondaryIidString);
    internal static readonly Guid BetaClsid = new("3a930e11-62b4-41d5-8c0a-510d92a70201");
    internal static readonly Guid BetaIid = new(BetaIidString);
    internal static readonly Guid ProxyStubClsid = new("3a930e11-62b4-41d5-8c0a-510d92a70301");
    internal static readonly Guid UnavailableClsid = new("3a930e11-62b4-41d5-8c0a-510d92a70401");
#pragma warning restore MA0176

    [ComImport]
    [Guid(AlphaIidString)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IProbeAlpha
    {
        [PreserveSig]
        int ComputeAlpha(int input, out int output);

        [PreserveSig]
        int GetModuleBaseAddress(out IntPtr moduleBase);
    }

    [ComImport]
    [Guid(SecondaryIidString)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IProbeSecondary
    {
        [PreserveSig]
        int GetSignature(out int signature);
    }

    [ComImport]
    [Guid(BetaIidString)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IProbeBeta
    {
        [PreserveSig]
        int CombineWithAlpha([MarshalAs(UnmanagedType.Interface)] IProbeAlpha alpha, int seed, out int output);

        [PreserveSig]
        int GetBetaTag(out int tag);
    }

    internal static string ProbeDllSourcePath
        => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "NativeLifetimeProbe.dll"));

    internal static (string WorkerDirectory, string DllPath, string ManifestPath) CreateIsolatedProbeWorker(
        string workerName,
        string threadingModel = "Apartment",
        string? alphaProgId = null,
        bool includeProxyStubTag = false)
    {
        var workerDir = Path.Combine(
            Path.GetTempPath(),
            "AdaskoComProbe_" + workerName + "_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workerDir);

        var dllPath = Path.Combine(workerDir, "NativeLifetimeProbe.dll");
        File.Copy(ProbeDllSourcePath, dllPath, overwrite: true);

        var progIdAttribute = string.IsNullOrEmpty(alphaProgId) ? string.Empty : $" progid=\"{alphaProgId}\"";
        var proxyStubElement = includeProxyStubTag
            ? $"    <comInterfaceProxyStub iid=\"{{{SecondaryIidString}}}\" name=\"IProbeSecondary\" proxyStubClsid32=\"{ProxyStubClsidString}\" />{Environment.NewLine}"
            : string.Empty;

        var manifestPath = Path.Combine(workerDir, "NativeLifetimeProbe.manifest");
        var manifestXml =
            $"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
              <assemblyIdentity type="win32" name="NativeLifetimeProbe" version="1.0.0.0" />
              <file name="NativeLifetimeProbe.dll">
                <comClass clsid="{AlphaClsidString}" threadingModel="{threadingModel}"{progIdAttribute} description="Licensee: ConfidentialProbeCustomer" />
                <comClass clsid="{BetaClsidString}" threadingModel="{threadingModel}" />
                <comClass clsid="{PersistProbeClsidString}" threadingModel="{threadingModel}" />
            {proxyStubElement}  </file>
            </assembly>
            """;

        File.WriteAllText(manifestPath, manifestXml);
        return (workerDir, dllPath, manifestPath);
    }

    internal static void CleanupIsolatedProbeWorker(string workerDir)
    {
        if (Directory.Exists(workerDir))
        {
            Directory.Delete(workerDir, recursive: true);
        }
    }

    internal static object CreateComInstance(Guid clsid)
    {
        var iid = new Guid("00000000-0000-0000-C000-000000000046");
        var hr = NativeMethods.CoCreateInstance(
            ref clsid,
            IntPtr.Zero,
            NativeMethods.CLSCTX_INPROC_SERVER,
            ref iid,
            out var unk);
        Marshal.ThrowExceptionForHR(hr);
        try
        {
            return Marshal.GetObjectForIUnknown(unk);
        }
        finally
        {
            _ = Marshal.Release(unk);
        }
    }

    internal static void RunWithCrossProcessRegistryGate(Action action)
    {
        using var mutex = new Mutex(initiallyOwned: false, @"Local\AdaskoTheBeAsT.Interop.COM.Test.HkcuRegistryGate");
        var acquired = false;
        try
        {
            try
            {
                acquired = mutex.WaitOne(TimeSpan.FromSeconds(30));
            }
            catch (AbandonedMutexException ex)
            {
                _ = ex;
                acquired = true;
            }

#pragma warning disable CC0031
            action();
#pragma warning restore CC0031
        }
        finally
        {
            if (acquired)
            {
                mutex.ReleaseMutex();
            }
        }
    }

    internal static class SafeNativeMethods
    {
        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int GetActiveComObjectCount();

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern void SetForceBusyUnload(int busy);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern void ResetProbeModuleState();

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int CreateDelphiBorrowedInt32SafeArray(
            [In] int[]? values,
            uint count,
            int lowerBound,
            out IntPtr safeArray);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int InspectAndFreeDelphiSafeArray(
            IntPtr safeArray,
            out ushort features,
            out uint locks);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int CreateOwnedDoubleSafeArray(
            [In] double[]? values,
            uint count,
            int lowerBound,
            out IntPtr safeArray);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        internal static extern int CreateBstrSafeArray(
            [In] IntPtr[]? bstrSourcePtrs,
            uint count,
            int lowerBound,
            out IntPtr safeArray);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        internal static extern int CreateMixedVariantSafeArray(
            int intValue,
            double doubleValue,
            [MarshalAs(UnmanagedType.LPWStr)] string? stringValue,
            int lowerBound,
            out IntPtr safeArray);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int CreateRank2Int32SafeArray(out IntPtr safeArray);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int DestroyNativeSafeArray(IntPtr safeArray);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int CreateNativePersistRunnableProbe(int initialValue, out IntPtr unknownPtr);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int CreateNativeMemoryStream(out IntPtr streamPtr);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int RewindNativeMemoryStream(IntPtr streamPtr);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int ExerciseNativeDispenserFactory(
            IntPtr classFactoryPtr,
            out int createdCount,
            out int destroyedCount,
            out int destroyedStringCount);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int InvokeImportedOleRun(IntPtr unknownPtr);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        internal static extern int PopulateProbeComErrorInfo(
            IntPtr interfaceIdPtr,
            [MarshalAs(UnmanagedType.LPWStr)] string? source,
            [MarshalAs(UnmanagedType.LPWStr)] string? description,
            [MarshalAs(UnmanagedType.LPWStr)] string? helpFile,
            uint helpContext,
            int returnHr);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        internal static extern int ParseDoubleWithThreadLocale(
            [MarshalAs(UnmanagedType.LPWStr)] string input,
            uint dwFlags,
            out double outValue,
            out uint outObservedLcid);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern int FormatDoubleWithThreadLocale(
            double input,
            uint dwFlags,
            out IntPtr outBstr,
            out uint outObservedLcid);

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern void InvokeImportedCoUninitializeTwice();

        [DllImport("NativeLifetimeProbe.dll", ExactSpelling = true)]
        internal static extern void InvokeImportedOleUninitializeTwice();
    }

    internal sealed class SimulatedFallbackActivationContextApi : ActivationContextApi
    {
        private int _registerAttempts;

        internal int StaInitCalls { get; private set; }

        internal int MtaInitCalls { get; private set; }

        internal int UninitCalls { get; private set; }

        internal bool ForceNotInitializedOnFirstRegister { get; set; }

        internal override IntPtr CreateActCtx(ref ActCtx context) => new(-1);

        internal override int GetLastWin32Error() => NativeMethods.ERROR_INVALID_FUNCTION;

        internal override int RegisterClassObject(ref Guid clsid, IntPtr classFactory, out uint registrationToken)
        {
            _registerAttempts++;
            if (ForceNotInitializedOnFirstRegister && _registerAttempts == 1)
            {
                registrationToken = 0;
                return NativeMethods.CO_E_NOTINITIALIZED;
            }

            return base.RegisterClassObject(ref clsid, classFactory, out registrationToken);
        }

        internal override int InitializeSingleThreadedApartment()
        {
            StaInitCalls++;
            return base.InitializeSingleThreadedApartment();
        }

        internal override int InitializeMultithreadedApartment()
        {
            MtaInitCalls++;
            return base.InitializeMultithreadedApartment();
        }

        internal override void UninitializeApartment()
        {
            UninitCalls++;
            base.UninitializeApartment();
        }
    }
}
#pragma warning restore SCS0018, SEC0116
#pragma warning restore SYSLIB1096
#pragma warning restore SYSLIB1054
#pragma warning restore SA1121
