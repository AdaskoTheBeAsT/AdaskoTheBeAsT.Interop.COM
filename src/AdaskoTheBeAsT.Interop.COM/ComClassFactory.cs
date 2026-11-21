using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1121 // Use built-in type alias
#pragma warning disable SYSLIB1096 // Direct COM RCW wrapping is the core purpose of ComClassFactory.
#pragma warning disable SCS0018, SEC0116 // Caller-supplied COM server DLL paths are intentional.
/// <summary>
/// Instantiates COM coclasses directly from a specific loaded native module (<c>HMODULE</c>) via
/// <c>DllGetClassObject</c> and <c>IClassFactory::CreateInstance</c>, completely bypassing global
/// <c>CoRegisterClassObject</c> tables, WinSxS activation stacks, and <c>HKCR</c>/<c>HKLM</c> registry lookups.
/// </summary>
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
public static class ComClassFactory
{
#pragma warning disable MA0176 // Optimize guid creation
    private static readonly Guid IidIUnknown = new("00000000-0000-0000-C000-000000000046");
    private static readonly Guid IidIClassFactory = new("00000001-0000-0000-C000-000000000046");
#pragma warning restore MA0176

    /// <summary>
    /// Creates an instance of <paramref name="clsid"/> directly from <paramref name="moduleHandle"/>
    /// without consulting global COM class tables or the Windows Registry.
    /// </summary>
    /// <param name="moduleHandle">Native <c>HMODULE</c> of the loaded COM server DLL.</param>
    /// <param name="clsid">Coclass identifier (<c>CLSID</c>) exported by <paramref name="moduleHandle"/>.</param>
    /// <param name="interfaceId">Optional native interface <c>IID</c> to request from <c>IClassFactory::CreateInstance</c> (defaults to <c>IID_IUnknown</c>).</param>
    /// <returns>A <see cref="ComCreateResult{T}"/> containing the created RCW on success.</returns>
    public static ComCreateResult<object> CreateInstance(
        IntPtr moduleHandle,
        Guid clsid,
        Guid? interfaceId = null)
    {
        if (moduleHandle == IntPtr.Zero)
        {
            return Failure<object>(new ArgumentException("Module handle cannot be zero.", nameof(moduleHandle)));
        }

        if (clsid == Guid.Empty)
        {
            return Failure<object>(new ArgumentException("CLSID cannot be Guid.Empty.", nameof(clsid)));
        }

        var targetIid = interfaceId ?? IidIUnknown;
        if (targetIid == Guid.Empty)
        {
            return Failure<object>(new ArgumentException("Interface ID cannot be Guid.Empty.", nameof(interfaceId)));
        }

        return CreateInstanceCore<object>(moduleHandle, clsid, targetIid);
    }

    /// <summary>
    /// Creates a strongly typed instance of <paramref name="clsid"/> directly from <paramref name="moduleHandle"/>
    /// without consulting global COM class tables or the Windows Registry.
    /// </summary>
    /// <typeparam name="TInterface">Managed COM interface or coclass wrapper type.</typeparam>
    /// <param name="moduleHandle">Native <c>HMODULE</c> of the loaded COM server DLL.</param>
    /// <param name="clsid">Coclass identifier (<c>CLSID</c>) exported by <paramref name="moduleHandle"/>.</param>
    /// <returns>A <see cref="ComCreateResult{T}"/> containing the typed RCW on success.</returns>
    public static ComCreateResult<TInterface> CreateInstance<TInterface>(
        IntPtr moduleHandle,
        Guid clsid)
        where TInterface : class
    {
        if (moduleHandle == IntPtr.Zero)
        {
            return Failure<TInterface>(new ArgumentException("Module handle cannot be zero.", nameof(moduleHandle)));
        }

        if (clsid == Guid.Empty)
        {
            return Failure<TInterface>(new ArgumentException("CLSID cannot be Guid.Empty.", nameof(clsid)));
        }

        var interfaceType = typeof(TInterface);
        var targetIid = interfaceType.IsInterface && interfaceType.GUID != Guid.Empty
            ? interfaceType.GUID
            : IidIUnknown;

        return CreateInstanceCore<TInterface>(moduleHandle, clsid, targetIid);
    }

    /// <summary>
    /// Loads <paramref name="dllPath"/> (retaining the returned <paramref name="moduleHandle"/> on success)
    /// and instantiates <paramref name="clsid"/> directly via its exported <c>DllGetClassObject</c>.
    /// </summary>
    /// <remarks>
    /// When creation succeeds, the caller owns the returned <paramref name="moduleHandle"/> load reference and must keep
    /// the module loaded for as long as any native objects from it remain alive. When creation fails, the temporary
    /// load reference is freed automatically and <paramref name="moduleHandle"/> is set to <see cref="IntPtr.Zero"/>.
    /// </remarks>
    /// <param name="dllPath">Path to the COM server DLL.</param>
    /// <param name="clsid">Coclass identifier (<c>CLSID</c>) to instantiate.</param>
    /// <param name="moduleHandle">Receives the loaded <c>HMODULE</c> when creation succeeds; otherwise <see cref="IntPtr.Zero"/>.</param>
    /// <param name="interfaceId">Optional native interface <c>IID</c> to request from <c>IClassFactory::CreateInstance</c>.</param>
    /// <returns>A <see cref="ComCreateResult{T}"/> containing the created RCW on success.</returns>
    public static ComCreateResult<object> CreateInstanceFromDll(
        string dllPath,
        Guid clsid,
        out IntPtr moduleHandle,
        Guid? interfaceId = null)
    {
        ThrowHelper.ThrowIfNull(dllPath, nameof(dllPath));
        if (string.IsNullOrWhiteSpace(dllPath))
        {
            throw new ArgumentException("COM DLL path cannot be empty or whitespace.", nameof(dllPath));
        }

        moduleHandle = IntPtr.Zero;
        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(dllPath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return Failure<object>(ex);
        }

        var loadedHandle = ActivationContextApi.Instance.LoadLibraryCore(fullPath);
        if (loadedHandle == IntPtr.Zero)
        {
            var lastError = Marshal.GetLastWin32Error();
            return Failure<object>(new Win32Exception(lastError != 0 ? lastError : NativeMethods.ERROR_MOD_NOT_FOUND));
        }

        var creation = CreateInstance(loadedHandle, clsid, interfaceId);
        if (!creation.Success)
        {
            _ = NativeMethods.FreeLibrary(loadedHandle);
            NanoServerOle2Shim.OnModuleUnloaded(loadedHandle);
            return creation;
        }

        moduleHandle = loadedHandle;
        return creation;
    }

    /// <summary>
    /// Loads <paramref name="dllPath"/> (retaining the returned <paramref name="moduleHandle"/> on success)
    /// and instantiates a strongly typed <typeparamref name="TInterface"/> wrapper for <paramref name="clsid"/>.
    /// </summary>
    /// <typeparam name="TInterface">Managed COM interface or coclass wrapper type.</typeparam>
    /// <param name="dllPath">Path to the COM server DLL.</param>
    /// <param name="clsid">Coclass identifier (<c>CLSID</c>) to instantiate.</param>
    /// <param name="moduleHandle">Receives the loaded <c>HMODULE</c> when creation succeeds; otherwise <see cref="IntPtr.Zero"/>.</param>
    /// <returns>A <see cref="ComCreateResult{T}"/> containing the typed RCW on success.</returns>
    public static ComCreateResult<TInterface> CreateInstanceFromDll<TInterface>(
        string dllPath,
        Guid clsid,
        out IntPtr moduleHandle)
        where TInterface : class
    {
        ThrowHelper.ThrowIfNull(dllPath, nameof(dllPath));
        if (string.IsNullOrWhiteSpace(dllPath))
        {
            throw new ArgumentException("COM DLL path cannot be empty or whitespace.", nameof(dllPath));
        }

        moduleHandle = IntPtr.Zero;
        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(dllPath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return Failure<TInterface>(ex);
        }

        var loadedHandle = ActivationContextApi.Instance.LoadLibraryCore(fullPath);
        if (loadedHandle == IntPtr.Zero)
        {
            var lastError = Marshal.GetLastWin32Error();
            return Failure<TInterface>(new Win32Exception(lastError != 0 ? lastError : NativeMethods.ERROR_MOD_NOT_FOUND));
        }

        var creation = CreateInstance<TInterface>(loadedHandle, clsid);
        if (!creation.Success)
        {
            _ = NativeMethods.FreeLibrary(loadedHandle);
            NanoServerOle2Shim.OnModuleUnloaded(loadedHandle);
            return creation;
        }

        moduleHandle = loadedHandle;
        return creation;
    }

    private static ComCreateResult<TResult> CreateInstanceCore<TResult>(
        IntPtr moduleHandle,
        Guid clsid,
        Guid targetIid)
        where TResult : class
    {
        try
        {
            NanoServerOptions.EnsureCurrentStaThreadInitialized();
            var procAddress = NativeMethods.GetProcAddress(moduleHandle, "DllGetClassObject");
            if (procAddress == IntPtr.Zero)
            {
                var lastError = Marshal.GetLastWin32Error();
                return Failure<TResult>(new Win32Exception(lastError != 0 ? lastError : NativeMethods.ERROR_PROC_NOT_FOUND));
            }

            if (!TryAcquireClassFactory(procAddress, clsid, out var classFactoryPtr, out var factoryError))
            {
                return Failure<TResult>(factoryError!);
            }

            try
            {
                return InvokeClassFactoryCreateInstance<TResult>(classFactoryPtr, targetIid);
            }
            finally
            {
                _ = Marshal.Release(classFactoryPtr);
            }
        }
        catch (Exception ex)
        {
            return Failure<TResult>(ex);
        }
    }

    private static bool TryAcquireClassFactory(
        IntPtr procAddress,
        Guid clsid,
        out IntPtr classFactoryPtr,
        out Exception? failure)
    {
        failure = null;
        var dllGetClassObject = Marshal.GetDelegateForFunctionPointer<NativeMethods.DllGetClassObjectDelegate>(procAddress);
        var clsidCopy = clsid;
        var iidFactory = IidIClassFactory;
        var hr = dllGetClassObject(ref clsidCopy, ref iidFactory, out classFactoryPtr);
        if (hr == NativeMethods.E_NOINTERFACE)
        {
            var iidUnknown = IidIUnknown;
            hr = dllGetClassObject(ref clsidCopy, ref iidUnknown, out var unkPtr);
            if (hr >= 0 && unkPtr != IntPtr.Zero)
            {
                try
                {
                    hr = NativeMethods.QueryInterface(unkPtr, ref iidFactory, out classFactoryPtr);
                }
                finally
                {
                    _ = Marshal.Release(unkPtr);
                }
            }
        }

        if (hr < 0 || classFactoryPtr == IntPtr.Zero)
        {
            classFactoryPtr = IntPtr.Zero;
            failure = ComErrorDiagnostics.CreateExceptionForHr(hr < 0 ? hr : NativeMethods.CLASS_E_CLASSNOTAVAILABLE);
            return false;
        }

        return true;
    }

    private static ComCreateResult<TResult> InvokeClassFactoryCreateInstance<TResult>(
        IntPtr classFactoryPtr,
        Guid targetIid)
        where TResult : class
    {
        var vtable = Marshal.ReadIntPtr(classFactoryPtr);
        if (vtable == IntPtr.Zero)
        {
            return Failure<TResult>(ComErrorDiagnostics.CreateExceptionForHr(NativeMethods.E_POINTER));
        }

        var createInstancePtr = Marshal.ReadIntPtr(vtable, 3 * IntPtr.Size);
        if (createInstancePtr == IntPtr.Zero)
        {
            return Failure<TResult>(ComErrorDiagnostics.CreateExceptionForHr(NativeMethods.E_POINTER));
        }

        var createInstance = Marshal.GetDelegateForFunctionPointer<NativeMethods.ClassFactoryCreateInstanceDelegate>(createInstancePtr);
        var iidCopy = targetIid;
        var hr = createInstance(classFactoryPtr, IntPtr.Zero, ref iidCopy, out var instancePtr);
        if (hr < 0 || instancePtr == IntPtr.Zero)
        {
            return Failure<TResult>(ComErrorDiagnostics.CreateExceptionForHr(hr < 0 ? hr : NativeMethods.E_NOINTERFACE));
        }

        try
        {
            return WrapNativeInstance<TResult>(instancePtr);
        }
        finally
        {
            _ = Marshal.Release(instancePtr);
        }
    }

    private static ComCreateResult<TResult> WrapNativeInstance<TResult>(IntPtr instancePtr)
        where TResult : class
    {
        var rcw = Marshal.GetObjectForIUnknown(instancePtr);
        if (rcw is TResult directTyped)
        {
            return new ComCreateResult<TResult>
            {
                Success = true,
                Value = directTyped,
            };
        }

        if (Marshal.IsComObject(rcw))
        {
            _ = Marshal.FinalReleaseComObject(rcw);
        }

        var typedRcw = Marshal.GetTypedObjectForIUnknown(instancePtr, typeof(TResult));
        if (typedRcw is TResult castTyped)
        {
            return new ComCreateResult<TResult>
            {
                Success = true,
                Value = castTyped,
            };
        }

        if (typedRcw is not null && Marshal.IsComObject(typedRcw))
        {
            _ = Marshal.FinalReleaseComObject(typedRcw);
        }

        return Failure<TResult>(new InvalidCastException(
            "The created COM object could not be cast to '" + typeof(TResult).FullName + "'."));
    }

    private static ComCreateResult<TResult> Failure<TResult>(Exception exception)
        where TResult : class
        => new()
        {
            Success = false,
            Exception = exception,
        };
}
#pragma warning restore SCS0018, SEC0116
#pragma warning restore SYSLIB1096
#pragma warning restore SA1121
