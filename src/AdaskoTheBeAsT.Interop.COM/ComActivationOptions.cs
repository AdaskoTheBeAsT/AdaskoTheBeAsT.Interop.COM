namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Options controlling registration-free COM activation, module lifetime management,
/// message pumping, and Windows Nano Server fallback behavior.
/// </summary>
public sealed class ComActivationOptions
{
    /// <summary>
    /// Gets or sets a value indicating whether pending Windows messages are pumped on the calling thread
    /// after activation callbacks and during handle or scope release.
    /// Defaults to <see langword="true"/>.
    /// </summary>
    public bool PumpPendingMessages { get; set; } = true;

    /// <summary>
    /// Gets or sets a value indicating whether manifest-referenced COM DLLs should be explicitly preloaded
    /// with reference-counted module leases when the scope activates.
    /// Defaults to <see langword="false"/>.
    /// </summary>
    public bool PreloadModules { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether COM server modules leased by the activation scope should be
    /// unloaded (<c>CoFreeUnusedLibrariesEx</c> and <c>FreeLibrary</c>) when the last scope lease ends and
    /// <c>DllCanUnloadNow</c> reports no live native objects.
    /// Defaults to <see langword="false"/>.
    /// </summary>
    public bool UnloadModulesOnDeactivate { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether <c>&lt;typelib&gt;</c> entries in the manifest are loaded into memory
    /// with <c>REGKIND_NONE</c> during Nano Server fallback activation.
    /// Defaults to <see langword="true"/>.
    /// </summary>
    public bool LoadTypeLibsOnFallback { get; set; } = true;

    /// <summary>
    /// Gets or sets a value indicating whether manifest <c>progid</c> entries are registered under
    /// <c>HKCU\Software\Classes</c> during Nano Server fallback activation and rolled back when the last
    /// owning scope deactivates.
    /// Defaults to <see langword="false"/>.
    /// </summary>
    public bool RegisterCurrentUserProgIdsOnFallback { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether manifest <c>&lt;typelib&gt;</c> entries are registered under
    /// <c>HKCU\Software\Classes\TypeLib</c> during Nano Server fallback activation and rolled back when the last
    /// owning scope deactivates.
    /// Defaults to <see langword="false"/>.
    /// </summary>
    public bool RegisterCurrentUserTypeLibsOnFallback { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether known proxy/stub classes (<c>PSFactoryBuffer</c> or classes
    /// referenced by <c>proxyStubClsid32</c> / <c>&lt;comInterfaceProxyStub&gt;</c>) that return
    /// <c>CLASS_E_CLASSNOTAVAILABLE</c> are skipped during Nano Server fallback activation.
    /// Ordinary coclasses returning <c>CLASS_E_CLASSNOTAVAILABLE</c> always fail activation.
    /// Defaults to <see langword="true"/>.
    /// </summary>
    public bool IgnoreUnavailableProxyStubClasses { get; set; } = true;

    /// <summary>
    /// Gets or sets a value indicating whether Win32 Side-by-Side (<c>CreateActCtxW</c>) is bypassed unconditionally
    /// in favor of the manifest-driven in-process class-factory bridge (<c>DllGetClassObject</c> +
    /// <c>CoRegisterClassObject</c>), even on Desktop Windows or Server Core.
    /// Useful when WinSxS <c>HKCR\CLSID</c> virtualization hides OLE DB service keys (<c>OLEDB_SERVICES</c>)
    /// or for deterministic parity testing with Windows Nano Server.
    /// Defaults to <see langword="false"/>.
    /// </summary>
    public bool ForceInProcessFallback { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether opt-in <c>HKCU\Software\Classes</c> writes are skipped for entries
    /// whose expected mapping is already present in <c>HKEY_CLASSES_ROOT</c> (<c>HKCR</c>).
    /// Defaults to <see langword="true"/>.
    /// </summary>
    public bool SkipCurrentUserRegistrationWhenPresentInClassesRoot { get; set; } = true;

    internal static ComActivationOptions Default { get; } = new();

    internal ComActivationOptions Clone()
        => new()
        {
            PumpPendingMessages = PumpPendingMessages,
            PreloadModules = PreloadModules,
            UnloadModulesOnDeactivate = UnloadModulesOnDeactivate,
            LoadTypeLibsOnFallback = LoadTypeLibsOnFallback,
            RegisterCurrentUserProgIdsOnFallback = RegisterCurrentUserProgIdsOnFallback,
            RegisterCurrentUserTypeLibsOnFallback = RegisterCurrentUserTypeLibsOnFallback,
            IgnoreUnavailableProxyStubClasses = IgnoreUnavailableProxyStubClasses,
            ForceInProcessFallback = ForceInProcessFallback,
            SkipCurrentUserRegistrationWhenPresentInClassesRoot = SkipCurrentUserRegistrationWhenPresentInClassesRoot,
        };
}
