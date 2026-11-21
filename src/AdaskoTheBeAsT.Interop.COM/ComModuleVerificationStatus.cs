namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Outcome of inspecting a COM object's native vtable module ownership against an expected isolated server module.
/// </summary>
public enum ComModuleVerificationStatus
{
    /// <summary>
    /// The inspected vtable method slot resides inside the expected native COM module.
    /// </summary>
    Match = 0,

    /// <summary>
    /// The inspected vtable method slot resides inside a different non-runtime native COM module
    /// (for example, a globally registered DLL or another worker's copy).
    /// </summary>
    Mismatch = 1,

    /// <summary>
    /// The inspected vtable method slot belongs to a shared COM/OLE marshaling proxy, tear-off helper,
    /// or language runtime module (such as <c>combase.dll</c>, <c>oleaut32.dll</c>, or <c>msvbvm60.dll</c>),
    /// or to dynamically generated memory not backed by a PE module, which cannot by itself prove or disprove
    /// server module isolation.
    /// </summary>
    Inconclusive = 2,
}
