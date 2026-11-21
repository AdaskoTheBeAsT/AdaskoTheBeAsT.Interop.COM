using System;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1121 // Use built-in type alias
/// <summary>
/// Detailed diagnostic outcome produced by <see cref="ComModuleDiagnostics"/> when inspecting
/// the native module backing a COM interface vtable.
/// </summary>
public sealed class ComModuleInspectionResult
{
    internal ComModuleInspectionResult(
        ComModuleVerificationStatus status,
        IntPtr moduleHandle,
        string? modulePath,
        (IntPtr Vtable, IntPtr Method, int Index) slot,
        Guid interfaceId,
        string diagnosticMessage)
    {
        Status = status;
        ModuleHandle = moduleHandle;
        ModulePath = modulePath;
        VtableAddress = slot.Vtable;
        MethodSlotAddress = slot.Method;
        MethodSlotIndex = slot.Index;
        InterfaceId = interfaceId;
        DiagnosticMessage = diagnosticMessage;
    }

    /// <summary>
    /// Gets the module verification classification (<see cref="ComModuleVerificationStatus.Match"/>,
    /// <see cref="ComModuleVerificationStatus.Mismatch"/>, or <see cref="ComModuleVerificationStatus.Inconclusive"/>).
    /// </summary>
    public ComModuleVerificationStatus Status { get; }

    /// <summary>
    /// Gets the resolved native <c>HMODULE</c> containing <see cref="MethodSlotAddress"/>, or <see cref="IntPtr.Zero"/>.
    /// </summary>
    public IntPtr ModuleHandle { get; }

    /// <summary>
    /// Gets the full file path of <see cref="ModuleHandle"/> when available.
    /// </summary>
    public string? ModulePath { get; }

    /// <summary>
    /// Gets the raw COM interface vtable address read from the interface pointer.
    /// </summary>
    public IntPtr VtableAddress { get; }

    /// <summary>
    /// Gets the function pointer stored at <see cref="MethodSlotIndex"/> in <see cref="VtableAddress"/>.
    /// </summary>
    public IntPtr MethodSlotAddress { get; }

    /// <summary>
    /// Gets the zero-based vtable slot index inspected.
    /// </summary>
    public int MethodSlotIndex { get; }

    /// <summary>
    /// Gets the COM interface identifier (<c>IID</c>) inspected on the object.
    /// </summary>
    public Guid InterfaceId { get; }

    /// <summary>
    /// Gets a human-readable diagnostic summary describing the inspection result.
    /// </summary>
    public string DiagnosticMessage { get; }
}
#pragma warning restore SA1121
