using System;
using System.Runtime.InteropServices;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SA1307 // Accessible fields should begin with upper-case letter
#pragma warning disable SA1121 // Use built-in type alias
/// <summary>
/// The ACTCTX structure is used by the CreateActCtx function to create the activation context.
/// </summary>
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
internal struct ActCtx
{
    /// <summary>
    /// The size, in bytes, of this structure. This is used to determine the version of this structure.
    /// </summary>
    internal Int32 cbSize;

    /// <summary>
    /// Flags that indicate how the values included in this structure are to be used.
    /// Set any undefined bits in dwFlags to 0. If any undefined bits are not set to 0,
    /// the call to CreateActCtx that creates the activation context fails
    /// and returns an invalid parameter error code.
    /// <list type="table">
    /// <listheader>
    /// <term>Bit flag</term>
    /// <description>Meaning</description>
    /// </listheader>
    /// <item>
    /// <term>ACTCTX_FLAG_PROCESSOR_ARCHITECTURE_VALID</term>
    /// <description>0x001</description>
    /// </item>
    /// <item>
    /// <term>ACTCTX_FLAG_LANGID_VALID</term>
    /// <description>0x002</description>
    /// </item>
    /// <item>
    /// <term>ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID</term>
    /// <description>0x004</description>
    /// </item>
    /// <item>
    /// <term>ACTCTX_FLAG_RESOURCE_NAME_VALID</term>
    /// <description>0x008</description>
    /// </item>
    /// <item>
    /// <term>ACTCTX_FLAG_SET_PROCESS_DEFAULT</term>
    /// <description>0x010</description>
    /// </item>
    /// <item>
    /// <term>ACTCTX_FLAG_APPLICATION_NAME_VALID</term>
    /// <description>0x020</description>
    /// </item>
    /// <item>
    /// <term>ACTCTX_FLAG_HMODULE_VALID</term>
    /// <description>0x080</description>
    /// </item>
    /// </list>
    /// </summary>
    internal UInt32 dwFlags;

    /// <summary>
    /// Null-terminated string specifying the path of the manifest file
    /// or PE image to be used to create the activation context.
    /// If this path refers to an EXE or DLL file, the lpResourceName
    /// member is required.
    /// </summary>
    [MarshalAs(UnmanagedType.LPWStr)]
    internal string lpSource;

    /// <summary>
    /// Identifies the type of processor used. Specifies the system's
    /// processor architecture.
    /// This value can be one of the following values:
    /// PROCESSOR_ARCHITECTURE_INTEL
    /// PROCESSOR_ARCHITECTURE_MIPS
    /// PROCESSOR_ARCHITECTURE_ALPHA
    /// PROCESSOR_ARCHITECTURE_PPC
    /// PROCESSOR_ARCHITECTURE_UNKNOWN.
    /// </summary>
    internal UInt16 wProcessorArchitecture;

    /// <summary>
    /// Specifies the language manifest that should be used. The default
    /// is the current user's current UI language.
    /// If the requested language cannot be found, an approximation
    /// is searched for using the following order:
    /// The current user's specific language. For example, for US English (1033).
    /// The current user's primary language. For example, for English (9).
    /// The current system's specific language.
    /// The current system's primary language.
    /// A nonspecific worldwide language. Language neutral (0).
    /// </summary>
    internal UInt16 wLangId;

    /// <summary>
    /// The base directory in which to perform private assembly probing
    /// if assemblies in the activation context are not present
    /// in the system-wide store.
    /// </summary>
    [MarshalAs(UnmanagedType.LPWStr)]
    internal string lpAssemblyDirectory;

    /// <summary>
    /// Pointer to a null-terminated string that contains the resource name
    /// to be loaded from the PE specified in hModule or lpSource.
    /// If the resource name is an integer, set this member using
    /// MAKEINTRESOURCE. This member is required if lpSource refers
    /// to an EXE or DLL.
    /// </summary>
    [MarshalAs(UnmanagedType.LPWStr)]
    internal string lpResourceName;

    /// <summary>
    /// The name of the current application. If the value of this member
    /// is set to null, the name of the executable that launched
    /// the current process is used.
    /// </summary>
    [MarshalAs(UnmanagedType.LPWStr)]
    internal string lpApplicationName;

    /// <summary>
    /// Use this member rather than lpSource if you have already
    /// loaded a DLL and wish to use it to create activation contexts
    /// rather than using a path in lpSource. See lpResourceName
    /// for the rules of looking up resources in this module.
    /// </summary>
    internal IntPtr hModule;
}
#pragma warning restore SA1121 // Use built-in type alias
#pragma warning restore SA1307 // Accessible fields should begin with upper-case letter
