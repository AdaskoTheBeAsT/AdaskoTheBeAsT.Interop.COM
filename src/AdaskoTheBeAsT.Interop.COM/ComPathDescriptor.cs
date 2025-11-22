using System;

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// Describes the paths to a COM assembly and its associated manifest file.
/// Used for registration-free COM activation.
/// </summary>
public sealed class ComPathDescriptor
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ComPathDescriptor"/> class.
    /// </summary>
    /// <param name="comAssemblyPath">Full path to the COM DLL assembly.</param>
    /// <param name="comManifestPath">Full path to the manifest file.</param>
    /// <exception cref="ArgumentNullException">Thrown when comAssemblyPath or comManifestPath is null.</exception>
    /// <exception cref="ArgumentException">Thrown when comAssemblyPath or comManifestPath is empty or whitespace.</exception>
    public ComPathDescriptor(
        string comAssemblyPath,
        string comManifestPath)
    {
        if (comAssemblyPath is null)
        {
            throw new ArgumentNullException(nameof(comAssemblyPath));
        }

        if (comManifestPath is null)
        {
            throw new ArgumentNullException(nameof(comManifestPath));
        }

        if (string.IsNullOrWhiteSpace(comAssemblyPath))
        {
            throw new ArgumentException("COM assembly path cannot be empty or whitespace.", nameof(comAssemblyPath));
        }

        if (string.IsNullOrWhiteSpace(comManifestPath))
        {
            throw new ArgumentException("COM manifest path cannot be empty or whitespace.", nameof(comManifestPath));
        }

        ComAssemblyPath = comAssemblyPath;
        ComManifestPath = comManifestPath;
    }

    /// <summary>
    /// Gets the full path to the COM DLL assembly.
    /// </summary>
    public string ComAssemblyPath { get; }

    /// <summary>
    /// Gets the full path to the manifest file describing the COM component.
    /// </summary>
    public string ComManifestPath { get; }
}
