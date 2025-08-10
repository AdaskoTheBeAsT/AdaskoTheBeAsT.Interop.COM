namespace AdaskoTheBeAsT.Interop.COM;

public sealed class ComPathDescriptor
{
    public ComPathDescriptor(
        string comAssemblyPath,
        string comManifestPath)
    {
        ComAssemblyPath = comAssemblyPath;
        ComManifestPath = comManifestPath;
    }

    public string ComAssemblyPath { get; }

    public string ComManifestPath { get; }
}
