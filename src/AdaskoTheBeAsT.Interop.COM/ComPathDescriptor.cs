namespace AdaskoTheBeAsT.Interop.COM;

public class ComPathDescriptor
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
