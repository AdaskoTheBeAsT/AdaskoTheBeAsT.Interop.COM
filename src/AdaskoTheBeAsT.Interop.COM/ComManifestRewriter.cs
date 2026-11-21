using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SCS0018, SEC0112, SEC0116 // Source and destination manifest paths are explicitly validated before file I/O.
/// <summary>
/// Structured, validated XML rewriter, companion <c>&lt;file&gt;</c> pruner, and <c>&lt;description&gt;</c>/comment
/// sanitizer for Win32 Side-by-Side (SxS) COM manifests staged into per-worker isolated directories.
/// </summary>
public sealed class ComManifestRewriter
{
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    private readonly XDocument _document;
    private readonly string? _sourceManifestPath;

    private ComManifestRewriter(XDocument document, string? sourceManifestPath)
    {
        _document = document;
        _sourceManifestPath = sourceManifestPath;
    }

    /// <summary>
    /// Loads a Win32 SxS manifest XML file from <paramref name="sourceManifestPath"/> with DTD processing prohibited.
    /// </summary>
    /// <param name="sourceManifestPath">Path to the template/source <c>.manifest</c> file.</param>
    /// <returns>A new <see cref="ComManifestRewriter"/> instance guarded against overwriting <paramref name="sourceManifestPath"/> in-place.</returns>
    public static ComManifestRewriter FromFile(string sourceManifestPath)
    {
        var normalizedPath = ValidateSafeFilePath(sourceManifestPath, nameof(sourceManifestPath));
        if (!File.Exists(normalizedPath))
        {
            throw new FileNotFoundException("Source COM manifest file was not found.", normalizedPath);
        }

        var settings = CreateSecureXmlReaderSettings();
        XDocument doc;
        try
        {
            using var stream = File.OpenRead(normalizedPath);
            using var reader = XmlReader.Create(stream, settings);
            doc = XDocument.Load(reader, LoadOptions.None);
        }
        catch (XmlException ex)
        {
            throw new InvalidOperationException("Source COM manifest contains malformed or prohibited XML.", ex);
        }

        EnsureValidAssemblyRoot(doc);
        return new ComManifestRewriter(doc, normalizedPath);
    }

    /// <summary>
    /// Loads a Win32 SxS manifest from an in-memory XML string with DTD processing prohibited.
    /// </summary>
    /// <param name="manifestXml">The source Win32 SxS manifest XML text.</param>
    /// <returns>A new <see cref="ComManifestRewriter"/> instance.</returns>
    public static ComManifestRewriter FromXml(string manifestXml)
    {
        ThrowHelper.ThrowIfNullOrWhiteSpace(manifestXml, nameof(manifestXml));

        var settings = CreateSecureXmlReaderSettings();
        XDocument doc;
        try
        {
            using var stringReader = new StringReader(manifestXml);
            using var reader = XmlReader.Create(stringReader, settings);
            doc = XDocument.Load(reader, LoadOptions.None);
        }
        catch (XmlException ex)
        {
            throw new InvalidOperationException("Manifest XML is malformed or contains prohibited constructs.", ex);
        }

        EnsureValidAssemblyRoot(doc);
        return new ComManifestRewriter(doc, sourceManifestPath: null);
    }

    /// <summary>
    /// Sets (or inserts) the top-level <c>&lt;assemblyIdentity name="..."&gt;</c> attribute value on the manifest.
    /// </summary>
    /// <param name="assemblyName">Safe bare assembly identity name (must not contain path separators, traversal segments, colons, or control characters).</param>
    /// <returns>This <see cref="ComManifestRewriter"/> instance for fluent chaining.</returns>
    public ComManifestRewriter WithAssemblyIdentityName(string assemblyName)
    {
        ValidateSafeBareName(assemblyName, nameof(assemblyName));
        var root = EnsureValidAssemblyRoot(_document);
        var identity = root
            .Elements()
            .FirstOrDefault(static e => string.Equals(e.Name.LocalName, "assemblyIdentity", StringComparison.Ordinal));

        if (identity is not null)
        {
            identity.SetAttributeValue("name", assemblyName);
        }
        else
        {
            var newIdentity = new XElement(
                root.Name.Namespace + "assemblyIdentity",
                new XAttribute("type", "win32"),
                new XAttribute("name", assemblyName),
                new XAttribute("version", "1.0.0.0"));
            root.AddFirst(newIdentity);
        }

        return this;
    }

    /// <summary>
    /// Renames a <c>&lt;file name="..."&gt;</c> entry in the manifest from <paramref name="originalFileName"/> to
    /// <paramref name="replacementFileName"/>, matching bare file names case-insensitively and succeeding idempotently
    /// when <paramref name="replacementFileName"/> is already applied.
    /// </summary>
    /// <param name="originalFileName">Expected bare file name in the template manifest (for example, <c>Vendor.dll</c>).</param>
    /// <param name="replacementFileName">Worker-isolated bare file name (for example, <c>Vendor_00.dll</c>).</param>
    /// <returns>This <see cref="ComManifestRewriter"/> instance for fluent chaining.</returns>
    /// <exception cref="InvalidOperationException">Thrown when neither <paramref name="originalFileName"/> nor <paramref name="replacementFileName"/> exists in the manifest, or when renaming would collide with a separate existing <c>&lt;file&gt;</c> entry.</exception>
    public ComManifestRewriter RenameFile(
        string originalFileName,
        string replacementFileName)
    {
        ValidateSafeBareName(originalFileName, nameof(originalFileName));
        ValidateSafeBareName(replacementFileName, nameof(replacementFileName));

        var root = EnsureValidAssemblyRoot(_document);
        var fileElements = root
            .Elements()
            .Where(static e => string.Equals(e.Name.LocalName, "file", StringComparison.Ordinal))
            .ToList();

        var matchingOriginal = fileElements
            .Where(e => string.Equals((string?)e.Attribute("name"), originalFileName, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var matchingReplacement = fileElements
            .Where(e => string.Equals((string?)e.Attribute("name"), replacementFileName, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (matchingOriginal.Count == 0 && matchingReplacement.Count == 0)
        {
            throw new InvalidOperationException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Manifest does not contain a <file> entry matching '{0}' or '{1}'.",
                    originalFileName,
                    replacementFileName));
        }

        if (matchingOriginal.Count > 0
            && matchingReplacement.Count > 0
            && !string.Equals(originalFileName, replacementFileName, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Cannot rename <file> '{0}' to '{1}' because a separate '{1}' <file> entry already exists in the manifest.",
                    originalFileName,
                    replacementFileName));
        }

        var targets = matchingOriginal.Count > 0 ? matchingOriginal : matchingReplacement;
        foreach (var fileElement in targets)
        {
            fileElement.SetAttributeValue("name", replacementFileName);
        }

        return this;
    }

    /// <summary>
    /// Applies multiple bare <c>&lt;file name="..."&gt;</c> renames after pre-validating every key and value.
    /// </summary>
    /// <param name="fileRenames">Map of original bare file names to worker-isolated replacement bare file names.</param>
    /// <returns>This <see cref="ComManifestRewriter"/> instance for fluent chaining.</returns>
    public ComManifestRewriter RenameFiles(
        IReadOnlyDictionary<string, string> fileRenames)
    {
        ThrowHelper.ThrowIfNull(fileRenames, nameof(fileRenames));
        foreach (var pair in fileRenames)
        {
            ValidateSafeBareName(pair.Key, nameof(fileRenames));
            ValidateSafeBareName(pair.Value, nameof(fileRenames));
        }

        foreach (var pair in fileRenames)
        {
            _ = RenameFile(pair.Key, pair.Value);
        }

        return this;
    }

    /// <summary>
    /// Removes all <c>&lt;file name="..."&gt;</c> elements matching <paramref name="fileName"/> case-insensitively and,
    /// when <paramref name="removeOrphanExternalProxyStubs"/> is <see langword="true"/>, removes any
    /// <c>&lt;comInterfaceExternalProxyStub&gt;</c> elements whose <c>tlbid</c> was owned exclusively by the removed file.
    /// </summary>
    /// <param name="fileName">Bare file name of the companion DLL entry to prune.</param>
    /// <param name="removeOrphanExternalProxyStubs">Whether to prune orphaned <c>&lt;comInterfaceExternalProxyStub&gt;</c> entries referencing removed <c>tlbid</c> GUIDs.</param>
    /// <returns>This <see cref="ComManifestRewriter"/> instance for fluent chaining.</returns>
    public ComManifestRewriter RemoveFile(
        string fileName,
        bool removeOrphanExternalProxyStubs = true)
    {
        ValidateSafeBareName(fileName, nameof(fileName));

        var root = EnsureValidAssemblyRoot(_document);
        var matchingFiles = root
            .Elements()
            .Where(e => string.Equals(e.Name.LocalName, "file", StringComparison.Ordinal)
                && string.Equals((string?)e.Attribute("name"), fileName, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (matchingFiles.Count == 0)
        {
            return this;
        }

        var removedTlbIds = new HashSet<Guid>();
        if (removeOrphanExternalProxyStubs)
        {
            foreach (var removedFile in matchingFiles)
            {
                CollectTypeLibIds(removedFile, removedTlbIds);
            }
        }

        foreach (var removedFile in matchingFiles)
        {
            removedFile.Remove();
        }

        if (removeOrphanExternalProxyStubs && removedTlbIds.Count > 0)
        {
            PruneOrphanExternalProxyStubs(root, removedTlbIds);
        }

        return this;
    }

    /// <summary>
    /// Removes all <c>&lt;description&gt;</c> elements and XML comments from the manifest so generator/licensee PII
    /// never lands in staged worker directories.
    /// </summary>
    /// <returns>This <see cref="ComManifestRewriter"/> instance for fluent chaining.</returns>
    public ComManifestRewriter StripDescriptions()
    {
        var descriptions = _document
            .Descendants()
            .Where(static e => string.Equals(e.Name.LocalName, "description", StringComparison.Ordinal))
            .ToList();
        foreach (var desc in descriptions)
        {
            desc.Remove();
        }

        var comments = _document
            .DescendantNodes()
            .OfType<XComment>()
            .ToList();
        foreach (var comment in comments)
        {
            comment.Remove();
        }

        return this;
    }

    /// <summary>
    /// Validates all structural manifest invariants and serializes the rewritten manifest as a UTF-8 XML string.
    /// </summary>
    /// <returns>The validated Win32 SxS manifest XML text.</returns>
    public string ToXmlString()
        => Utf8NoBom.GetString(ToUtf8Bytes());

    /// <summary>
    /// Validates all structural manifest invariants via the library's strict manifest parser and returns the UTF-8
    /// (without BOM) byte representation of the rewritten manifest.
    /// </summary>
    /// <returns>Validated UTF-8 bytes without a byte-order mark.</returns>
    public byte[] ToUtf8Bytes()
        => SerializeAndValidate(validationDirectory: null);

    /// <summary>
    /// Validates the rewritten manifest, refuses in-place overwrites of the source manifest path, compares against any
    /// existing file at <paramref name="destinationManifestPath"/>, and atomically writes UTF-8 (without BOM) via a
    /// sibling temporary file only when the content differs.
    /// </summary>
    /// <param name="destinationManifestPath">Destination path for the staged worker <c>.manifest</c> file.</param>
    /// <returns><see langword="true"/> if <paramref name="destinationManifestPath"/> was created or updated; <see langword="false"/> if an identical file was already present.</returns>
    public bool SaveIfChanged(string destinationManifestPath)
    {
        var normalizedDestination = ValidateSafeFilePath(destinationManifestPath, nameof(destinationManifestPath));
        if (_sourceManifestPath is not null
            && string.Equals(normalizedDestination, _sourceManifestPath, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Refusing to overwrite source manifest '{0}' in-place. Write staged worker manifests to an isolated destination path.",
                    normalizedDestination));
        }

        var destinationDirectory = Path.GetDirectoryName(normalizedDestination);
        if (destinationDirectory is null || string.IsNullOrWhiteSpace(destinationDirectory))
        {
            throw new ArgumentException("Destination manifest path must include a parent directory.", nameof(destinationManifestPath));
        }

        var utf8Bytes = SerializeAndValidate(destinationDirectory);
        if (File.Exists(normalizedDestination))
        {
            var existingBytes = File.ReadAllBytes(normalizedDestination);
            if (BytesEqual(existingBytes, utf8Bytes))
            {
                return false;
            }
        }

        Directory.CreateDirectory(destinationDirectory);
        WriteFileAtomically(destinationDirectory, normalizedDestination, utf8Bytes);
        return true;
    }

    private static XmlReaderSettings CreateSecureXmlReaderSettings()
        => new()
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        };

    private static XElement EnsureValidAssemblyRoot(XDocument document)
    {
        var root = document.Root;
        if (root is null || !string.Equals(root.Name.LocalName, "assembly", StringComparison.Ordinal))
        {
            throw new InvalidOperationException("COM manifest XML must have a root <assembly> element.");
        }

        return root;
    }

    private static string ValidateSafeFilePath(string path, string paramName)
    {
        ThrowHelper.ThrowIfNullOrWhiteSpace(path, paramName);
        if (ComManifestParser.ContainsUnsafePathCharacters(path, allowDriveColon: true)
            || ComManifestParser.ContainsParentTraversalSegment(path))
        {
            throw new ArgumentException("Manifest file path contains unsafe characters or parent traversal segments.", paramName);
        }

        string normalized;
        try
        {
            normalized = Path.GetFullPath(path);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            throw new ArgumentException("Manifest file path is invalid.", paramName, ex);
        }

        if (ComManifestParser.IsUncOrDevicePath(normalized))
        {
            throw new ArgumentException("UNC and Win32 device paths are not permitted for manifest files.", paramName);
        }

        return normalized;
    }

    private static void ValidateSafeBareName(string value, string paramName)
    {
        ThrowHelper.ThrowIfNullOrWhiteSpace(value, paramName);

        if (string.Equals(value, ".", StringComparison.Ordinal)
            || string.Equals(value, "..", StringComparison.Ordinal)
            || value.IndexOf("..", StringComparison.Ordinal) >= 0
            || ComManifestParser.ContainsUnsafePathCharacters(value, allowDriveColon: false)
            || Path.IsPathRooted(value))
        {
            throw new ArgumentException(
                string.Format(CultureInfo.InvariantCulture, "Value '{0}' must be a bare name without path traversal or rooted segments.", value),
                paramName);
        }

        var invalidFileNameChars = Path.GetInvalidFileNameChars();
        foreach (var ch in value)
        {
            if (ch == Path.DirectorySeparatorChar
                || ch == Path.AltDirectorySeparatorChar
                || ch == ':')
            {
                throw new ArgumentException(
                    string.Format(CultureInfo.InvariantCulture, "Value '{0}' must not contain directory or volume separators.", value),
                    paramName);
            }

            if (Array.IndexOf(invalidFileNameChars, ch) >= 0)
            {
                throw new ArgumentException(
                    string.Format(CultureInfo.InvariantCulture, "Value '{0}' contains an invalid file-name character.", value),
                    paramName);
            }
        }
    }

    private static void CollectTypeLibIds(XElement container, HashSet<Guid> targetSet)
    {
        foreach (var typelib in container.Descendants().Where(static e => string.Equals(e.Name.LocalName, "typelib", StringComparison.Ordinal)))
        {
            var rawTlbId = (string?)typelib.Attribute("tlbid");
            if (rawTlbId is not null && Guid.TryParse(rawTlbId, out var parsedGuid) && parsedGuid != Guid.Empty)
            {
                _ = targetSet.Add(parsedGuid);
            }
        }
    }

    private static void PruneOrphanExternalProxyStubs(XElement root, HashSet<Guid> removedTlbIds)
    {
        var remainingTlbIds = new HashSet<Guid>();
        CollectTypeLibIds(root, remainingTlbIds);

        var orphanStubs = root
            .Elements()
            .Where(e =>
            {
                if (!string.Equals(e.Name.LocalName, "comInterfaceExternalProxyStub", StringComparison.Ordinal))
                {
                    return false;
                }

                var rawTlbId = (string?)e.Attribute("tlbid");
                return rawTlbId is not null
                    && Guid.TryParse(rawTlbId, out var stubTlbId)
                    && removedTlbIds.Contains(stubTlbId)
                    && !remainingTlbIds.Contains(stubTlbId);
            })
            .ToList();

        foreach (var stub in orphanStubs)
        {
            stub.Remove();
        }
    }

    private static bool BytesEqual(byte[] left, byte[] right)
    {
        if (left.Length != right.Length)
        {
            return false;
        }

        for (var i = 0; i < left.Length; i++)
        {
            if (left[i] != right[i])
            {
                return false;
            }
        }

        return true;
    }

    private static void WriteFileAtomically(string destinationDirectory, string normalizedDestination, byte[] utf8Bytes)
    {
        var tempFileName = "." + Path.GetFileName(normalizedDestination) + "." + Guid.NewGuid().ToString("N") + ".tmp";
        var tempFilePath = Path.Combine(destinationDirectory, tempFileName);
        try
        {
            using (var fileStream = new FileStream(tempFilePath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                fileStream.Write(utf8Bytes, 0, utf8Bytes.Length);
                fileStream.Flush(flushToDisk: true);
            }

            if (File.Exists(normalizedDestination))
            {
                File.Replace(tempFilePath, normalizedDestination, destinationBackupFileName: null, ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(tempFilePath, normalizedDestination);
            }
        }
        finally
        {
            if (File.Exists(tempFilePath))
            {
                try
                {
                    File.Delete(tempFilePath);
                }
                catch (IOException ex)
                {
                    _ = ex;
                }
            }
        }
    }

    private byte[] SerializeAndValidate(string? validationDirectory)
    {
        _ = EnsureValidAssemblyRoot(_document);

        var standalone = _document.Declaration?.Standalone;
        if (string.IsNullOrWhiteSpace(standalone))
        {
            standalone = "yes";
        }

        _document.Declaration = new XDeclaration("1.0", "UTF-8", standalone);

        byte[] utf8Bytes;
        using (var memoryStream = new MemoryStream())
        {
            var writerSettings = new XmlWriterSettings
            {
                Encoding = Utf8NoBom,
                Indent = true,
                IndentChars = "  ",
                NewLineChars = "\r\n",
                NewLineHandling = NewLineHandling.Replace,
                OmitXmlDeclaration = false,
            };

            using (var xmlWriter = XmlWriter.Create(memoryStream, writerSettings))
            {
                _document.Save(xmlWriter);
            }

            utf8Bytes = memoryStream.ToArray();
        }

        var xmlString = Utf8NoBom.GetString(utf8Bytes);
        if (!ComManifestParser.TryValidateXmlContent(xmlString, _sourceManifestPath, validationDirectory, out var errorCode))
        {
            throw new InvalidOperationException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Rewritten COM manifest failed structural validation (Win32 error {0}). Ensure at least one valid <comClass> remains and all <typelib>, ProgID, and proxy/stub attributes are well-formed.",
                    errorCode));
        }

        return utf8Bytes;
    }
}
#pragma warning restore SCS0018, SEC0112, SEC0116
