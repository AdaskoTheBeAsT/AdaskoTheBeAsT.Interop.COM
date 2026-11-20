using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable SCS0018, SEC0116 // Manifest and assembly paths are validated before activation.
internal static class ComManifestParser
{
#pragma warning disable MA0176 // Optimize guid creation
    private static readonly Guid ClsidPsDispatch = new("00020420-0000-0000-C000-000000000046");
    private static readonly Guid ClsidPsOaInterface = new("00020424-0000-0000-C000-000000000046");
#pragma warning restore MA0176

    internal enum ComThreadingModel
    {
        Unspecified = 0,
        Apartment = 1,
        Single = 2,
        Free = 3,
        Both = 4,
        Neutral = 5,
    }

    internal static bool TryParse(
        string manifestPath,
        string? assemblyDirectory,
        string? defaultAssemblyPath,
        out ParsedManifest? parsedManifest,
        out int errorCode)
    {
        parsedManifest = null;
        if (string.IsNullOrWhiteSpace(manifestPath) || ContainsUnsafePathCharacters(manifestPath, allowDriveColon: true))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        string normalizedManifestPath;
        try
        {
            normalizedManifestPath = Path.GetFullPath(manifestPath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        if (IsUncOrDevicePath(normalizedManifestPath))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        if (!File.Exists(normalizedManifestPath))
        {
            errorCode = NativeMethods.ERROR_FILE_NOT_FOUND;
            return false;
        }

        var builder = new ManifestParseState(assemblyDirectory, defaultAssemblyPath);
        try
        {
            if (!ParseManifestFile(normalizedManifestPath, builder, out errorCode))
            {
                return false;
            }

            parsedManifest = builder.Build(out errorCode);
            return parsedManifest is not null;
        }
        catch (Exception ex) when (ex is XmlException or IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            errorCode = ex is FileNotFoundException or DirectoryNotFoundException
                ? NativeMethods.ERROR_FILE_NOT_FOUND
                : NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }
    }

    internal static bool TryValidateXmlContent(
        string manifestXml,
        string? sourceManifestPath,
        string? assemblyDirectory,
        out int errorCode)
    {
        if (string.IsNullOrWhiteSpace(manifestXml))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        var effectiveDirectory = assemblyDirectory;
        if ((effectiveDirectory is null || string.IsNullOrWhiteSpace(effectiveDirectory))
            && sourceManifestPath is not null
            && !string.IsNullOrWhiteSpace(sourceManifestPath))
        {
            effectiveDirectory = Path.GetDirectoryName(sourceManifestPath);
        }

        if (effectiveDirectory is null || string.IsNullOrWhiteSpace(effectiveDirectory))
        {
            effectiveDirectory = AppContext.BaseDirectory;
        }

        var builder = new ManifestParseState(effectiveDirectory, defaultAssemblyPath: null, allowUnresolvedDependencies: true);
        try
        {
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
            };

            using var stringReader = new StringReader(manifestXml);
            using var reader = XmlReader.Create(stringReader, settings);
            if (!ReadManifestNodes(reader, sourceManifestPath ?? string.Empty, builder, out errorCode))
            {
                return false;
            }

            var parsed = builder.Build(out errorCode);
            if (parsed is null)
            {
                return false;
            }

            if (parsed.Classes.Count == 0 && parsed.TypeLibs.Count == 0)
            {
                errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
                return false;
            }

            return true;
        }
        catch (Exception ex) when (ex is XmlException or IOException or ArgumentException or NotSupportedException)
        {
            _ = ex;
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }
    }

    internal static bool IsSafeProgId(string? progId)
    {
        if (progId is null || string.IsNullOrWhiteSpace(progId))
        {
            return false;
        }

        if (progId.Length > 255 || progId[0] == '.' || progId[progId.Length - 1] == '.')
        {
            return false;
        }

        var previousWasDot = false;
        foreach (var ch in progId)
        {
            if (ch == '.')
            {
                if (previousWasDot)
                {
                    return false;
                }

                previousWasDot = true;
                continue;
            }

            previousWasDot = false;
            var isValidChar = (ch >= 'A' && ch <= 'Z')
                || (ch >= 'a' && ch <= 'z')
                || (ch >= '0' && ch <= '9')
                || ch == '_'
                || ch == '-';
            if (!isValidChar)
            {
                return false;
            }
        }

        return true;
    }

    internal static bool ContainsUnsafePathCharacters(string path, bool allowDriveColon)
    {
        var invalidChars = Path.GetInvalidPathChars();
        for (var i = 0; i < path.Length; i++)
        {
            var ch = path[i];
            if (char.IsControl(ch) || ch is '*' or '?' or '"' or '<' or '>' or '|')
            {
                return true;
            }

            if (ch == ':' && (!allowDriveColon || i != 1 || !IsAsciiDriveLetter(path[0])))
            {
                return true;
            }

            for (var j = 0; j < invalidChars.Length; j++)
            {
                if (ch == invalidChars[j])
                {
                    return true;
                }
            }
        }

        return false;
    }

    internal static bool IsUncOrDevicePath(string path)
        => path.Length >= 2
            && (path[0] == Path.DirectorySeparatorChar || path[0] == Path.AltDirectorySeparatorChar)
            && (path[1] == Path.DirectorySeparatorChar || path[1] == Path.AltDirectorySeparatorChar);

    internal static bool ContainsParentTraversalSegment(string path)
    {
        var segments = path.Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar], StringSplitOptions.RemoveEmptyEntries);
        for (var i = 0; i < segments.Length; i++)
        {
            if (string.Equals(segments[i], "..", StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static bool ParseManifestFile(
        string manifestPath,
        ManifestParseState state,
        out int errorCode)
    {
        errorCode = 0;
        if (!state.VisitedManifests.Add(manifestPath))
        {
            return true;
        }

        if (!File.Exists(manifestPath))
        {
            errorCode = NativeMethods.ERROR_FILE_NOT_FOUND;
            return false;
        }

        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        };

        using var stream = File.OpenRead(manifestPath);
        using var reader = XmlReader.Create(stream, settings);
        return ReadManifestNodes(reader, manifestPath, state, out errorCode);
    }

    private static bool ReadManifestNodes(
        XmlReader reader,
        string manifestPath,
        ManifestParseState state,
        out int errorCode)
    {
        errorCode = 0;
        var scope = new ElementScope();
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.SignificantWhitespace)
            {
                if (scope.ReadingProgIdElement && !scope.TryAppendProgIdText(reader.Value))
                {
                    errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
                    return false;
                }

                continue;
            }

            if (reader.NodeType == XmlNodeType.EndElement)
            {
                if (!scope.OnEndElement(reader, out errorCode))
                {
                    return false;
                }

                continue;
            }

            if (reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            if (!scope.HasAssemblyRoot)
            {
                if (!string.Equals(reader.LocalName, "assembly", StringComparison.Ordinal))
                {
                    errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
                    return false;
                }

                scope.HasAssemblyRoot = true;
                continue;
            }

            if (!ProcessElement(reader, manifestPath, state, scope, out errorCode))
            {
                return false;
            }
        }

        if (!scope.HasAssemblyRoot)
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        return true;
    }

    private static bool ProcessElement(
        XmlReader reader,
        string manifestPath,
        ManifestParseState state,
        ElementScope scope,
        out int errorCode)
    {
        errorCode = 0;
        return reader.LocalName switch
        {
            "file" => ProcessFileElement(reader, state, scope, out errorCode),
            "dependentAssembly" => ProcessDependentAssemblyElement(reader, scope),
            "comClass" => ProcessComClassElement(reader, state, scope, out errorCode),
            "progid" => ProcessProgIdChildElement(reader, scope, out errorCode),
            "typelib" => ProcessTypeLibElement(reader, state, scope, out errorCode),
            "comInterfaceProxyStub" => ProcessInterfaceProxyStubElement(reader, state, scope, isInternalToFile: true, out errorCode),
            "comInterfaceExternalProxyStub" => ProcessInterfaceProxyStubElement(reader, state, scope, isInternalToFile: false, out errorCode),
            "assemblyIdentity" when scope.DependentDepth >= 0 => ProcessDependentAssemblyIdentity(reader, manifestPath, state, out errorCode),
            _ => true,
        };
    }

    private static bool ProcessFileElement(
        XmlReader reader,
        ManifestParseState state,
        ElementScope scope,
        out int errorCode)
    {
        errorCode = 0;
        var name = reader.GetAttribute("name");
        if (name is not null && (string.IsNullOrWhiteSpace(name) || !TryValidateManifestFileAttribute(name, state.AssemblyDirectory)))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        if (!reader.IsEmptyElement)
        {
            scope.CurrentFileName = name;
            scope.FileDepth = reader.Depth;
        }

        return true;
    }

    private static bool ProcessDependentAssemblyElement(XmlReader reader, ElementScope scope)
    {
        if (!reader.IsEmptyElement)
        {
            scope.DependentDepth = reader.Depth;
        }

        return true;
    }

    private static bool ProcessComClassElement(
        XmlReader reader,
        ManifestParseState state,
        ElementScope scope,
        out int errorCode)
    {
        errorCode = 0;
        if (!TryParseRequiredNonEmptyGuid(reader.GetAttribute("clsid"), out var clsid)
            || !TryParseOptionalNonEmptyGuid(reader.GetAttribute("tlbid"), out var tlbid)
            || !TryReadOptionalSafeProgIdAttribute(reader, "progid", out var progIdAttr)
            || !TryReadOptionalSafeProgIdAttribute(reader, "versionIndependentProgId", out var versionIndependentProgIdAttr))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        var dllPath = ResolveComDllPath(scope.CurrentFileName, state.AssemblyDirectory, state.DefaultAssemblyPath);
        if (dllPath is null)
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        // Check only exact "PSFactoryBuffer" marker; never retain or log arbitrary description/licensee strings.
        var isPsFactoryBuffer = string.Equals(
            reader.GetAttribute("description"),
            "PSFactoryBuffer",
            StringComparison.OrdinalIgnoreCase);

        var threadingModel = NormalizeThreadingModel(reader.GetAttribute("threadingModel"));
        var entry = new MutableComClassEntry(dllPath, clsid, tlbid, threadingModel, isPsFactoryBuffer)
        {
            ExplicitVersionIndependentProgId = versionIndependentProgIdAttr,
        };
        AppendUniqueProgId(entry, progIdAttr);
        AppendUniqueProgId(entry, versionIndependentProgIdAttr);

        if (!state.TryAddComClass(entry))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        if (!reader.IsEmptyElement)
        {
            scope.CurrentComClass = entry;
            scope.ComClassDepth = reader.Depth;
        }

        return true;
    }

    private static bool TryReadOptionalSafeProgIdAttribute(
        XmlReader reader,
        string attributeName,
        out string? progId)
    {
        progId = reader.GetAttribute(attributeName);
        return progId is null || IsSafeProgId(progId);
    }

    private static void AppendUniqueProgId(MutableComClassEntry entry, string? progId)
    {
        if (progId is not null
            && !entry.ProgIds.Exists(existing => string.Equals(existing, progId, StringComparison.OrdinalIgnoreCase)))
        {
            entry.ProgIds.Add(progId);
        }
    }

    private static bool ProcessProgIdChildElement(
        XmlReader reader,
        ElementScope scope,
        out int errorCode)
    {
        errorCode = 0;
        if (scope.CurrentComClass is null)
        {
            return true;
        }

        if (reader.IsEmptyElement)
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        scope.BeginProgIdElement(reader.Depth);
        return true;
    }

    private static bool ProcessTypeLibElement(
        XmlReader reader,
        ManifestParseState state,
        ElementScope scope,
        out int errorCode)
    {
        errorCode = 0;
        if (!TryParseRequiredNonEmptyGuid(reader.GetAttribute("tlbid"), out var tlbid))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        var dllPath = ResolveComDllPath(scope.CurrentFileName, state.AssemblyDirectory, state.DefaultAssemblyPath);
        if (dllPath is null)
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        var rawVersion = reader.GetAttribute("version");
        var version = rawVersion is not null && !string.IsNullOrWhiteSpace(rawVersion)
            ? rawVersion.Trim()
            : "1.0";
        if (!TryParseTypeLibVersion(version, out var majorVersion, out var minorVersion))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        var helpDir = reader.GetAttribute("helpdir");
        string resolvedHelpDir;
        if (helpDir is null || string.IsNullOrWhiteSpace(helpDir))
        {
            resolvedHelpDir = Path.GetDirectoryName(dllPath) ?? string.Empty;
        }
        else if (!TryResolveHelpDirectory(helpDir, Path.GetDirectoryName(dllPath), out resolvedHelpDir))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        var resourceId = reader.GetAttribute("resourceid");
        if (resourceId is not null && !IsSafeResourceIdentifier(resourceId))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        var flagsText = reader.GetAttribute("flags");
        var flagsValue = ParseTypeLibFlags(flagsText);
        state.AddTypeLib(new ManifestTypeLibEntry(dllPath, tlbid, version, majorVersion, minorVersion, resolvedHelpDir, resourceId, flagsValue));
        return true;
    }

    private static bool ProcessInterfaceProxyStubElement(
        XmlReader reader,
        ManifestParseState state,
        ElementScope scope,
        bool isInternalToFile,
        out int errorCode)
    {
        errorCode = 0;
        if (!TryParseRequiredNonEmptyGuid(reader.GetAttribute("iid"), out var iid)
            || !TryParseOptionalNonEmptyGuid(reader.GetAttribute("tlbid"), out var tlbid)
            || !TryParseOptionalNonEmptyGuid(reader.GetAttribute("proxyStubClsid32"), out var proxyStubClsid)
            || !TryParseOptionalNonEmptyGuid(reader.GetAttribute("baseInterface"), out _))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        var effectiveProxyStubClsid = proxyStubClsid ?? (isInternalToFile ? iid : ClsidPsOaInterface);
        if (!IsSystemAutomationProxyStub(effectiveProxyStubClsid))
        {
            state.ReferencedProxyStubClsids.Add(effectiveProxyStubClsid);
            if (isInternalToFile && scope.CurrentFileName is not null)
            {
                var dllPath = ResolveComDllPath(scope.CurrentFileName, state.AssemblyDirectory, state.DefaultAssemblyPath);
                if (dllPath is not null)
                {
                    state.ImplicitProxyStubClasses.Add(new ImplicitProxyStubEntry(dllPath, effectiveProxyStubClsid));
                }
            }
        }

        state.Interfaces.Add(new ManifestInterfaceEntry(iid, tlbid, effectiveProxyStubClsid));
        return true;
    }

    private static bool ProcessDependentAssemblyIdentity(
        XmlReader reader,
        string currentManifestPath,
        ManifestParseState state,
        out int errorCode)
    {
        var dependencyName = reader.GetAttribute("name");
        if (dependencyName is null || string.IsNullOrWhiteSpace(dependencyName) || !IsSafeDependencyName(dependencyName))
        {
            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        var dependentManifestPath = ResolveDependentManifestPath(dependencyName, currentManifestPath, state.AssemblyDirectory);
        if (dependentManifestPath is null)
        {
            var publicKeyToken = reader.GetAttribute("publicKeyToken");
            if (!string.IsNullOrWhiteSpace(publicKeyToken) || state.AllowUnresolvedDependencies)
            {
                errorCode = 0;
                return true;
            }

            errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
            return false;
        }

        return ParseManifestFile(dependentManifestPath, state, out errorCode);
    }

    private static bool IsSystemAutomationProxyStub(Guid clsid)
        => clsid == ClsidPsOaInterface || clsid == ClsidPsDispatch;

    private static ComThreadingModel NormalizeThreadingModel(string? value)
    {
        if (value is null || string.IsNullOrWhiteSpace(value))
        {
            return ComThreadingModel.Unspecified;
        }

        return value.Trim().ToUpperInvariant() switch
        {
            "APARTMENT" => ComThreadingModel.Apartment,
            "SINGLE" => ComThreadingModel.Single,
            "FREE" => ComThreadingModel.Free,
            "BOTH" => ComThreadingModel.Both,
            "NEUTRAL" => ComThreadingModel.Neutral,
            _ => ComThreadingModel.Unspecified,
        };
    }

    private static bool TryParseRequiredNonEmptyGuid(string? value, out Guid guid)
    {
        if (value is null || string.IsNullOrWhiteSpace(value) || !Guid.TryParse(value, out guid) || guid == Guid.Empty)
        {
            guid = Guid.Empty;
            return false;
        }

        return true;
    }

    private static bool TryParseOptionalNonEmptyGuid(string? value, out Guid? guid)
    {
        if (value is null || string.IsNullOrWhiteSpace(value))
        {
            guid = null;
            return true;
        }

        if (!Guid.TryParse(value, out var parsed) || parsed == Guid.Empty)
        {
            guid = null;
            return false;
        }

        guid = parsed;
        return true;
    }

    private static bool TryParseTypeLibVersion(string version, out ushort major, out ushort minor)
    {
        major = 0;
        minor = 0;
        var parts = version.Split('.');
        if (parts.Length is < 1 or > 2)
        {
            return false;
        }

        if (!ushort.TryParse(parts[0], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out major)
            && !ushort.TryParse(parts[0], System.Globalization.NumberStyles.AllowHexSpecifier, System.Globalization.CultureInfo.InvariantCulture, out major))
        {
            return false;
        }

        if (parts.Length == 1)
        {
            return true;
        }

        return ushort.TryParse(parts[1], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out minor)
            || ushort.TryParse(parts[1], System.Globalization.NumberStyles.AllowHexSpecifier, System.Globalization.CultureInfo.InvariantCulture, out minor);
    }

    private static int ParseTypeLibFlags(string? flags)
    {
        if (flags is null || string.IsNullOrWhiteSpace(flags))
        {
            return 0;
        }

        var result = 0;
        var parts = flags.Split([',', '|', ' '], StringSplitOptions.RemoveEmptyEntries);
        for (var i = 0; i < parts.Length; i++)
        {
            result |= parts[i].Trim().ToUpperInvariant() switch
            {
                "RESTRICTED" => 1,
                "CONTROL" => 2,
                "HIDDEN" => 4,
                "HASDISKIMAGE" => 8,
                _ => 0,
            };
        }

        return result;
    }

    private static bool IsSafeResourceIdentifier(string resourceId)
    {
        if (resourceId.Length == 0)
        {
            return true;
        }

        for (var i = 0; i < resourceId.Length; i++)
        {
            if (resourceId[i] is < '0' or > '9')
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryResolveHelpDirectory(string helpDir, string? baseDirectory, out string resolved)
    {
        resolved = string.Empty;
        if (ContainsUnsafePathCharacters(helpDir, allowDriveColon: true)
            || IsUncOrDevicePath(helpDir)
            || ContainsParentTraversalSegment(helpDir))
        {
            return false;
        }

        try
        {
            if (Path.IsPathRooted(helpDir))
            {
                resolved = Path.GetFullPath(helpDir);
                return !IsUncOrDevicePath(resolved);
            }

            if (baseDirectory is null || string.IsNullOrWhiteSpace(baseDirectory))
            {
                return false;
            }

            var normalizedBase = EnsureTrailingSeparator(Path.GetFullPath(baseDirectory));
            var combined = Path.GetFullPath(Path.Combine(normalizedBase, helpDir));
            if (!EnsureTrailingSeparator(combined).StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            resolved = combined;
            return true;
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return false;
        }
    }

    private static bool TryValidateManifestFileAttribute(string fileName, string? assemblyDirectory)
    {
        if (ContainsUnsafePathCharacters(fileName, allowDriveColon: true)
            || IsUncOrDevicePath(fileName)
            || ContainsParentTraversalSegment(fileName))
        {
            return false;
        }

        var endsWithSeparator = fileName[fileName.Length - 1] == Path.DirectorySeparatorChar
            || fileName[fileName.Length - 1] == Path.AltDirectorySeparatorChar;
        if (endsWithSeparator)
        {
            return false;
        }

        try
        {
            if (Path.IsPathRooted(fileName))
            {
                return IsValidRootedManifestFile(fileName);
            }

            if (assemblyDirectory is not null && !string.IsNullOrWhiteSpace(assemblyDirectory))
            {
                var normalizedBase = EnsureTrailingSeparator(Path.GetFullPath(assemblyDirectory));
                var candidate = Path.GetFullPath(Path.Combine(normalizedBase, fileName));
                if (!candidate.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                return !string.IsNullOrWhiteSpace(Path.GetFileName(candidate));
            }

            return !string.IsNullOrWhiteSpace(Path.GetFileName(fileName));
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return false;
        }
    }

    private static bool IsValidRootedManifestFile(string fileName)
    {
        // Reject root-relative ("\evil.dll") and drive-relative ("C:evil.dll") paths.
        if (fileName.Length < 3
            || !IsAsciiDriveLetter(fileName[0])
            || fileName[1] != ':'
            || (fileName[2] != Path.DirectorySeparatorChar && fileName[2] != Path.AltDirectorySeparatorChar))
        {
            return false;
        }

        var rooted = Path.GetFullPath(fileName);
        return !IsUncOrDevicePath(rooted) && !string.IsNullOrWhiteSpace(Path.GetFileName(rooted));
    }

    private static bool IsSafeDependencyName(string dependencyName)
    {
        if (ContainsUnsafePathCharacters(dependencyName, allowDriveColon: false)
            || ContainsParentTraversalSegment(dependencyName)
            || Path.IsPathRooted(dependencyName))
        {
            return false;
        }

        foreach (var ch in dependencyName)
        {
            if (ch == Path.DirectorySeparatorChar || ch == Path.AltDirectorySeparatorChar)
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsAsciiDriveLetter(char ch)
        => (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');

    private static string EnsureTrailingSeparator(string directoryPath)
        => directoryPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            ? directoryPath
            : directoryPath + Path.DirectorySeparatorChar;

    private static string? ResolveComDllPath(
        string? currentFileName,
        string? assemblyDirectory,
        string? defaultAssemblyPath)
    {
        if (currentFileName is not null && !string.IsNullOrWhiteSpace(currentFileName))
        {
            string? candidate = null;
            if (Path.IsPathRooted(currentFileName))
            {
                candidate = Path.GetFullPath(currentFileName);
            }
            else if (assemblyDirectory is not null && !string.IsNullOrWhiteSpace(assemblyDirectory))
            {
                candidate = Path.GetFullPath(Path.Combine(assemblyDirectory, currentFileName));
            }

            if (candidate is not null
                && !string.IsNullOrWhiteSpace(candidate)
                && (File.Exists(candidate) || string.IsNullOrWhiteSpace(defaultAssemblyPath)))
            {
                return candidate;
            }
        }

        return defaultAssemblyPath is null || string.IsNullOrWhiteSpace(defaultAssemblyPath)
            ? null
            : Path.GetFullPath(defaultAssemblyPath);
    }

    private static string? ResolveDependentManifestPath(
        string assemblyName,
        string currentManifestPath,
        string? assemblyDirectory)
    {
        var manifestFileName = assemblyName + ".manifest";
        if (assemblyDirectory is not null && !string.IsNullOrWhiteSpace(assemblyDirectory))
        {
            var directCandidate = Path.GetFullPath(Path.Combine(assemblyDirectory, manifestFileName));
            if (File.Exists(directCandidate))
            {
                return directCandidate;
            }

            var subFolderCandidate = Path.GetFullPath(Path.Combine(assemblyDirectory, assemblyName, manifestFileName));
            if (File.Exists(subFolderCandidate))
            {
                return subFolderCandidate;
            }
        }

        var manifestDirectory = Path.GetDirectoryName(currentManifestPath);
        if (manifestDirectory is not null && !string.IsNullOrWhiteSpace(manifestDirectory))
        {
            var siblingCandidate = Path.GetFullPath(Path.Combine(manifestDirectory, manifestFileName));
            if (File.Exists(siblingCandidate))
            {
                return siblingCandidate;
            }
        }

        return null;
    }

    internal sealed class ParsedManifest
    {
        internal ParsedManifest(
            IReadOnlyList<ManifestComClassEntry> classes,
            IReadOnlyList<ManifestTypeLibEntry> typeLibs,
            IReadOnlyList<ManifestInterfaceEntry> interfaces,
            IReadOnlyList<string> referencedDllPaths)
        {
            Classes = classes;
            TypeLibs = typeLibs;
            Interfaces = interfaces;
            ReferencedDllPaths = referencedDllPaths;
        }

        internal IReadOnlyList<ManifestComClassEntry> Classes { get; }

        internal IReadOnlyList<ManifestTypeLibEntry> TypeLibs { get; }

        internal IReadOnlyList<ManifestInterfaceEntry> Interfaces { get; }

        internal IReadOnlyList<string> ReferencedDllPaths { get; }
    }

    internal sealed class ManifestComClassEntry
    {
        internal ManifestComClassEntry(
            string dllPath,
            Guid clsid,
            Guid? tlbId,
            ComThreadingModel threadingModel,
            bool isKnownProxyStub,
            IReadOnlyList<string> progIds,
            string? primaryProgId,
            string? versionIndependentProgId)
        {
            DllPath = dllPath;
            Clsid = clsid;
            TlbId = tlbId;
            ThreadingModel = threadingModel;
            IsKnownProxyStub = isKnownProxyStub;
            ProgIds = progIds;
            PrimaryProgId = primaryProgId;
            VersionIndependentProgId = versionIndependentProgId;
        }

        internal string DllPath { get; }

        internal Guid Clsid { get; }

        internal Guid? TlbId { get; }

        internal ComThreadingModel ThreadingModel { get; }

        internal bool IsKnownProxyStub { get; }

        internal IReadOnlyList<string> ProgIds { get; }

        internal string? PrimaryProgId { get; }

        internal string? VersionIndependentProgId { get; }
    }

    internal sealed class ManifestTypeLibEntry
    {
        internal ManifestTypeLibEntry(
            string dllPath,
            Guid tlbId,
            string version,
            ushort majorVersion,
            ushort minorVersion,
            string helpDirectory,
            string? resourceId,
            int flags)
        {
            DllPath = dllPath;
            TlbId = tlbId;
            Version = version;
            MajorVersion = majorVersion;
            MinorVersion = minorVersion;
            HelpDirectory = helpDirectory;
            ResourceId = resourceId;
            Flags = flags;
        }

        internal string DllPath { get; }

        internal Guid TlbId { get; }

        internal string Version { get; }

        internal ushort MajorVersion { get; }

        internal ushort MinorVersion { get; }

        internal string HelpDirectory { get; }

        internal string? ResourceId { get; }

        internal int Flags { get; }
    }

    internal sealed class ManifestInterfaceEntry
    {
        internal ManifestInterfaceEntry(Guid iid, Guid? tlbId, Guid proxyStubClsid)
        {
            Iid = iid;
            TlbId = tlbId;
            ProxyStubClsid = proxyStubClsid;
        }

        internal Guid Iid { get; }

        internal Guid? TlbId { get; }

        internal Guid ProxyStubClsid { get; }
    }

    private sealed class ImplicitProxyStubEntry
    {
        internal ImplicitProxyStubEntry(string dllPath, Guid clsid)
        {
            DllPath = dllPath;
            Clsid = clsid;
        }

        internal string DllPath { get; }

        internal Guid Clsid { get; }
    }

    private sealed class MutableComClassEntry
    {
        internal MutableComClassEntry(
            string dllPath,
            Guid clsid,
            Guid? tlbId,
            ComThreadingModel threadingModel,
            bool isPsFactoryBuffer)
        {
            DllPath = dllPath;
            Clsid = clsid;
            TlbId = tlbId;
            ThreadingModel = threadingModel;
            IsPsFactoryBuffer = isPsFactoryBuffer;
        }

        internal string DllPath { get; }

        internal Guid Clsid { get; }

        internal Guid? TlbId { get; }

        internal ComThreadingModel ThreadingModel { get; }

        internal bool IsPsFactoryBuffer { get; }

        internal string? ExplicitVersionIndependentProgId { get; set; }

        internal List<string> ProgIds { get; } = [];
    }

    private sealed class ElementScope
    {
        private System.Text.StringBuilder? _progIdBuilder;
        private int _progIdDepth = -1;

        internal bool HasAssemblyRoot { get; set; }

        internal string? CurrentFileName { get; set; }

        internal int FileDepth { get; set; } = -1;

        internal int DependentDepth { get; set; } = -1;

        internal MutableComClassEntry? CurrentComClass { get; set; }

        internal int ComClassDepth { get; set; } = -1;

        internal bool ReadingProgIdElement => _progIdDepth >= 0;

        internal void BeginProgIdElement(int depth)
        {
            _progIdDepth = depth;
            _progIdBuilder = new System.Text.StringBuilder(64);
        }

        internal bool TryAppendProgIdText(string text)
        {
            if (_progIdBuilder is null)
            {
                return false;
            }

            if (_progIdBuilder.Length + text.Length > 260)
            {
                return false;
            }

            _progIdBuilder.Append(text);
            return true;
        }

        internal bool OnEndElement(XmlReader reader, out int errorCode)
        {
            errorCode = 0;
            if (reader.Depth == _progIdDepth && string.Equals(reader.LocalName, "progid", StringComparison.Ordinal))
            {
                var progId = _progIdBuilder?.ToString().Trim();
                _progIdBuilder = null;
                _progIdDepth = -1;
                if (progId is null || !IsSafeProgId(progId))
                {
                    errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
                    return false;
                }

                if (CurrentComClass is not null && !ContainsIgnoreCase(CurrentComClass.ProgIds, progId))
                {
                    CurrentComClass.ProgIds.Add(progId);
                }

                return true;
            }

            if (reader.Depth == ComClassDepth && string.Equals(reader.LocalName, "comClass", StringComparison.Ordinal))
            {
                CurrentComClass = null;
                ComClassDepth = -1;
            }
            else if (reader.Depth == FileDepth && string.Equals(reader.LocalName, "file", StringComparison.Ordinal))
            {
                CurrentFileName = null;
                FileDepth = -1;
            }
            else if (reader.Depth == DependentDepth && string.Equals(reader.LocalName, "dependentAssembly", StringComparison.Ordinal))
            {
                DependentDepth = -1;
            }

            return true;
        }

        private static bool ContainsIgnoreCase(List<string> items, string candidate)
            => items.Exists(item => string.Equals(item, candidate, StringComparison.OrdinalIgnoreCase));
    }

    private sealed class ManifestParseState
    {
        private readonly List<MutableComClassEntry> _classes = [];
        private readonly Dictionary<Guid, MutableComClassEntry> _classesByClsid = [];
        private readonly List<ManifestTypeLibEntry> _typeLibs = [];
        private readonly Dictionary<string, ManifestTypeLibEntry> _typeLibsByKey = new(StringComparer.OrdinalIgnoreCase);

        internal ManifestParseState(string? assemblyDirectory, string? defaultAssemblyPath, bool allowUnresolvedDependencies = false)
        {
            AssemblyDirectory = assemblyDirectory;
            DefaultAssemblyPath = defaultAssemblyPath;
            AllowUnresolvedDependencies = allowUnresolvedDependencies;
        }

        internal string? AssemblyDirectory { get; }

        internal string? DefaultAssemblyPath { get; }

        internal bool AllowUnresolvedDependencies { get; }

        internal HashSet<string> VisitedManifests { get; } = new(StringComparer.OrdinalIgnoreCase);

        internal HashSet<Guid> ReferencedProxyStubClsids { get; } = [];

        internal List<ImplicitProxyStubEntry> ImplicitProxyStubClasses { get; } = [];

        internal List<ManifestInterfaceEntry> Interfaces { get; } = [];

        internal static void ResolveProgIdRoles(
            MutableComClassEntry item,
            out string? primaryProgId,
            out string? versionIndependentProgId)
        {
            if (item.ProgIds.Count == 0)
            {
                primaryProgId = null;
                versionIndependentProgId = null;
                return;
            }

            primaryProgId = item.ProgIds[0];
            if (item.ExplicitVersionIndependentProgId is not null)
            {
                versionIndependentProgId = item.ExplicitVersionIndependentProgId;
                return;
            }

            if (item.ProgIds.Count < 2)
            {
                versionIndependentProgId = null;
                return;
            }

            for (var i = 1; i < item.ProgIds.Count; i++)
            {
                if (!HasNumericVersionSuffix(item.ProgIds[i]))
                {
                    versionIndependentProgId = item.ProgIds[i];
                    return;
                }
            }

            versionIndependentProgId = !HasNumericVersionSuffix(primaryProgId)
                ? primaryProgId
                : item.ProgIds[1];
        }

        internal static bool HasNumericVersionSuffix(string progId)
        {
            var lastDot = progId.LastIndexOf('.');
            if (lastDot < 0 || lastDot == progId.Length - 1)
            {
                return false;
            }

            for (var i = lastDot + 1; i < progId.Length; i++)
            {
                if (progId[i] is < '0' or > '9')
                {
                    return false;
                }
            }

            return true;
        }

        internal bool TryAddComClass(MutableComClassEntry entry)
        {
            // Reject any duplicate or conflicting CLSID within the normalized manifest graph.
            if (_classesByClsid.ContainsKey(entry.Clsid))
            {
                return false;
            }

            _classesByClsid.Add(entry.Clsid, entry);
            _classes.Add(entry);
            return true;
        }

        internal void AddTypeLib(ManifestTypeLibEntry entry)
        {
            var key = entry.TlbId.ToString("N") + ":" + entry.Version;
            if (!_typeLibsByKey.ContainsKey(key))
            {
                _typeLibsByKey.Add(key, entry);
                _typeLibs.Add(entry);
            }
        }

        internal ParsedManifest? Build(out int errorCode)
        {
            errorCode = 0;
            var progIdOwners = new Dictionary<string, Guid>(StringComparer.OrdinalIgnoreCase);
            var dllPaths = new List<string>();
            var dllPathSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var finalizedClasses = new List<ManifestComClassEntry>(_classes.Count + ImplicitProxyStubClasses.Count);

            if (!TryFinalizeExplicitClasses(progIdOwners, finalizedClasses, dllPaths, dllPathSet))
            {
                errorCode = NativeMethods.ERROR_SXS_CANT_GEN_ACTCTX;
                return null;
            }

            AppendImplicitProxyStubClasses(finalizedClasses, dllPaths, dllPathSet);
            AppendTypeLibAndDefaultPaths(dllPaths, dllPathSet);
            return new ParsedManifest(finalizedClasses, _typeLibs, Interfaces, dllPaths);
        }

        private bool TryFinalizeExplicitClasses(
            Dictionary<string, Guid> progIdOwners,
            List<ManifestComClassEntry> finalizedClasses,
            List<string> dllPaths,
            HashSet<string> dllPathSet)
        {
            foreach (var item in _classes)
            {
                foreach (var progId in item.ProgIds)
                {
                    if (progIdOwners.TryGetValue(progId, out var existingOwner) && existingOwner != item.Clsid)
                    {
                        return false;
                    }

                    progIdOwners[progId] = item.Clsid;
                }

                var isKnownProxyStub = item.IsPsFactoryBuffer || ReferencedProxyStubClsids.Contains(item.Clsid);
                ResolveProgIdRoles(item, out var primaryProgId, out var versionIndependentProgId);
                finalizedClasses.Add(new ManifestComClassEntry(
                    item.DllPath,
                    item.Clsid,
                    item.TlbId,
                    item.ThreadingModel,
                    isKnownProxyStub,
                    [.. item.ProgIds],
                    primaryProgId,
                    versionIndependentProgId));

                if (dllPathSet.Add(item.DllPath))
                {
                    dllPaths.Add(item.DllPath);
                }
            }

            return true;
        }

        private void AppendImplicitProxyStubClasses(
            List<ManifestComClassEntry> finalizedClasses,
            List<string> dllPaths,
            HashSet<string> dllPathSet)
        {
            foreach (var ps in ImplicitProxyStubClasses)
            {
                if (_classesByClsid.ContainsKey(ps.Clsid))
                {
                    continue;
                }

                _classesByClsid.Add(ps.Clsid, new MutableComClassEntry(ps.DllPath, ps.Clsid, tlbId: null, ComThreadingModel.Both, isPsFactoryBuffer: true));
                finalizedClasses.Add(new ManifestComClassEntry(
                    ps.DllPath,
                    ps.Clsid,
                    tlbId: null,
                    ComThreadingModel.Both,
                    isKnownProxyStub: true,
                    progIds: [],
                    primaryProgId: null,
                    versionIndependentProgId: null));

                if (dllPathSet.Add(ps.DllPath))
                {
                    dllPaths.Add(ps.DllPath);
                }
            }
        }

        private void AppendTypeLibAndDefaultPaths(List<string> dllPaths, HashSet<string> dllPathSet)
        {
            foreach (var dllPath in _typeLibs.Select(static tlb => tlb.DllPath).Where(dllPathSet.Add))
            {
                dllPaths.Add(dllPath);
            }

            if (dllPaths.Count == 0 && DefaultAssemblyPath is not null && !string.IsNullOrWhiteSpace(DefaultAssemblyPath))
            {
                var fullDefault = Path.GetFullPath(DefaultAssemblyPath);
                if (dllPathSet.Add(fullDefault))
                {
                    dllPaths.Add(fullDefault);
                }
            }
        }
    }
}
#pragma warning restore SCS0018, SEC0116
