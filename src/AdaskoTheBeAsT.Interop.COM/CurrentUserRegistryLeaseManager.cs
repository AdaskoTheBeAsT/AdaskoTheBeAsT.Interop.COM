using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using Microsoft.Win32;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

#pragma warning disable MA0158 // Use System.Threading.Lock
#pragma warning disable SA1121 // Use built-in type alias
#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
internal static class CurrentUserRegistryLeaseManager
{
    private const string SoftwareClassesPrefix = @"Software\Classes\";

#if NET10_0_OR_GREATER
    private static readonly Lock RegistryLock = new();
#endif
#if NET472_OR_GREATER
    private static readonly object RegistryLock = new();
#endif
    private static readonly Dictionary<string, SharedValueLease> ActiveValueLeases = new(StringComparer.OrdinalIgnoreCase);

    internal static bool TryAcquireFallbackRegistryLease(
        ComManifestParser.ParsedManifest manifest,
        bool registerProgIds,
        bool registerTypeLibs,
        out RegistryScopeLease? scopeLease)
        => TryAcquireFallbackRegistryLease(
            manifest,
            registerProgIds,
            registerTypeLibs,
            skipWhenPresentInClassesRoot: true,
            out scopeLease,
            classesRootValueReader: null);

    internal static bool TryAcquireFallbackRegistryLease(
        ComManifestParser.ParsedManifest manifest,
        bool registerProgIds,
        bool registerTypeLibs,
        bool skipWhenPresentInClassesRoot,
        out RegistryScopeLease? scopeLease,
        Func<string, string, object?>? classesRootValueReader = null)
    {
        scopeLease = null;
        if (!registerProgIds && !registerTypeLibs)
        {
            return true;
        }

        var desiredWrites = BuildDesiredWrites(manifest, registerProgIds, registerTypeLibs);
        if (desiredWrites.Count == 0)
        {
            return true;
        }

        var effectiveReader = classesRootValueReader ?? ReadFromClassesRoot;

        lock (RegistryLock)
        {
            // Validate no active conflicting write exists before touching HKCU.
            foreach (var write in desiredWrites)
            {
                if (ActiveValueLeases.TryGetValue(write.CompositeKey, out var existing)
                    && !ValuesMatch(existing.WrittenValue, write.ValueData, write.IsGuidValue))
                {
                    return false;
                }
            }

            var acquiredEntries = new List<SharedValueLease>(desiredWrites.Count);
            try
            {
                foreach (var write in desiredWrites)
                {
                    var entry = AcquireSingleValueUnderLock(write, skipWhenPresentInClassesRoot, effectiveReader);
                    acquiredEntries.Add(entry);
                }

                scopeLease = new RegistryScopeLease(acquiredEntries);
                return true;
            }
            catch (Exception ex)
            {
                _ = ex;
                RollbackEntriesUnderLock(acquiredEntries);
                return false;
            }
        }
    }

    private static List<DesiredRegistryWrite> BuildDesiredWrites(
        ComManifestParser.ParsedManifest manifest,
        bool registerProgIds,
        bool registerTypeLibs)
    {
        var writes = new List<DesiredRegistryWrite>();
        var seenCompositeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (registerProgIds)
        {
            AppendProgIdWrites(manifest, writes, seenCompositeKeys);
        }

        if (registerTypeLibs)
        {
            AppendTypeLibWrites(manifest, writes, seenCompositeKeys);
        }

        return writes;
    }

    private static void AppendProgIdWrites(
        ComManifestParser.ParsedManifest manifest,
        List<DesiredRegistryWrite> writes,
        HashSet<string> seenCompositeKeys)
    {
        foreach (var cls in manifest.Classes)
        {
            if (cls.ProgIds.Count == 0)
            {
                continue;
            }

            var clsidText = FormatBracedUpperGuid(cls.Clsid);
            foreach (var progId in cls.ProgIds)
            {
                AddDesiredWrite(
                    writes,
                    seenCompositeKeys,
                    SoftwareClassesPrefix + progId + @"\CLSID",
                    valueName: string.Empty,
                    clsidText,
                    isGuidValue: true);
            }

            if (cls.PrimaryProgId is not null)
            {
                AddDesiredWrite(
                    writes,
                    seenCompositeKeys,
                    SoftwareClassesPrefix + @"CLSID\" + clsidText + @"\ProgID",
                    valueName: string.Empty,
                    cls.PrimaryProgId,
                    isGuidValue: false);
            }

            if (cls.VersionIndependentProgId is not null)
            {
                AddDesiredWrite(
                    writes,
                    seenCompositeKeys,
                    SoftwareClassesPrefix + @"CLSID\" + clsidText + @"\VersionIndependentProgID",
                    valueName: string.Empty,
                    cls.VersionIndependentProgId,
                    isGuidValue: false);
            }
        }
    }

    private static void AppendTypeLibWrites(
        ComManifestParser.ParsedManifest manifest,
        List<DesiredRegistryWrite> writes,
        HashSet<string> seenCompositeKeys)
    {
        var registeredTlbVersions = new Dictionary<Guid, string>();
        var platformSubkey = IntPtr.Size == 8 ? "win64" : "win32";
        foreach (var tlb in manifest.TypeLibs)
        {
            var tlbText = FormatBracedUpperGuid(tlb.TlbId);
            var versionBase = SoftwareClassesPrefix + @"TypeLib\" + tlbText + @"\" + tlb.Version;
            var loadPath = string.IsNullOrEmpty(tlb.ResourceId)
                ? tlb.DllPath
                : tlb.DllPath + System.IO.Path.DirectorySeparatorChar + tlb.ResourceId;

            AddDesiredWrite(writes, seenCompositeKeys, versionBase + @"\0\" + platformSubkey, string.Empty, loadPath, isGuidValue: false);
            AddDesiredWrite(
                writes,
                seenCompositeKeys,
                versionBase + @"\FLAGS",
                string.Empty,
                tlb.Flags.ToString(CultureInfo.InvariantCulture),
                isGuidValue: false);
            AddDesiredWrite(writes, seenCompositeKeys, versionBase + @"\HELPDIR", string.Empty, tlb.HelpDirectory, isGuidValue: false);
            registeredTlbVersions[tlb.TlbId] = tlb.Version;
        }

        foreach (var iface in manifest.Interfaces)
        {
            if (iface.TlbId is null || !registeredTlbVersions.TryGetValue(iface.TlbId.Value, out var version))
            {
                continue;
            }

            var iidBase = SoftwareClassesPrefix + @"Interface\" + FormatBracedUpperGuid(iface.Iid);
            AddDesiredWrite(writes, seenCompositeKeys, iidBase + @"\ProxyStubClsid32", string.Empty, FormatBracedUpperGuid(iface.ProxyStubClsid), isGuidValue: true);
            AddDesiredWrite(writes, seenCompositeKeys, iidBase + @"\TypeLib", string.Empty, FormatBracedUpperGuid(iface.TlbId.Value), isGuidValue: true);
            AddDesiredWrite(writes, seenCompositeKeys, iidBase + @"\TypeLib", "Version", version, isGuidValue: false);
        }
    }

    private static void AddDesiredWrite(
        List<DesiredRegistryWrite> writes,
        HashSet<string> seenCompositeKeys,
        string subKeyPath,
        string valueName,
        string valueData,
        bool isGuidValue)
    {
        var compositeKey = subKeyPath + "|" + valueName;
        if (seenCompositeKeys.Add(compositeKey))
        {
            writes.Add(new DesiredRegistryWrite(compositeKey, subKeyPath, valueName, valueData, isGuidValue));
        }
    }

    private static string FormatBracedUpperGuid(Guid guid)
        => guid.ToString("B", CultureInfo.InvariantCulture).ToUpperInvariant();

    private static bool ValuesMatch(string candidate, string expected, bool isGuidValue)
    {
        if (isGuidValue
            && Guid.TryParse(candidate, out var candidateGuid)
            && Guid.TryParse(expected, out var expectedGuid))
        {
            return candidateGuid == expectedGuid;
        }

        return string.Equals(candidate.Trim(), expected.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static object? ReadFromClassesRoot(string classesRootSubKeyPath, string valueName)
    {
        using var key = Registry.ClassesRoot.OpenSubKey(classesRootSubKeyPath, writable: false);
        return key?.GetValue(valueName, defaultValue: null, RegistryValueOptions.DoNotExpandEnvironmentNames);
    }

    private static bool IsAlreadyPresentInClassesRoot(
        DesiredRegistryWrite write,
        Func<string, string, object?> classesRootValueReader)
    {
        if (!write.SubKeyPath.StartsWith(SoftwareClassesPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var classesRootSubKeyPath = write.SubKeyPath.Substring(SoftwareClassesPrefix.Length);
#pragma warning disable CC0031 // Caller guarantees non-null classesRootValueReader.
        var raw = classesRootValueReader(classesRootSubKeyPath, write.ValueName);
#pragma warning restore CC0031
        return raw is string existingString
            && !string.IsNullOrWhiteSpace(existingString)
            && ValuesMatch(existingString, write.ValueData, write.IsGuidValue);
    }

    private static SharedValueLease AcquireSingleValueUnderLock(
        DesiredRegistryWrite write,
        bool skipWhenPresentInClassesRoot,
        Func<string, string, object?> classesRootValueReader)
    {
        if (ActiveValueLeases.TryGetValue(write.CompositeKey, out var existing))
        {
            if (existing.IsNoOpClassesRootMatch && !skipWhenPresentInClassesRoot)
            {
                PromoteNoOpLeaseToCurrentUserWriteUnderLock(existing);
            }

            existing.ActiveCount++;
            return existing;
        }

        if (skipWhenPresentInClassesRoot && IsAlreadyPresentInClassesRoot(write, classesRootValueReader))
        {
            var noOp = new SharedValueLease(
                write.CompositeKey,
                write.SubKeyPath,
                write.ValueName,
                write.ValueData,
                hadPriorValue: false,
                priorValue: null,
                RegistryValueKind.String,
                createdSubKeys: [],
                isNoOpClassesRootMatch: true);
            ActiveValueLeases.Add(write.CompositeKey, noOp);
            return noOp;
        }

        var createdSubKeys = CaptureMissingSubKeys(write.SubKeyPath);
        var hadPriorValue = TryReadPriorValue(write.SubKeyPath, write.ValueName, out var priorValue, out var priorValueKind);

        WriteCurrentUserStringValue(write.SubKeyPath, write.ValueName, write.ValueData);

        var created = new SharedValueLease(
            write.CompositeKey,
            write.SubKeyPath,
            write.ValueName,
            write.ValueData,
            hadPriorValue,
            priorValue,
            priorValueKind,
            createdSubKeys,
            isNoOpClassesRootMatch: false);
        ActiveValueLeases.Add(write.CompositeKey, created);
        return created;
    }

    private static void PromoteNoOpLeaseToCurrentUserWriteUnderLock(SharedValueLease lease)
    {
        var createdSubKeys = CaptureMissingSubKeys(lease.SubKeyPath);
        var hadPriorValue = TryReadPriorValue(lease.SubKeyPath, lease.ValueName, out var priorValue, out var priorValueKind);
        WriteCurrentUserStringValue(lease.SubKeyPath, lease.ValueName, lease.WrittenValue);
        lease.PromoteToMaterializedWrite(hadPriorValue, priorValue, priorValueKind, createdSubKeys);
    }

    private static void WriteCurrentUserStringValue(string subKeyPath, string valueName, string valueData)
    {
        using var key = Registry.CurrentUser.CreateSubKey(subKeyPath, writable: true);
        if (key is null)
        {
            throw new InvalidOperationException("Failed to open HKCU subkey: " + subKeyPath);
        }

        key.SetValue(valueName, valueData, RegistryValueKind.String);
    }

    private static List<string> CaptureMissingSubKeys(string subKeyPath)
    {
        var missing = new List<string>();
        var segments = subKeyPath.Split('\\');
        var current = string.Empty;
        for (var i = 0; i < segments.Length; i++)
        {
            current = i == 0 ? segments[0] : current + @"\" + segments[i];
            if (i < 2)
            {
                // Never attempt to delete "Software" or "Software\Classes".
                continue;
            }

            using var existing = Registry.CurrentUser.OpenSubKey(current, writable: false);
            if (existing is null)
            {
                missing.Add(current);
            }
        }

        return missing;
    }

    private static bool TryReadPriorValue(
        string subKeyPath,
        string valueName,
        out object? priorValue,
        out RegistryValueKind priorKind)
    {
        priorValue = null;
        priorKind = RegistryValueKind.String;

        using var key = Registry.CurrentUser.OpenSubKey(subKeyPath, writable: false);
        if (key is null)
        {
            return false;
        }

        var raw = key.GetValue(valueName, defaultValue: null, RegistryValueOptions.DoNotExpandEnvironmentNames);
        if (raw is null)
        {
            return false;
        }

        priorValue = raw;
        priorKind = key.GetValueKind(valueName);
        return true;
    }

    private static void RollbackEntriesUnderLock(List<SharedValueLease> entries)
    {
        for (var i = entries.Count - 1; i >= 0; i--)
        {
            var entry = entries[i];
            entry.ActiveCount--;
            if (entry.ActiveCount > 0)
            {
                continue;
            }

            ActiveValueLeases.Remove(entry.CompositeKey);
            RestoreOrDeleteOwnedValue(entry);
        }
    }

    private static void RestoreOrDeleteOwnedValue(SharedValueLease entry)
    {
        if (entry.IsNoOpClassesRootMatch)
        {
            return;
        }

        try
        {
            using (var key = Registry.CurrentUser.OpenSubKey(entry.SubKeyPath, writable: true))
            {
                if (key is null)
                {
                    return;
                }

                var current = key.GetValue(entry.ValueName, defaultValue: null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string;

                // Roll back ONLY values still owned by our scope; never clobber an external modification.
                if (!string.Equals(current, entry.WrittenValue, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                if (entry.HadPriorValue && entry.PriorValue is not null)
                {
                    key.SetValue(entry.ValueName, entry.PriorValue, entry.PriorValueKind);
                    return;
                }

                key.DeleteValue(entry.ValueName, throwOnMissingValue: false);
            }

            // Delete subkeys created by us in reverse depth order, only when completely empty.
            for (var i = entry.CreatedSubKeys.Count - 1; i >= 0; i--)
            {
                var candidate = entry.CreatedSubKeys[i];
                using var checkKey = Registry.CurrentUser.OpenSubKey(candidate, writable: false);
                if (checkKey is null || checkKey.SubKeyCount > 0 || checkKey.ValueCount > 0)
                {
                    break;
                }

                checkKey.Close();
                Registry.CurrentUser.DeleteSubKey(candidate, throwOnMissingSubKey: false);
            }
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or System.Security.SecurityException or ArgumentException or ObjectDisposedException)
        {
            _ = ex;
        }
    }

    internal sealed class RegistryScopeLease
    {
        private readonly List<SharedValueLease> _entries;
        private int _rolledBack;

        internal RegistryScopeLease(List<SharedValueLease> entries)
        {
            _entries = entries;
        }

        internal void Rollback()
        {
            if (Interlocked.CompareExchange(ref _rolledBack, 1, 0) != 0)
            {
                return;
            }

            lock (RegistryLock)
            {
                RollbackEntriesUnderLock(_entries);
            }
        }
    }

#pragma warning disable S107 // Internal lease record captures full prior-value and HKCR short-circuit metadata.
    internal sealed class SharedValueLease
    {
        internal SharedValueLease(
            string compositeKey,
            string subKeyPath,
            string valueName,
            string writtenValue,
            bool hadPriorValue,
            object? priorValue,
            RegistryValueKind priorValueKind,
            List<string> createdSubKeys,
            bool isNoOpClassesRootMatch)
        {
            CompositeKey = compositeKey;
            SubKeyPath = subKeyPath;
            ValueName = valueName;
            WrittenValue = writtenValue;
            HadPriorValue = hadPriorValue;
            PriorValue = priorValue;
            PriorValueKind = priorValueKind;
            CreatedSubKeys = createdSubKeys;
            IsNoOpClassesRootMatch = isNoOpClassesRootMatch;
            ActiveCount = 1;
        }

        internal string CompositeKey { get; }

        internal string SubKeyPath { get; }

        internal string ValueName { get; }

        internal string WrittenValue { get; }

        internal bool HadPriorValue { get; private set; }

        internal object? PriorValue { get; private set; }

        internal RegistryValueKind PriorValueKind { get; private set; }

        internal List<string> CreatedSubKeys { get; private set; }

        internal bool IsNoOpClassesRootMatch { get; private set; }

        internal int ActiveCount { get; set; }

        internal void PromoteToMaterializedWrite(
            bool hadPriorValue,
            object? priorValue,
            RegistryValueKind priorValueKind,
            List<string> createdSubKeys)
        {
            HadPriorValue = hadPriorValue;
            PriorValue = priorValue;
            PriorValueKind = priorValueKind;
            CreatedSubKeys = createdSubKeys;
            IsNoOpClassesRootMatch = false;
        }
    }
#pragma warning restore S107

    private sealed class DesiredRegistryWrite
    {
        internal DesiredRegistryWrite(
            string compositeKey,
            string subKeyPath,
            string valueName,
            string valueData,
            bool isGuidValue)
        {
            CompositeKey = compositeKey;
            SubKeyPath = subKeyPath;
            ValueName = valueName;
            ValueData = valueData;
            IsGuidValue = isGuidValue;
        }

        internal string CompositeKey { get; }

        internal string SubKeyPath { get; }

        internal string ValueName { get; }

        internal string ValueData { get; }

        internal bool IsGuidValue { get; }
    }
}
#pragma warning restore SA1121
#pragma warning restore MA0158
