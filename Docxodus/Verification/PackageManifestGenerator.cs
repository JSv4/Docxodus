// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Buffers;
using System.Buffers.Binary;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Docxodus.Verification;

/// <summary>
/// Builds non-mutating, deterministic package manifests directly from supplied DOCX/OPC bytes.
/// Malformed and encrypted inputs produce structured findings rather than partially opening a
/// mutable <c>WordprocessingDocument</c>.
/// </summary>
public static class PackageManifestGenerator
{
    private const string ContentTypesUri = "/[Content_Types].xml";
    private const string ContentTypesMime = "application/vnd.openxmlformats-package.content-types+xml";
    private const string RelationshipsMime = "application/vnd.openxmlformats-package.relationships+xml";
    private const string AnnotationNamespace = "http://docxodus.dev/annotations/v1";
    private const string TransitionalPackageRelationshipsNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string TransitionalContentTypesNamespace =
        "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string TransitionalWordNamespace =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StrictWordNamespace =
        "http://purl.oclc.org/ooxml/wordprocessingml/main";
    private const string TransitionalOfficeRelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string StrictOfficeRelationshipNamespace =
        "http://purl.oclc.org/ooxml/officeDocument/relationships";
    private const string TransitionalOfficeRelationshipTypePrefix =
        TransitionalOfficeRelationshipNamespace + "/";
    private const string StrictOfficeRelationshipTypePrefix =
        StrictOfficeRelationshipNamespace + "/";
    private const string Word2012Namespace =
        "http://schemas.microsoft.com/office/word/2012/wordml";
    private static readonly AsciiCaseInsensitiveComparer PartNameComparer =
        AsciiCaseInsensitiveComparer.Instance;
    private static readonly uint[] Crc32Table = CreateCrc32Table();
    private static readonly byte[] OleSignature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
    private static readonly HashSet<string> WordprocessingXmlContentTypes = new(
        StringComparer.OrdinalIgnoreCase)
    {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.stylesWithEffects+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.glossaryDocument+xml",
        "application/vnd.ms-word.document.macroEnabled.main+xml",
        "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
    };
    private static readonly HashSet<string> KnownOoxmlXmlContentTypes = new(
        WordprocessingXmlContentTypes, StringComparer.OrdinalIgnoreCase)
    {
        ContentTypesMime,
        RelationshipsMime,
        "application/vnd.openxmlformats-package.core-properties+xml",
        "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
        "application/vnd.openxmlformats-officedocument.theme+xml",
        "application/vnd.openxmlformats-officedocument.themeOverride+xml",
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
        "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml",
        "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
        "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
        "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
        "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
        "application/vnd.openxmlformats-officedocument.vmlDrawing",
        "application/vnd.ms-word.commentsExt+xml",
        "application/vnd.ms-word.commentsIds+xml",
        "application/vnd.ms-word.people+xml",
        "application/vnd.ms-office.chartstyle+xml",
        "application/vnd.ms-office.chartcolorstyle+xml",
    };

    /// <summary>Generate a schema-v1 package manifest.</summary>
    public static PackageManifest Generate(byte[] packageBytes, PackageManifestOptions? options = null) =>
        GenerateInspection(packageBytes, options, includeInspectionEntries: false).Manifest;

    /// <summary>
    /// Generate a manifest and retain parsed XML from that same bounded inspection pass for
    /// internal verification consumers. No archive, stream, or duplicate raw payload escapes.
    /// </summary>
    internal static PackageManifestInspection Inspect(
        byte[] packageBytes,
        PackageManifestOptions? options = null) =>
        GenerateInspection(packageBytes, options, includeInspectionEntries: true);

    private static PackageManifestInspection GenerateInspection(
        byte[] packageBytes,
        PackageManifestOptions? options,
        bool includeInspectionEntries)
    {
        ArgumentNullException.ThrowIfNull(packageBytes);
        options ??= new PackageManifestOptions();
        options.Validate();

        var rawDigest = Digest(packageBytes);
        var findings = new List<VerificationFinding>();
        if (packageBytes.AsSpan().StartsWith(OleSignature))
            return new PackageManifestInspection(
                OleManifest(packageBytes, rawDigest, findings),
                Array.Empty<PackageManifestInspectionEntry>());

        try
        {
            using var stream = new MemoryStream(packageBytes, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            return GenerateZipManifest(
                archive, packageBytes, rawDigest, options, findings, includeInspectionEntries);
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or ArgumentException
            or OverflowException)
        {
            AddFinding(findings, "malformed_package", VerificationFindingSeverity.Error,
                $"The supplied bytes are not a readable ZIP/OPC package ({ex.GetType().Name}).");
            return new PackageManifestInspection(
                FinalizeManifest(
                    "malformed", rawDigest, null, null,
                    Array.Empty<PackageManifestEntry>(),
                    Array.Empty<PackageContentTypeDeclaration>(),
                    Array.Empty<PackageRelationship>(),
                    new PackageManifestFacts(), findings),
                Array.Empty<PackageManifestInspectionEntry>());
        }
    }

    /// <summary>Generate the canonical schema-v1 JSON representation.</summary>
    public static string GenerateJson(
        byte[] packageBytes,
        PackageManifestOptions? options = null,
        bool indented = false) => Generate(packageBytes, options).ToJson(indented);

    private static PackageManifestInspection GenerateZipManifest(
        ZipArchive archive,
        byte[] packageBytes,
        VerificationDigest rawDigest,
        PackageManifestOptions options,
        List<VerificationFinding> findings,
        bool includeInspectionEntries)
    {
        var archiveEntries = archive.Entries.ToList();
        var centralMetadata = TryReadCentralDirectoryMetadata(packageBytes, archiveEntries.Count);
        if (centralMetadata is null)
        {
            AddFinding(findings, "zip_encryption_detection_unavailable",
                VerificationFindingSeverity.Error,
                "ZIP central-directory encryption flags could not be parsed authoritatively.",
                new ChangeLocation { PropertyPath = "entries[].isEncrypted" });
        }

        // Keep a name-only index for the complete central directory. A capped inspection cannot
        // read payloads beyond the cap, but it must not invent missing targets/owners/overrides or
        // misclassify an OPC package merely because their entries occur after the cutoff.
        var allEntryUris = archiveEntries
            .Select(entry => TryCanonicalizeEntryName(entry.FullName, out var uri) ? uri : null)
            .Where(uri => uri is not null)
            .Select(uri => uri!)
            .ToHashSet(PartNameComparer);

        // Declared expansion is measured over the whole central directory. Measuring it only over
        // the entries we go on to inspect would let a package dodge the budget by also breaching
        // the entry-count limit.
        var declaredTotal = SumDeclaredSizes(archiveEntries);
        var entryCountExceeded = archiveEntries.Count > options.MaxEntryCount;
        if (entryCountExceeded)
        {
            AddFinding(findings, "entry_count_limit_exceeded", VerificationFindingSeverity.Error,
                $"Package has {archiveEntries.Count.ToString(CultureInfo.InvariantCulture)} entries; " +
                $"the inspection limit is {options.MaxEntryCount.ToString(CultureInfo.InvariantCulture)}.",
                new ChangeLocation { PropertyPath = "entries" });
            archiveEntries = archiveEntries.Take(options.MaxEntryCount).ToList();
        }

        var works = new List<EntryWork>(archiveEntries.Count);
        for (var index = 0; index < archiveEntries.Count; index++)
        {
            var archiveEntry = archiveEntries[index];
            var validEntryPath = TryCanonicalizeEntryName(archiveEntry.FullName, out var uri);
            var isDirectory = validEntryPath && uri.EndsWith("/", StringComparison.Ordinal);
            if (!validEntryPath)
            {
                AddFinding(findings, "unsafe_entry_path", VerificationFindingSeverity.Error,
                    "ZIP entry name does not satisfy the OPC part-name segment grammar.",
                    new ChangeLocation { EntryUri = uri });
            }
            if (uri.Length > options.MaxUriLength)
            {
                AddFinding(findings, "entry_uri_limit_exceeded", VerificationFindingSeverity.Error,
                    $"Entry URI exceeds the {options.MaxUriLength.ToString(CultureInfo.InvariantCulture)} character limit.",
                    new ChangeLocation { EntryUri = uri });
            }

            long length;
            long compressedLength;
            try
            {
                length = archiveEntry.Length;
                compressedLength = archiveEntry.CompressedLength;
            }
            catch (InvalidDataException)
            {
                length = 0;
                compressedLength = 0;
                AddFinding(findings, "malformed_entry", VerificationFindingSeverity.Error,
                    "ZIP entry metadata could not be read.", new ChangeLocation { EntryUri = uri });
            }

            var entryLimitExceeded = length > options.MaxEntryUncompressedBytes;
            if (entryLimitExceeded)
            {
                AddFinding(findings, "entry_size_limit_exceeded",
                    VerificationFindingSeverity.Error,
                    $"Declared entry size exceeds the " +
                    $"{options.MaxEntryUncompressedBytes.ToString(CultureInfo.InvariantCulture)} byte limit.",
                    new ChangeLocation { EntryUri = uri });
            }
            bool? encrypted = centralMetadata is not null && index < centralMetadata.Count
                ? centralMetadata[index].IsEncrypted
                : null;
            if (encrypted == true)
            {
                AddFinding(findings, "unsupported_zip_encryption", VerificationFindingSeverity.Error,
                    "Encrypted ZIP entries are not supported.",
                    new ChangeLocation { EntryUri = uri });
            }

            var ratio = compressedLength == 0
                ? (length == 0 ? 0d : double.PositiveInfinity)
                : (double)length / compressedLength;
            var ratioExceeded = ratio > options.MaxCompressionRatio;
            if (ratioExceeded)
            {
                AddFinding(findings, "compression_ratio_limit_exceeded", VerificationFindingSeverity.Error,
                    "Entry expansion ratio exceeds the configured safety limit.",
                    new ChangeLocation { EntryUri = uri });
            }

            if (isDirectory && length == 0)
            {
                AddFinding(findings, "directory_entry", VerificationFindingSeverity.Warning,
                    "OPC packages should not contain directory-only ZIP entries.",
                    new ChangeLocation { EntryUri = uri });
            }
            else if (isDirectory)
            {
                AddFinding(findings, "nonempty_directory_entry", VerificationFindingSeverity.Error,
                    "A trailing-slash ZIP entry contains payload bytes and is not a directory artifact.",
                    new ChangeLocation { EntryUri = uri });
            }

            works.Add(new EntryWork
            {
                ArchiveEntry = archiveEntry,
                ArchiveIndex = index,
                Uri = uri,
                IsDirectory = isDirectory,
                Size = length,
                CompressedSize = compressedLength,
                IsEncrypted = encrypted,
                ExpectedCrc32 = centralMetadata is not null && index < centralMetadata.Count
                    ? centralMetadata[index].Crc32
                    : null,
                RatioExceeded = ratioExceeded,
                EntryLimitExceeded = entryLimitExceeded,
            });
        }

        var totalLimitExceeded = declaredTotal > options.MaxTotalUncompressedBytes;
        if (totalLimitExceeded)
        {
            AddFinding(findings, "total_expansion_limit_exceeded", VerificationFindingSeverity.Error,
                $"Declared uncompressed package size exceeds the " +
                $"{options.MaxTotalUncompressedBytes.ToString(CultureInfo.InvariantCulture)} byte limit.",
                new ChangeLocation { PropertyPath = "entries[].size" });
        }

        FindDuplicateEntryNames(works, findings);
        FindInterleavedPartNames(works, findings);
        var readBudget = new ActualReadBudget(options.MaxTotalUncompressedBytes);
        var contentTypeMap = ReadContentTypes(
            works, allEntryUris, totalLimitExceeded, options, readBudget, findings);
        foreach (var work in works)
        {
            (work.ContentType, work.ContentTypeSource) = contentTypeMap.Resolve(work.Uri);
            work.IsXml = !work.IsDirectory && IsXml(work.Uri, work.ContentType);

            // Only a readable [Content_Types].xml lets us claim a *part* has no declaration.
            // Reporting it per entry when the map itself was never parsed turns one systemic
            // failure into one error per entry and misnames the cause.
            if (work.ContentType is null && !work.IsDirectory && contentTypeMap.IsAvailable)
            {
                AddFinding(findings, "missing_content_type", VerificationFindingSeverity.Error,
                    "No content-type Override or Default matches this package entry.",
                    new ChangeLocation { EntryUri = work.Uri });
            }
        }
        if (!contentTypeMap.IsAvailable
            && works.Any(work => PartNameComparer.Equals(work.Uri, ContentTypesUri)))
        {
            AddFinding(findings, "content_types_unreadable", VerificationFindingSeverity.Error,
                "[Content_Types].xml is present but could not be used, so no part content type was resolved.",
                new ChangeLocation { EntryUri = ContentTypesUri });
        }
        ValidateContentTypeTargets(contentTypeMap, allEntryUris, findings);

        var payloadsInspected = !totalLimitExceeded && !readBudget.Exceeded;
        if (payloadsInspected)
        {
            foreach (var work in works)
            {
                if (!ReadEntry(work, options, readBudget, findings))
                    break;
            }
        }

        FindConflictingEntries(works, findings);
        AssignStableOccurrences(works);
        var relationships = ReadRelationships(
            works, allEntryUris, options, payloadsInspected, findings, out var unreadableOwners);
        ValidateRelationshipReferences(works, relationships, unreadableOwners, findings);
        var facts = BuildFacts(works, relationships);

        // A truncated inspection has not seen the whole package, so it cannot state a content
        // identity: two packages differing only past the cut would otherwise compare equal.
        VerificationDigest? orderedContentDigest = null;
        if (!totalLimitExceeded && !entryCountExceeded
            && works.All(work => work.RawBytesDigest is not null
                && work.IsEncrypted == false && !work.RatioExceeded
                && !work.EntryLimitExceeded))
        {
            orderedContentDigest = ComputeOrderedContentDigest(works);
        }

        // An XML part skipped for budget reasons might have normalized under a larger budget, so
        // substituting its raw bytes would make the package identity a function of the caller's
        // options. Only bytes that are provably not XML get the raw-byte fallback.
        VerificationDigest? semanticDigest = null;
        if (orderedContentDigest is not null
            && works.All(work => !work.IsXml
                || work.NormalizedXmlDigest is not null || work.XmlUnparsable))
        {
            semanticDigest = ComputeSemanticDigest(works);
        }

        var orderedWorks = works
            .OrderBy(work => work.Uri, StringComparer.Ordinal)
            .ThenBy(work => work.Occurrence)
            .ToList();
        var entryModels = orderedWorks
            .Select(work => new PackageManifestEntry
            {
                Uri = work.Uri,
                Occurrence = work.Occurrence,
                ContentType = work.ContentType,
                ContentTypeSource = work.ContentTypeSource,
                Size = work.Size,
                CompressedSize = work.CompressedSize,
                RawBytesDigest = work.RawBytesDigest,
                NormalizedXmlDigest = work.NormalizedXmlDigest,
                IsXml = work.IsXml,
                IsEncrypted = work.IsEncrypted,
            })
            .ToList();

        var hasContentTypes = allEntryUris.Contains(ContentTypesUri);
        var isEncrypted = centralMetadata?.Any(entry => entry.IsEncrypted) == true;
        var packageKind = isEncrypted
            ? "zip-encrypted"
            : hasContentTypes ? "opc" : "zip";
        var manifest = FinalizeManifest(
            packageKind, rawDigest, orderedContentDigest, semanticDigest,
            entryModels, contentTypeMap.Declarations, relationships, facts, findings);
        if (!includeInspectionEntries)
            return new PackageManifestInspection(
                manifest, Array.Empty<PackageManifestInspectionEntry>());

        var inspectedEntries = orderedWorks
            .Select((work, index) => new PackageManifestInspectionEntry(
                entryModels[index], work.Xml))
            .ToList();
        return new PackageManifestInspection(manifest, inspectedEntries);
    }

    private static PackageManifest OleManifest(
        byte[] bytes,
        VerificationDigest rawDigest,
        List<VerificationFinding> findings)
    {
        var encrypted = ContainsEncryptedOoxmlStreams(bytes);
        AddFinding(findings,
            encrypted ? "unsupported_ole_encryption" : "unsupported_compound_file",
            VerificationFindingSeverity.Error,
            encrypted
                ? "Password-encrypted OOXML is wrapped in an OLE compound file and is not supported."
                : "OLE compound files are not OPC/ZIP packages and are not supported.");
        return FinalizeManifest(
            encrypted ? "ole-encrypted" : "ole", rawDigest, null, null,
            Array.Empty<PackageManifestEntry>(),
            Array.Empty<PackageContentTypeDeclaration>(),
            Array.Empty<PackageRelationship>(),
            new PackageManifestFacts(), findings);
    }

    private static PackageManifest FinalizeManifest(
        string packageKind,
        VerificationDigest rawDigest,
        VerificationDigest? orderedContentDigest,
        VerificationDigest? semanticDigest,
        IReadOnlyList<PackageManifestEntry> entries,
        IReadOnlyList<PackageContentTypeDeclaration> contentTypes,
        IReadOnlyList<PackageRelationship> relationships,
        PackageManifestFacts facts,
        List<VerificationFinding> findings)
    {
        var orderedFindings = findings
            .OrderByDescending(finding => finding.Severity)
            .ThenBy(finding => finding.Code, StringComparer.Ordinal)
            .ThenBy(finding => finding.Location?.EntryUri ?? string.Empty, StringComparer.Ordinal)
            .ThenBy(finding => finding.Location?.OwnerUri ?? string.Empty, StringComparer.Ordinal)
            .ThenBy(finding => finding.Location?.RelationshipId ?? string.Empty, StringComparer.Ordinal)
            .ThenBy(finding => finding.Location?.TargetUri ?? string.Empty, StringComparer.Ordinal)
            .ThenBy(finding => finding.Location?.PropertyPath ?? string.Empty, StringComparer.Ordinal)
            .ThenBy(finding => finding.Message, StringComparer.Ordinal)
            .ToList();
        return new PackageManifest
        {
            PackageKind = packageKind,
            IsValid = orderedFindings.All(finding =>
                finding.Severity != VerificationFindingSeverity.Error),
            RawPackageBytesDigest = rawDigest,
            OrderedOpcContentDigest = orderedContentDigest,
            NormalizedSemanticDigest = semanticDigest,
            Entries = entries,
            ContentTypes = contentTypes,
            Relationships = relationships,
            Facts = facts,
            Findings = orderedFindings,
        };
    }

    private static ContentTypeMap ReadContentTypes(
        IReadOnlyList<EntryWork> works,
        IReadOnlySet<string> allEntryUris,
        bool totalLimitExceeded,
        PackageManifestOptions options,
        ActualReadBudget readBudget,
        List<VerificationFinding> findings)
    {
        var candidates = works.Where(work =>
                PartNameComparer.Equals(work.Uri, ContentTypesUri))
            .OrderBy(work => work.ArchiveIndex)
            .ToList();
        if (candidates.Count == 0)
        {
            if (!allEntryUris.Contains(ContentTypesUri))
            {
                AddFinding(findings, "missing_content_types", VerificationFindingSeverity.Error,
                    "The package has no [Content_Types].xml entry.",
                    new ChangeLocation { EntryUri = ContentTypesUri });
            }
            return ContentTypeMap.Empty;
        }
        if (totalLimitExceeded)
            return ContentTypeMap.Empty;

        var selected = candidates[0];
        if (selected.IsEncrypted != false || selected.RatioExceeded
            || selected.EntryLimitExceeded)
            return ContentTypeMap.Empty;
        if (selected.Size > options.MaxXmlPartBytes)
        {
            selected.XmlLimitReported = true;
            AddFinding(findings, "xml_size_limit_exceeded", VerificationFindingSeverity.Error,
                "[Content_Types].xml exceeds the configured XML parsing limit.",
                new ChangeLocation { EntryUri = selected.Uri });
            return ContentTypeMap.Empty;
        }

        try
        {
            var bytes = ReadAllBounded(
                selected.ArchiveEntry,
                options.MaxXmlPartBytes,
                options.MaxEntryUncompressedBytes,
                ExpansionCeiling(selected.CompressedSize, options.MaxCompressionRatio),
                readBudget);
            selected.PreloadedBytes = bytes;
            var document = XmlSemanticNormalizer.Parse(bytes,
                XmlCharacterLimit(options.MaxXmlPartBytes));
            return ContentTypeMap.Parse(document, options.MaxUriLength, findings);
        }
        catch (ManifestSafetyException ex)
        {
            selected.ReadBlocked = true;
            AddFinding(findings, SafetyFindingCode(ex.Kind),
                VerificationFindingSeverity.Error,
                SafetyFindingMessage(ex.Kind, "[Content_Types].xml"),
                new ChangeLocation { EntryUri = ContentTypesUri });
            return ContentTypeMap.Empty;
        }
        catch (XmlDepthLimitException)
        {
            selected.XmlLimitReported = true;
            AddFinding(findings, "xml_depth_limit_exceeded", VerificationFindingSeverity.Error,
                $"[Content_Types].xml exceeds the {XmlSemanticNormalizer.MaxElementDepth.ToString(CultureInfo.InvariantCulture)} level XML nesting limit.",
                new ChangeLocation { EntryUri = ContentTypesUri });
            return ContentTypeMap.Empty;
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or XmlException
            or UnauthorizedAccessException)
        {
            AddFinding(findings, "malformed_content_types", VerificationFindingSeverity.Error,
                $"[Content_Types].xml is not readable XML ({ex.GetType().Name}).",
                new ChangeLocation { EntryUri = ContentTypesUri });
            return ContentTypeMap.Empty;
        }
    }

    private static bool ReadEntry(
        EntryWork work,
        PackageManifestOptions options,
        ActualReadBudget readBudget,
        List<VerificationFinding> findings)
    {
        if (work.IsEncrypted != false || work.RatioExceeded || work.EntryLimitExceeded)
            return true;
        if (work.ReadBlocked)
            return true;

        var captureXml = work.IsXml && work.Size <= options.MaxXmlPartBytes;
        if (work.IsXml && !captureXml && !work.XmlLimitReported)
        {
            work.XmlLimitReported = true;
            AddFinding(findings, "xml_size_limit_exceeded", VerificationFindingSeverity.Error,
                "XML entry exceeds the configured XML parsing limit; its raw digest is retained.",
                new ChangeLocation { EntryUri = work.Uri });
        }

        try
        {
            if (work.PreloadedBytes is { } preloaded)
            {
                work.ActualSize = preloaded.LongLength;
                work.RawBytesDigest = Digest(preloaded);
                ValidateCrc32(work, ComputeCrc32(preloaded), findings);
                if (work.ActualSize != work.Size)
                {
                    AddFinding(findings, "entry_size_mismatch", VerificationFindingSeverity.Error,
                        "The decompressed byte count differs from the central-directory size.",
                        new ChangeLocation { EntryUri = work.Uri });
                }
                if (captureXml)
                    NormalizeXml(work, preloaded, options, findings);
                return true;
            }

            using var input = work.ArchiveEntry.Open();
            using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            using var xml = captureXml ? new MemoryStream((int)Math.Min(work.Size, int.MaxValue)) : null;
            var buffer = new byte[81920];
            long readTotal = 0;
            var expansionRemaining = ExpansionCeiling(
                work.CompressedSize, options.MaxCompressionRatio);
            var crc32 = uint.MaxValue;
            var actualXmlLimitExceeded = false;
            int read;
            while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                if (read > expansionRemaining)
                    throw new ManifestSafetyException(SafetyLimitKind.EntryExpansion);
                expansionRemaining -= read;
                if (read > options.MaxEntryUncompressedBytes - readTotal)
                    throw new ManifestSafetyException(SafetyLimitKind.EntrySize);
                if (!readBudget.TryConsume(read))
                    throw new ManifestSafetyException(SafetyLimitKind.TotalExpansion);
                if (readTotal > long.MaxValue - read)
                    throw new ManifestSafetyException(SafetyLimitKind.EntryExpansion);
                readTotal += read;
                hash.AppendData(buffer, 0, read);
                crc32 = AppendCrc32(crc32, buffer.AsSpan(0, read));
                if (xml is not null && !actualXmlLimitExceeded)
                {
                    if (readTotal > options.MaxXmlPartBytes)
                    {
                        actualXmlLimitExceeded = true;
                        AddFinding(findings, "xml_size_limit_exceeded",
                            VerificationFindingSeverity.Error,
                            "Actual XML entry bytes exceed the configured XML parsing limit; " +
                            "its raw digest is retained.",
                            new ChangeLocation { EntryUri = work.Uri });
                    }
                    else
                    {
                        xml.Write(buffer, 0, read);
                    }
                }
            }
            work.ActualSize = readTotal;
            ValidateCrc32(work, ~crc32, findings);
            if (readTotal != work.Size)
            {
                AddFinding(findings, "entry_size_mismatch", VerificationFindingSeverity.Error,
                    "The decompressed byte count differs from the central-directory size.",
                    new ChangeLocation { EntryUri = work.Uri });
            }
            work.RawBytesDigest = new VerificationDigest
            {
                Algorithm = "SHA-256",
                Value = Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant(),
            };

            if (xml is null || actualXmlLimitExceeded)
                return true;
            NormalizeXml(work, xml.ToArray(), options, findings);
            return true;
        }
        catch (ManifestSafetyException ex)
        {
            AddFinding(findings, SafetyFindingCode(ex.Kind), VerificationFindingSeverity.Error,
                SafetyFindingMessage(ex.Kind, "Entry"),
                new ChangeLocation { EntryUri = work.Uri });
            return false;
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException
            or NotSupportedException or CryptographicException)
        {
            AddFinding(findings, "unreadable_entry", VerificationFindingSeverity.Error,
                $"Entry payload could not be read ({ex.GetType().Name}).",
                new ChangeLocation { EntryUri = work.Uri });
            return true;
        }
    }

    private static void NormalizeXml(
        EntryWork work,
        byte[] bytes,
        PackageManifestOptions options,
        List<VerificationFinding> findings)
    {
        try
        {
            var document = XmlSemanticNormalizer.Parse(bytes,
                XmlCharacterLimit(options.MaxXmlPartBytes));
            work.Xml = document;
            work.NormalizedXmlDigest = XmlSemanticNormalizer.Digest(
                document, work.Uri, IsKnownOoxmlXml(work));
        }
        catch (XmlDepthLimitException)
        {
            work.XmlLimitReported = true;
            AddFinding(findings, "xml_depth_limit_exceeded", VerificationFindingSeverity.Error,
                $"XML entry exceeds the {XmlSemanticNormalizer.MaxElementDepth.ToString(CultureInfo.InvariantCulture)} level nesting limit.",
                new ChangeLocation { EntryUri = work.Uri });
        }
        catch (Exception ex) when (ex is XmlException or InvalidOperationException)
        {
            // The bytes themselves are not XML, so no budget would ever normalize them. That is
            // a stable fact about the package, unlike an entry we merely declined to read.
            work.XmlUnparsable = true;
            AddFinding(findings, "malformed_xml", VerificationFindingSeverity.Error,
                $"XML entry could not be normalized ({ex.GetType().Name}).",
                new ChangeLocation { EntryUri = work.Uri });
        }
    }

    private static IReadOnlyList<PackageRelationship> ReadRelationships(
        IReadOnlyList<EntryWork> works,
        IReadOnlySet<string> allEntryUris,
        PackageManifestOptions options,
        bool payloadsInspected,
        List<VerificationFinding> findings,
        out HashSet<string> unreadableOwners)
    {
        var entryUris = allEntryUris;
        unreadableOwners = new HashSet<string>(PartNameComparer);
        var relationships = new List<PackageRelationship>();
        foreach (var work in works.Where(work => IsRelationshipPart(work.Uri))
                     .OrderBy(work => work.Uri, StringComparer.Ordinal)
                     .ThenBy(work => work.Occurrence))
        {
            if (work.Xml?.Root is null)
            {
                // The part exists but was never parsed. Say so, rather than emitting an empty
                // relationship set that reads as "this part declares nothing" — but only when
                // payloads were inspected at all. If a package-wide limit stopped every read,
                // that breach is already reported once and blaming each .rels part repeats it.
                var skippedOwner = RelationshipOwner(work.Uri);
                if (skippedOwner is not null)
                    unreadableOwners.Add(skippedOwner);
                if (payloadsInspected)
                {
                    AddFinding(findings, "relationship_part_unreadable", VerificationFindingSeverity.Error,
                        "Relationship part could not be parsed, so its relationships are unknown.",
                        new ChangeLocation { EntryUri = work.Uri, OwnerUri = skippedOwner });
                }
                continue;
            }
            if (work.Xml.Root.Name.LocalName != "Relationships"
                || !IsPackageRelationshipsNamespace(work.Xml.Root.Name.NamespaceName))
            {
                AddFinding(findings, "malformed_relationship_part",
                    VerificationFindingSeverity.Error,
                    "Relationship part root must use an OPC Relationships namespace.",
                    new ChangeLocation { EntryUri = work.Uri });
                continue;
            }
            var owner = RelationshipOwner(work.Uri);
            if (owner is null)
            {
                AddFinding(findings, "malformed_relationship_part", VerificationFindingSeverity.Error,
                    "Relationship part URI cannot be mapped to an owner part.",
                    new ChangeLocation { EntryUri = work.Uri });
                continue;
            }
            if (owner != "/" && !entryUris.Contains(owner))
            {
                AddFinding(findings, "missing_relationship_owner", VerificationFindingSeverity.Error,
                    "Relationship part owner does not exist in the package.",
                    new ChangeLocation { EntryUri = work.Uri, OwnerUri = owner });
            }

            foreach (var element in work.Xml.Root.Elements()
                         .Where(element => element.Name.LocalName == "Relationship"
                             && element.Name.NamespaceName == work.Xml.Root.Name.NamespaceName))
            {
                var id = UnqualifiedAttribute(element, "Id");
                var type = UnqualifiedAttribute(element, "Type");
                var target = UnqualifiedAttribute(element, "Target");
                var rawMode = UnqualifiedAttribute(element, "TargetMode");
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(type)
                    || string.IsNullOrEmpty(target))
                {
                    AddFinding(findings, "malformed_relationship", VerificationFindingSeverity.Error,
                        "Relationship is missing Id, Type, or Target.",
                        new ChangeLocation { EntryUri = work.Uri, OwnerUri = owner,
                            RelationshipId = id, TargetUri = target });
                    continue;
                }

                var external = string.Equals(rawMode, "External", StringComparison.Ordinal);
                var validTargetMode = string.IsNullOrEmpty(rawMode)
                    || external
                    || string.Equals(rawMode, "Internal", StringComparison.Ordinal);
                if (!validTargetMode)
                {
                    AddFinding(findings, "invalid_target_mode", VerificationFindingSeverity.Error,
                        "Relationship TargetMode must be Internal or External.",
                        new ChangeLocation { EntryUri = work.Uri, OwnerUri = owner,
                            RelationshipId = id, TargetUri = target });
                }

                string? resolved = null;
                bool? targetPresent = null;
                if (validTargetMode && !external)
                {
                    resolved = ResolveRelationshipTarget(
                        owner, target, options.MaxUriLength, out var invalidTarget);
                    targetPresent = resolved is not null && entryUris.Contains(resolved);
                    if (invalidTarget || resolved is null)
                    {
                        AddFinding(findings, "invalid_relationship_target", VerificationFindingSeverity.Error,
                            "Internal relationship target cannot be resolved within the package root.",
                            new ChangeLocation { EntryUri = work.Uri, OwnerUri = owner,
                                RelationshipId = id, TargetUri = target });
                    }
                    else if (targetPresent == false)
                    {
                        AddFinding(findings, "missing_target", VerificationFindingSeverity.Error,
                            "Internal relationship target is absent from the package.",
                            new ChangeLocation { EntryUri = work.Uri, OwnerUri = owner,
                                RelationshipId = id, TargetUri = resolved });
                    }
                }

                relationships.Add(new PackageRelationship
                {
                    OwnerUri = owner,
                    Id = id,
                    Type = type,
                    Target = target,
                    // Missing means Internal by OPC default. Invalid spellings are reported but
                    // use the closed Internal fallback on the wire so malformed relationships
                    // remain consumable and do not disappear from the inventory.
                    TargetMode = external ? "External" : "Internal",
                    ResolvedTargetUri = resolved,
                    IsTargetPresent = targetPresent,
                });
            }
        }

        foreach (var group in relationships.GroupBy(
                     relationship => (relationship.OwnerUri, relationship.Id),
                     OwnerIdComparer.Instance))
        {
            if (group.Count() <= 1)
                continue;
            var conflicting = group.Select(relationship =>
                    (relationship.Type, relationship.Target, relationship.TargetMode))
                .Distinct()
                .Skip(1)
                .Any();
            AddFinding(findings,
                conflicting ? "conflicting_relationship" : "duplicate_relationship",
                VerificationFindingSeverity.Error,
                conflicting
                    ? "Relationship ID is repeated with conflicting type/target values."
                    : "Relationship ID is repeated within the same owner.",
                new ChangeLocation { OwnerUri = group.Key.OwnerUri,
                    RelationshipId = group.Key.Id });
        }

        return relationships
            .OrderBy(relationship => relationship.OwnerUri, StringComparer.Ordinal)
            .ThenBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ThenBy(relationship => relationship.Type, StringComparer.Ordinal)
            .ThenBy(relationship => relationship.Target, StringComparer.Ordinal)
            .ThenBy(relationship => relationship.TargetMode, StringComparer.Ordinal)
            .ToList();
    }

    private static void ValidateRelationshipReferences(
        IReadOnlyList<EntryWork> works,
        IReadOnlyList<PackageRelationship> relationships,
        HashSet<string> unreadableOwners,
        List<VerificationFinding> findings)
    {
        var idsByOwner = relationships
            .GroupBy(relationship => relationship.OwnerUri, PartNameComparer)
            .ToDictionary(group => group.Key,
                group => group.Select(relationship => relationship.Id)
                    .ToHashSet(StringComparer.Ordinal),
                PartNameComparer);
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        foreach (var work in works.Where(work => work.Xml?.Root is not null
                     && !IsRelationshipPart(work.Uri)
                     && !PartNameComparer.Equals(work.Uri, ContentTypesUri)))
        {
            // Without a parsed .rels part we do not know which IDs the part defines, so every
            // reference would look dangling. relationship_part_unreadable already named the cause.
            if (unreadableOwners.Contains(work.Uri))
                continue;
            idsByOwner.TryGetValue(work.Uri, out var ids);
            foreach (var attribute in work.Xml!.Descendants()
                         .Attributes()
                         .Where(attribute =>
                             IsOfficeRelationshipNamespace(attribute.Name.NamespaceName)
                             && attribute.Name.LocalName is
                                 "id" or "embed" or "link" or "dm" or "lo" or "qs" or "cs" or "txbx"
                             && !string.IsNullOrWhiteSpace(attribute.Value)))
            {
                if (ids?.Contains(attribute.Value) == true)
                    continue;
                var key = work.Uri + "\0" + attribute.Value;
                if (!emitted.Add(key))
                    continue;
                AddFinding(findings, "dangling_relationship", VerificationFindingSeverity.Error,
                    "XML references a relationship ID that its owning part does not define.",
                    new ChangeLocation { EntryUri = work.Uri, OwnerUri = work.Uri,
                        RelationshipId = attribute.Value });
            }
        }
    }
    private static PackageManifestFacts BuildFacts(
        IReadOnlyList<EntryWork> works,
        IReadOnlyList<PackageRelationship> relationships)
    {
        var mainDocumentUri = relationships
            .Where(relationship => relationship.OwnerUri == "/"
                && IsOfficeRelationshipType(relationship.Type, "officeDocument")
                && relationship.TargetMode == "Internal")
            .Select(relationship => relationship.ResolvedTargetUri)
            .FirstOrDefault(uri => uri is not null && works.Any(work =>
                PartNameComparer.Equals(work.Uri, uri)
                && IsWordprocessingMainPart(work)));

        var sectionCount = 0;
        var paragraphCount = 0;
        var tableCount = 0;
        var drawingCount = 0;
        var altChunkCount = 0;
        var fieldCount = 0;
        var footnoteCount = 0;
        var endnoteCount = 0;
        var styleCount = 0;
        var numberingCount = 0;
        var insertions = 0;
        var deletions = 0;
        var moveFrom = 0;
        var moveTo = 0;
        var propertyChanges = 0;
        var structuralChanges = 0;
        var otherChanges = 0;
        var comments = 0;
        var replies = 0;
        var commentMetadata = 0;
        var resolvedComments = 0;
        var people = 0;
        var annotations = 0;
        var isStrict = relationships.Any(relationship =>
            relationship.Type.StartsWith(StrictOfficeRelationshipTypePrefix, StringComparison.Ordinal));

        foreach (var work in works.Where(work => work.Xml?.Root is not null))
        {
            var root = work.Xml!.Root!;
            var elements = root.DescendantsAndSelf().ToList();
            var isWordPart = IsWordprocessingPart(work);
            var wordElements = isWordPart
                ? elements.Where(IsWordprocessingElement).ToList()
                : new List<XElement>();
            var storyElements = IsWordprocessingStoryPart(work)
                ? wordElements
                : new List<XElement>();
            isStrict |= isWordPart && wordElements.Any(element =>
                element.Name.NamespaceName == StrictWordNamespace);
            paragraphCount += storyElements.Count(element => element.Name.LocalName == "p");
            tableCount += storyElements.Count(element => element.Name.LocalName == "tbl");
            drawingCount += storyElements.Count(element =>
                element.Name.LocalName is "drawing" or "pict");
            altChunkCount += storyElements.Count(element => element.Name.LocalName == "altChunk");
            fieldCount += storyElements.Count(element => element.Name.LocalName == "fldSimple")
                + storyElements.Count(element => element.Name.LocalName == "fldChar"
                    && string.Equals(ElementNamespaceAttribute(element, "fldCharType"), "begin",
                        StringComparison.OrdinalIgnoreCase));
            insertions += storyElements.Count(element => element.Name.LocalName == "ins");
            deletions += storyElements.Count(element => element.Name.LocalName == "del");
            moveFrom += storyElements.Count(element => element.Name.LocalName == "moveFrom");
            moveTo += storyElements.Count(element => element.Name.LocalName == "moveTo");
            propertyChanges += storyElements.Count(element =>
                IsPropertyRevisionName(element.Name.LocalName));
            structuralChanges += storyElements.Count(element =>
                element.Name.LocalName is "cellIns" or "cellDel" or "cellMerge");
            otherChanges += storyElements.Count(element =>
                IsCustomXmlRevisionRangeStart(element.Name.LocalName));

            if (mainDocumentUri is not null
                && isWordPart
                && PartNameComparer.Equals(work.Uri, mainDocumentUri))
                sectionCount += wordElements.Count(element => element.Name.LocalName == "sectPr");
            if (IsContentType(work, "footnotes"))
                footnoteCount += CountPositiveDefinitions(wordElements, "footnote");
            if (IsContentType(work, "endnotes"))
                endnoteCount += CountPositiveDefinitions(wordElements, "endnote");
            if (IsContentType(work, "styles"))
                styleCount += wordElements.Count(element => element.Name.LocalName == "style");
            if (IsContentType(work, "numbering"))
                numberingCount += wordElements.Count(element =>
                    element.Name.LocalName is "abstractNum" or "num");
            if (IsCommentsPart(work))
                comments += wordElements.Count(element => element.Name.LocalName == "comment");

            var exElements = IsMime(work, "application/vnd.ms-word.commentsExt+xml")
                ? elements.Where(element => element.Name.LocalName == "commentEx"
                    && element.Name.NamespaceName == Word2012Namespace).ToList()
                : new List<XElement>();
            commentMetadata += exElements.Count;
            replies += exElements.Count(element =>
                !string.IsNullOrEmpty(ElementNamespaceAttribute(element, "paraIdParent")));
            resolvedComments += exElements.Count(element =>
                IsOn(ElementNamespaceAttribute(element, "done")));
            if (IsMime(work, "application/vnd.ms-word.people+xml"))
                people += elements.Count(element => element.Name.LocalName == "person"
                    && element.Name.NamespaceName == Word2012Namespace);
            if (root.Name.NamespaceName == AnnotationNamespace
                && root.Name.LocalName == "annotations")
                annotations += root.Elements().Count(element => element.Name.LocalName == "annotation");
        }

        var revisions = new PackageRevisionCounts
        {
            Insertions = insertions,
            Deletions = deletions,
            MoveFrom = moveFrom,
            MoveTo = moveTo,
            PropertyChanges = propertyChanges,
            StructuralChanges = structuralChanges,
            OtherChanges = otherChanges,
            Total = insertions + deletions + moveFrom + moveTo + propertyChanges
                + structuralChanges + otherChanges,
        };
        return new PackageManifestFacts
        {
            MainDocumentUri = mainDocumentUri,
            IsStrictOoxml = isStrict,
            IsMacroEnabled = works.Any(work =>
                IsMime(work, "application/vnd.ms-word.document.macroEnabled.main+xml")
                || IsMime(work, "application/vnd.ms-word.template.macroEnabledTemplate.main+xml")
                || IsMime(work, "application/vnd.ms-office.vbaProject")),
            HasCoreProperties = works.Any(work =>
                IsMime(work, "application/vnd.openxmlformats-package.core-properties+xml")),
            HasExtendedProperties = works.Any(work =>
                IsMime(work, "application/vnd.openxmlformats-officedocument.extended-properties+xml")),
            HasCustomProperties = works.Any(work =>
                IsMime(work, "application/vnd.openxmlformats-officedocument.custom-properties+xml")),
            SectionCount = sectionCount,
            ParagraphCount = paragraphCount,
            TableCount = tableCount,
            HeaderPartCount = works.Count(work => IsContentType(work, "header")),
            FooterPartCount = works.Count(work => IsContentType(work, "footer")),
            FootnoteCount = footnoteCount,
            EndnoteCount = endnoteCount,
            StyleCount = styleCount,
            NumberingDefinitionCount = numberingCount,
            ThemePartCount = works.Count(work =>
                IsMime(work, "application/vnd.openxmlformats-officedocument.theme+xml")),
            MediaPartCount = works.Count(work => work.ContentType is { } contentType
                && (IsMediaTypeFamily(contentType, "image")
                    || IsMediaTypeFamily(contentType, "audio")
                    || IsMediaTypeFamily(contentType, "video"))),
            CustomXmlPartCount = works.Count(IsCustomXmlDataPart),
            DrawingCount = drawingCount,
            AltChunkCount = altChunkCount,
            FieldCount = fieldCount,
            Revisions = revisions,
            Annotations = new PackageAnnotationCounts
            {
                Comments = comments,
                CommentReplies = replies,
                ThreadedCommentMetadata = commentMetadata,
                ResolvedComments = resolvedComments,
                People = people,
                DocxodusAnnotations = annotations,
            },
        };
    }

    private static VerificationDigest ComputeOrderedContentDigest(
        IReadOnlyList<EntryWork> works)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        foreach (var work in StableEntryOrder(works))
        {
            AppendString(hash, work.Uri);
            AppendInt32(hash, work.Occurrence);
            AppendInt64(hash, work.ActualSize);
            AppendString(hash, work.RawBytesDigest!.Value);
        }
        return new VerificationDigest
        {
            Algorithm = "SHA-256",
            Value = Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant(),
        };
    }

    private static VerificationDigest ComputeSemanticDigest(IReadOnlyList<EntryWork> works)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        foreach (var work in StableEntryOrder(works))
        {
            AppendString(hash, work.Uri);
            AppendInt32(hash, work.Occurrence);
            AppendString(hash, work.ContentType ?? string.Empty);

            // 'X' normalized XML, 'B' opaque binary, 'U' an entry that claims to be XML but could
            // not be parsed. 'U' falls back to the exact bytes so one unparsable part costs that
            // part its serialization-independence rather than costing the package its identity.
            var normalized = work.IsXml ? work.NormalizedXmlDigest : null;
            var kind = !work.IsXml ? (byte)'B' : normalized is not null ? (byte)'X' : (byte)'U';
            hash.AppendData(new[] { kind });
            AppendString(hash, (normalized ?? work.RawBytesDigest)!.Value);
        }
        return new VerificationDigest
        {
            Algorithm = "SHA-256",
            Value = Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant(),
        };
    }

    private static void AssignStableOccurrences(IReadOnlyList<EntryWork> works)
    {
        foreach (var group in works.GroupBy(work => work.Uri, PartNameComparer))
        {
            var occurrence = 0;
            foreach (var work in group
                         .OrderBy(work => work.RawBytesDigest?.Value ?? string.Empty, StringComparer.Ordinal)
                         .ThenBy(work => work.Size)
                         .ThenBy(work => work.Uri, StringComparer.Ordinal)
                         .ThenBy(work => work.ArchiveIndex))
                work.Occurrence = occurrence++;
        }
    }

    // Empty directory-only entries carry no content, so they stay out of both content identities:
    // a repack that adds or drops folder entries is packaging, not a document change. A malformed
    // trailing-slash entry with bytes is retained so its payload cannot disappear from identity.
    private static IEnumerable<EntryWork> StableEntryOrder(IReadOnlyList<EntryWork> works) =>
        works.Where(work => !work.IsDirectory || work.Size != 0)
            .OrderBy(work => work.Uri, StringComparer.Ordinal)
            .ThenBy(work => work.Occurrence);

    private static void FindDuplicateEntryNames(
        IReadOnlyList<EntryWork> works,
        List<VerificationFinding> findings)
    {
        foreach (var group in works.GroupBy(work => work.Uri, PartNameComparer)
                     .Where(group => group.Count() > 1))
        {
            AddFinding(findings, "duplicate_entry", VerificationFindingSeverity.Error,
                "Package contains multiple ZIP entries for the same canonical URI.",
                new ChangeLocation { EntryUri = group.Key });
        }
    }

    private static void FindInterleavedPartNames(
        IReadOnlyList<EntryWork> works,
        List<VerificationFinding> findings)
    {
        var partNames = works
            .Where(work => !work.IsDirectory
                && !PartNameComparer.Equals(work.Uri, ContentTypesUri))
            .Select(work => work.Uri)
            .ToHashSet(PartNameComparer);
        foreach (var partName in partNames.OrderBy(name => name, StringComparer.Ordinal))
        {
            for (var separator = partName.IndexOf('/', 1);
                 separator >= 0;
                 separator = partName.IndexOf('/', separator + 1))
            {
                var prefix = partName[..separator];
                if (!partNames.Contains(prefix))
                    continue;
                AddFinding(findings, "interleaved_part_names", VerificationFindingSeverity.Error,
                    "A part name is derived from another part name by appending path segments.",
                    new ChangeLocation { EntryUri = partName, TargetUri = prefix });
            }
        }
    }

    private static void FindConflictingEntries(
        IReadOnlyList<EntryWork> works,
        List<VerificationFinding> findings)
    {
        foreach (var group in works.GroupBy(work => work.Uri, PartNameComparer)
                     .Where(group => group.Count() > 1))
        {
            if (group.Select(work => work.RawBytesDigest?.Value)
                .Distinct(StringComparer.Ordinal)
                .Skip(1)
                .Any())
            {
                AddFinding(findings, "conflicting_entry", VerificationFindingSeverity.Error,
                    "Duplicate ZIP entries for this URI have different uncompressed bytes.",
                    new ChangeLocation { EntryUri = group.Key });
            }
        }
    }

    private static void ValidateContentTypeTargets(
        ContentTypeMap map,
        IReadOnlySet<string> allEntryUris,
        List<VerificationFinding> findings)
    {
        foreach (var declaration in map.Declarations.Where(declaration => declaration.Kind == "override"))
        {
            if (!allEntryUris.Contains(declaration.Key))
            {
                AddFinding(findings, "missing_content_type_target", VerificationFindingSeverity.Error,
                    "Content-type Override names a part that is absent from the package.",
                    new ChangeLocation { EntryUri = ContentTypesUri, TargetUri = declaration.Key });
            }
        }
    }

    private static bool IsXml(string uri, string? contentType)
    {
        if (PartNameComparer.Equals(uri, ContentTypesUri) || IsRelationshipPart(uri))
            return true;
        if (contentType is not null)
        {
            var essence = MediaTypeEssence(contentType);
            return essence.EndsWith("+xml", StringComparison.OrdinalIgnoreCase)
                || essence.Equals("application/xml", StringComparison.OrdinalIgnoreCase)
                || essence.Equals("text/xml", StringComparison.OrdinalIgnoreCase);
        }
        return uri.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            || uri.EndsWith(".vml", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsKnownOoxmlXml(EntryWork work) =>
        PartNameComparer.Equals(work.Uri, ContentTypesUri)
        || IsRelationshipPart(work.Uri)
        || (work.ContentType is not null
            && KnownOoxmlXmlContentTypes.Contains(MediaTypeEssence(work.ContentType)));

    private static bool IsRelationshipPart(string uri) =>
        XmlSemanticNormalizer.IsRelationshipPart(uri);

    private static string? RelationshipOwner(string relationshipPartUri)
    {
        if (PartNameComparer.Equals(relationshipPartUri, "/_rels/.rels"))
            return "/";
        var marker = relationshipPartUri.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
        if (marker < 0 || !relationshipPartUri.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            return null;
        var directory = relationshipPartUri[..marker];
        var file = relationshipPartUri[(marker + "/_rels/".Length)..^".rels".Length];
        if (file.Length == 0)
            return null;
        var owner = directory + "/" + file;
        return TryCanonicalizePartName(owner, out var canonicalOwner) ? canonicalOwner : null;
    }

    private static string? ResolveRelationshipTarget(
        string owner,
        string target,
        int maximumUriLength,
        out bool invalidTarget)
    {
        invalidTarget = false;
        // A leading slash is a valid package-absolute path. Uri.TryCreate(..., Absolute)
        // interprets it as a file URI and invents the "file" scheme, which made every SDK-authored
        // target such as /word/document.xml fail preflight. Only an RFC 3986 scheme written in the
        // relationship value makes an internal target invalid.
        if (HasUriScheme(target))
        {
            invalidTarget = true;
            return null;
        }

        var delimiter = target.IndexOfAny(['?', '#']);
        var rawPath = delimiter < 0 ? target : target[..delimiter];
        if (!TryCanonicalizeOpcPath(rawPath, requireAbsolute: false, allowRelative: true,
                allowDotSegments: true, out var targetIsAbsolute, out var targetSegments))
        {
            invalidTarget = true;
            return null;
        }

        var ownerDirectoryEnd = owner.LastIndexOf('/');
        var resolvedSegments = targetIsAbsolute || owner == "/" || ownerDirectoryEnd <= 0
            ? new List<string>()
            : owner[1..ownerDirectoryEnd]
                .Split('/', StringSplitOptions.RemoveEmptyEntries).ToList();
        if (targetSegments[^1] is "." or "..")
        {
            invalidTarget = true;
            return null;
        }
        foreach (var segment in targetSegments)
        {
            if (segment == ".")
                continue;
            if (segment == "..")
            {
                if (resolvedSegments.Count == 0)
                {
                    invalidTarget = true;
                    return null;
                }
                resolvedSegments.RemoveAt(resolvedSegments.Count - 1);
                continue;
            }
            resolvedSegments.Add(segment);
        }

        var resolved = "/" + string.Join('/', resolvedSegments);
        if (!TryCanonicalizePartName(resolved, out var canonicalResolved)
            || canonicalResolved.Length > maximumUriLength)
        {
            invalidTarget = true;
            return null;
        }
        return canonicalResolved;
    }

    private static bool HasUriScheme(string value)
    {
        if (value.Length < 2 || !IsAsciiLetter(value[0]))
            return false;
        for (var index = 1; index < value.Length; index++)
        {
            var character = value[index];
            if (character == ':')
                return true;
            if (!(IsAsciiLetter(character) || char.IsAsciiDigit(character)
                    || character is '+' or '-' or '.'))
                return false;
        }
        return false;
    }

    private static bool IsAsciiLetter(char value) =>
        value is >= 'A' and <= 'Z' or >= 'a' and <= 'z';

    private static bool TryCanonicalizeEntryName(string name, out string canonical)
    {
        canonical = "/" + name.Replace('\\', '/').TrimStart('/');

        // The content-types stream is a package metadata item, not an OPC part name. Its square
        // brackets are therefore the one deliberate exception to the isegment-nz grammar.
        if (PartNameComparer.Equals(name, "[Content_Types].xml"))
        {
            canonical = ContentTypesUri;
            return true;
        }

        // A single trailing forward slash marks a directory-only entry. Those are packaging
        // artifacts, not OPC parts, so the grammar is applied to the path they name and the
        // slash is kept in the canonical URI so a folder can never collide with a real part.
        // A trailing backslash is still a malformed path and is left to fail below.
        var isDirectory = name.EndsWith("/", StringComparison.Ordinal);
        var body = isDirectory ? name[..^1] : name;
        if (body.Length == 0 || body.Any(character => character > 0x7f)
            || !TryMapZipItemNameToLogicalName(body, out var logicalBody))
            return false;
        if (!TryCanonicalizeOpcPath(logicalBody, requireAbsolute: false, allowRelative: true,
                allowDotSegments: false, out var isAbsolute, out var segments)
            || isAbsolute)
            return false;
        var joined = "/" + string.Join('/', segments);
        canonical = isDirectory ? joined + "/" : joined;
        return true;
    }

    // ZIP item names are the ASCII physical mapping of logical OPC part names. During the inverse
    // mapping, valid UTF-8 percent sequences for non-ASCII scalars become literal Unicode. ASCII
    // escapes and opaque octets remain escaped for the logical part-name validator to interpret.
    private static bool TryMapZipItemNameToLogicalName(string itemName, out string logicalName)
    {
        var builder = new StringBuilder(itemName.Length);
        for (var index = 0; index < itemName.Length;)
        {
            if (itemName[index] != '%')
            {
                builder.Append(itemName[index++]);
                continue;
            }
            if (!TryReadPercentByte(itemName, index, out var value))
            {
                logicalName = string.Empty;
                return false;
            }
            if (TryDecodeNonAsciiUtf8Escape(itemName, index, out var rune, out var consumed))
            {
                builder.Append(rune.ToString());
                index += consumed;
                continue;
            }
            AppendCanonicalPercentEscape(builder, value);
            index += 3;
        }
        logicalName = builder.ToString();
        return true;
    }

    private static long SumDeclaredSizes(IReadOnlyList<ZipArchiveEntry> entries)
    {
        long total = 0;
        foreach (var entry in entries)
        {
            try
            {
                total = checked(total + entry.Length);
            }
            catch (InvalidDataException)
            {
                // Unreadable metadata is reported per entry while inspecting it.
            }
            catch (OverflowException)
            {
                return long.MaxValue;
            }
        }
        return total;
    }

    private static bool TryCanonicalizePartName(string rawName, out string canonical)
    {
        canonical = rawName;
        if (!TryCanonicalizeOpcPath(rawName, requireAbsolute: true, allowRelative: false,
                allowDotSegments: false, out _, out var segments))
            return false;
        canonical = "/" + string.Join('/', segments);
        return true;
    }

    private static bool TryCanonicalizeOpcPath(
        string rawPath,
        bool requireAbsolute,
        bool allowRelative,
        bool allowDotSegments,
        out bool isAbsolute,
        out List<string> canonicalSegments)
    {
        canonicalSegments = new List<string>();
        isAbsolute = rawPath.StartsWith("/", StringComparison.Ordinal);
        if (rawPath.Length == 0 || rawPath.Contains('\\')
            || (requireAbsolute && !isAbsolute) || (!allowRelative && !isAbsolute)
            || rawPath.StartsWith("//", StringComparison.Ordinal))
            return false;

        var pathBody = isAbsolute ? rawPath[1..] : rawPath;
        if (pathBody.Length == 0)
            return false;
        foreach (var rawSegment in pathBody.Split('/'))
        {
            if (!TryCanonicalizeOpcSegment(rawSegment, allowDotSegments, out var segment))
                return false;
            canonicalSegments.Add(segment);
        }
        return true;
    }

    private static bool TryCanonicalizeOpcSegment(
        string rawSegment,
        bool allowDotSegments,
        out string canonical)
    {
        canonical = string.Empty;
        if (rawSegment.Length == 0)
            return false;
        var builder = new StringBuilder(rawSegment.Length);
        for (var index = 0; index < rawSegment.Length;)
        {
            if (rawSegment[index] == '%')
            {
                if (!TryReadPercentByte(rawSegment, index, out var encoded)
                    || encoded is (byte)'/' or (byte)'\\'
                    || encoded <= 0x7f && IsAsciiUnreserved(encoded))
                    return false;

                var escapeCharactersConsumed = 3;
                if (TryDecodeNonAsciiUtf8Escape(rawSegment, index, out var rune,
                        out var utf8Consumed))
                {
                    // A logical part name must spell RFC 3987 iunreserved characters literally.
                    // Other valid sequences, and opaque bytes such as %FC, remain percent escapes.
                    if (IsIUnreserved(rune))
                        return false;
                    escapeCharactersConsumed = utf8Consumed;
                }
                for (var escape = index;
                     escape < index + escapeCharactersConsumed;
                     escape += 3)
                {
                    TryReadPercentByte(rawSegment, escape, out var escapedByte);
                    AppendCanonicalPercentEscape(builder, escapedByte);
                }
                index += escapeCharactersConsumed;
                continue;
            }

            var status = Rune.DecodeFromUtf16(rawSegment.AsSpan(index), out var literal,
                out var consumed);
            if (status != OperationStatus.Done || !IsIPChar(literal))
                return false;
            builder.Append(rawSegment, index, consumed);
            index += consumed;
        }

        canonical = builder.ToString();
        if (canonical is "." or "..")
            return allowDotSegments;
        return !canonical.EndsWith(".", StringComparison.Ordinal);
    }

    private static bool TryDecodeNonAsciiUtf8Escape(
        string value,
        int offset,
        out Rune rune,
        out int charactersConsumed)
    {
        rune = Rune.ReplacementChar;
        charactersConsumed = 0;
        if (!TryReadPercentByte(value, offset, out var first))
            return false;
        var byteCount = first switch
        {
            >= 0xc2 and <= 0xdf => 2,
            >= 0xe0 and <= 0xef => 3,
            >= 0xf0 and <= 0xf4 => 4,
            _ => 0,
        };
        if (byteCount == 0)
            return false;

        Span<byte> bytes = stackalloc byte[4];
        bytes[0] = first;
        for (var byteIndex = 1; byteIndex < byteCount; byteIndex++)
        {
            if (!TryReadPercentByte(value, offset + byteIndex * 3, out bytes[byteIndex]))
                return false;
        }
        if (Rune.DecodeFromUtf8(bytes[..byteCount], out rune, out var bytesConsumed)
                != OperationStatus.Done
            || bytesConsumed != byteCount || rune.Value <= 0x7f)
        {
            return false;
        }
        charactersConsumed = byteCount * 3;
        return true;
    }

    private static bool TryReadPercentByte(string value, int offset, out byte decoded)
    {
        decoded = 0;
        if (offset < 0 || offset + 2 >= value.Length || value[offset] != '%'
            || !TryHex(value[offset + 1], out var high)
            || !TryHex(value[offset + 2], out var low))
            return false;
        decoded = (byte)((high << 4) | low);
        return true;
    }

    private static void AppendCanonicalPercentEscape(StringBuilder builder, byte value)
    {
        builder.Append('%');
        builder.Append(value.ToString("X2", CultureInfo.InvariantCulture));
    }

    private static bool IsIPChar(Rune value)
    {
        var scalar = value.Value;
        if (scalar > 0x7f)
            return IsUcsChar(scalar);
        return IsAsciiUnreserved((byte)scalar)
            || scalar is '!' or '$' or '&' or '\'' or '(' or ')' or '*' or '+' or ','
                or ';' or '=' or ':' or '@';
    }

    private static bool IsIUnreserved(Rune value) =>
        value.Value <= 0x7f
            ? IsAsciiUnreserved((byte)value.Value)
            : IsUcsChar(value.Value);

    // RFC 3987 ucschar, used by OPC's isegment-nz grammar. Private-use characters and Unicode
    // noncharacters are intentionally outside these ranges.
    private static bool IsUcsChar(int value) =>
        value is >= 0x00a0 and <= 0xd7ff
            or >= 0xf900 and <= 0xfdcf
            or >= 0xfdf0 and <= 0xffef
            or >= 0x10000 and <= 0x1fffd
            or >= 0x20000 and <= 0x2fffd
            or >= 0x30000 and <= 0x3fffd
            or >= 0x40000 and <= 0x4fffd
            or >= 0x50000 and <= 0x5fffd
            or >= 0x60000 and <= 0x6fffd
            or >= 0x70000 and <= 0x7fffd
            or >= 0x80000 and <= 0x8fffd
            or >= 0x90000 and <= 0x9fffd
            or >= 0xa0000 and <= 0xafffd
            or >= 0xb0000 and <= 0xbfffd
            or >= 0xc0000 and <= 0xcfffd
            or >= 0xd0000 and <= 0xdfffd
            or >= 0xe1000 and <= 0xefffd;

    private static bool TryHex(char value, out int nibble)
    {
        if (value is >= '0' and <= '9')
        {
            nibble = value - '0';
            return true;
        }
        if (value is >= 'a' and <= 'f')
        {
            nibble = value - 'a' + 10;
            return true;
        }
        if (value is >= 'A' and <= 'F')
        {
            nibble = value - 'A' + 10;
            return true;
        }
        nibble = 0;
        return false;
    }

    private static bool IsAsciiUnreserved(byte value) =>
        value is >= (byte)'A' and <= (byte)'Z'
        || value is >= (byte)'a' and <= (byte)'z'
        || value is >= (byte)'0' and <= (byte)'9'
        || value is (byte)'-' or (byte)'.' or (byte)'_' or (byte)'~';

    private static IReadOnlyList<ZipCentralEntryMetadata>? TryReadCentralDirectoryMetadata(
        byte[] bytes,
        int expectedCount)
    {
        const uint eocdSignature = 0x06054b50;
        const uint zip64EocdSignature = 0x06064b50;
        const uint zip64LocatorSignature = 0x07064b50;
        const uint centralSignature = 0x02014b50;
        if (bytes.Length < 22)
            return null;
        var minimum = Math.Max(0, bytes.Length - (22 + ushort.MaxValue));
        var eocd = -1;
        for (var offset = bytes.Length - 22; offset >= minimum; offset--)
        {
            if (ReadUInt32(bytes, offset) == eocdSignature
                && offset + 22 + ReadUInt16(bytes, offset + 20) == bytes.Length)
            {
                eocd = offset;
                break;
            }
        }
        if (eocd < 0)
            return null;

        var diskNumber = ReadUInt16(bytes, eocd + 4);
        var centralDisk = ReadUInt16(bytes, eocd + 6);
        var entriesOnDisk16 = ReadUInt16(bytes, eocd + 8);
        var entryCount16 = ReadUInt16(bytes, eocd + 10);
        var centralSize32 = ReadUInt32(bytes, eocd + 12);
        var centralOffset32 = ReadUInt32(bytes, eocd + 16);
        ulong entryCount;
        ulong centralSize;
        ulong centralOffset;
        ulong centralBoundary = (ulong)eocd;
        var zip64 = entriesOnDisk16 == ushort.MaxValue || entryCount16 == ushort.MaxValue
            || centralSize32 == uint.MaxValue || centralOffset32 == uint.MaxValue;
        if (zip64)
        {
            var locator = eocd - 20;
            if (locator < 0 || ReadUInt32(bytes, locator) != zip64LocatorSignature
                || ReadUInt32(bytes, locator + 4) != 0
                || ReadUInt32(bytes, locator + 16) != 1)
                return null;
            var zip64Offset = ReadUInt64(bytes, locator + 8);
            if (zip64Offset > int.MaxValue)
                return null;
            var record = (int)zip64Offset;
            if (record < 0 || record + 56 > bytes.Length
                || ReadUInt32(bytes, record) != zip64EocdSignature)
                return null;
            var recordSize = ReadUInt64(bytes, record + 4);
            if (recordSize < 44 || recordSize > int.MaxValue
                || (ulong)record + 12 + recordSize != (ulong)locator
                || ReadUInt32(bytes, record + 16) != 0
                || ReadUInt32(bytes, record + 20) != 0)
                return null;
            centralBoundary = (ulong)record;
            var entriesOnDisk = ReadUInt64(bytes, record + 24);
            entryCount = ReadUInt64(bytes, record + 32);
            centralSize = ReadUInt64(bytes, record + 40);
            centralOffset = ReadUInt64(bytes, record + 48);
            if (entriesOnDisk != entryCount)
                return null;
        }
        else
        {
            if (diskNumber != 0 || centralDisk != 0 || entriesOnDisk16 != entryCount16)
                return null;
            entryCount = entryCount16;
            centralSize = centralSize32;
            centralOffset = centralOffset32;
        }

        if (entryCount != (ulong)expectedCount || entryCount > int.MaxValue
            || centralOffset > int.MaxValue || centralSize > int.MaxValue
            || centralOffset + centralSize > centralBoundary)
            return null;

        var count = (int)entryCount;
        var metadata = new List<ZipCentralEntryMetadata>(count);
        var position = (int)centralOffset;
        var centralEnd = centralOffset + centralSize;
        for (var index = 0; index < count; index++)
        {
            if (position < 0 || position + 46 > bytes.Length
                || (ulong)(position + 46) > centralEnd
                || ReadUInt32(bytes, position) != centralSignature)
                return null;
            metadata.Add(new ZipCentralEntryMetadata(
                IsEncrypted: (ReadUInt16(bytes, position + 8) & 1) != 0,
                Crc32: ReadUInt32(bytes, position + 16)));
            var nameLength = ReadUInt16(bytes, position + 28);
            var extraLength = ReadUInt16(bytes, position + 30);
            var commentLength = ReadUInt16(bytes, position + 32);
            var nextPosition = (long)position + 46 + nameLength + extraLength + commentLength;
            if (nextPosition > bytes.Length || nextPosition > int.MaxValue
                || (ulong)nextPosition > centralEnd)
                return null;
            position = (int)nextPosition;
        }
        return metadata;
    }

    private static void ValidateCrc32(
        EntryWork work,
        uint actualCrc32,
        List<VerificationFinding> findings)
    {
        if (work.ExpectedCrc32 is not { } expectedCrc32 || actualCrc32 == expectedCrc32)
            return;
        AddFinding(findings, "crc_mismatch", VerificationFindingSeverity.Error,
            "Entry payload CRC-32 does not match the ZIP central-directory value.",
            new ChangeLocation { EntryUri = work.Uri });
    }

    private static uint ComputeCrc32(ReadOnlySpan<byte> bytes) =>
        ~AppendCrc32(uint.MaxValue, bytes);

    private static uint AppendCrc32(uint state, ReadOnlySpan<byte> bytes)
    {
        foreach (var value in bytes)
            state = Crc32Table[(state ^ value) & 0xff] ^ (state >> 8);
        return state;
    }

    private static uint[] CreateCrc32Table()
    {
        var table = new uint[256];
        for (uint index = 0; index < table.Length; index++)
        {
            var value = index;
            for (var bit = 0; bit < 8; bit++)
                value = (value & 1) != 0 ? 0xedb88320U ^ (value >> 1) : value >> 1;
            table[index] = value;
        }
        return table;
    }

    private static ushort ReadUInt16(byte[] bytes, int offset) =>
        offset >= 0 && offset + sizeof(ushort) <= bytes.Length
            ? BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset, sizeof(ushort)))
            : (ushort)0;

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        offset >= 0 && offset + sizeof(uint) <= bytes.Length
            ? BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, sizeof(uint)))
            : 0;

    private static ulong ReadUInt64(byte[] bytes, int offset) =>
        offset >= 0 && offset + sizeof(ulong) <= bytes.Length
            ? BinaryPrimitives.ReadUInt64LittleEndian(bytes.AsSpan(offset, sizeof(ulong)))
            : 0;

    private static byte[] ReadAllBounded(
        ZipArchiveEntry entry,
        long xmlMaximum,
        long entryMaximum,
        long expansionMaximum,
        ActualReadBudget readBudget)
    {
        using var input = entry.Open();
        using var output = new MemoryStream((int)Math.Min(entry.Length, int.MaxValue));
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
        {
            if (read > expansionMaximum - total)
                throw new ManifestSafetyException(SafetyLimitKind.EntryExpansion);
            if (read > entryMaximum - total)
                throw new ManifestSafetyException(SafetyLimitKind.EntrySize);
            if (!readBudget.TryConsume(read))
                throw new ManifestSafetyException(SafetyLimitKind.TotalExpansion);
            if (read > xmlMaximum - total)
                throw new ManifestSafetyException(SafetyLimitKind.XmlSize);
            total += read;
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

    private static long ExpansionCeiling(long compressedSize, double maximumRatio)
    {
        if (compressedSize <= 0)
            return 0;
        var product = compressedSize * maximumRatio;
        return double.IsInfinity(product) || product >= long.MaxValue
            ? long.MaxValue
            : (long)Math.Floor(product);
    }

    private static long XmlCharacterLimit(long maximumXmlBytes) =>
        maximumXmlBytes > long.MaxValue / 2
            ? long.MaxValue
            : Math.Max(maximumXmlBytes * 2, 1);

    private static string SafetyFindingCode(SafetyLimitKind kind) => kind switch
    {
        SafetyLimitKind.EntryExpansion => "compression_ratio_limit_exceeded",
        SafetyLimitKind.EntrySize => "entry_size_limit_exceeded",
        SafetyLimitKind.XmlSize => "xml_size_limit_exceeded",
        _ => "entry_expansion_limit_exceeded",
    };

    private static string SafetyFindingMessage(SafetyLimitKind kind, string subject) => kind switch
    {
        SafetyLimitKind.EntryExpansion =>
            $"{subject} produced more actual bytes than compressed size × MaxCompressionRatio permits.",
        SafetyLimitKind.EntrySize =>
            $"{subject} produced more actual bytes than MaxEntryUncompressedBytes permits.",
        SafetyLimitKind.XmlSize => $"{subject} exceeded the configured XML parsing limit.",
        _ => $"{subject} exceeded the remaining total package expansion budget.",
    };

    private static VerificationDigest Digest(byte[] bytes) => new()
    {
        Algorithm = "SHA-256",
        Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
    };

    private static void AppendString(IncrementalHash hash, string value)
    {
        var bytes = Encoding.UTF8.GetBytes(value);
        AppendInt32(hash, bytes.Length);
        hash.AppendData(bytes);
    }

    private static void AppendInt32(IncrementalHash hash, int value)
    {
        Span<byte> bytes = stackalloc byte[sizeof(int)];
        BinaryPrimitives.WriteInt32LittleEndian(bytes, value);
        hash.AppendData(bytes);
    }

    private static void AppendInt64(IncrementalHash hash, long value)
    {
        Span<byte> bytes = stackalloc byte[sizeof(long)];
        BinaryPrimitives.WriteInt64LittleEndian(bytes, value);
        hash.AppendData(bytes);
    }

    private static string? ElementNamespaceAttribute(XElement element, string localName) =>
        element.Attribute(XName.Get(localName, element.Name.NamespaceName))?.Value;

    private static string? UnqualifiedAttribute(XElement element, string localName) =>
        element.Attribute(XName.Get(localName))?.Value;

    private static bool IsWordprocessingElement(XElement element) =>
        element.Name.NamespaceName is TransitionalWordNamespace or StrictWordNamespace;

    private static bool IsWordprocessingPart(EntryWork work)
        => work.ContentType is not null
            && WordprocessingXmlContentTypes.Contains(MediaTypeEssence(work.ContentType));

    private static bool IsWordprocessingStoryPart(EntryWork work) =>
        IsWordprocessingMainPart(work)
        || IsContentType(work, "header")
        || IsContentType(work, "footer")
        || IsContentType(work, "footnotes")
        || IsContentType(work, "endnotes")
        || IsContentType(work, "comments")
        || IsContentType(work, "glossaryDocument");

    private static bool IsWordprocessingMainPart(EntryWork work) =>
        IsMime(work, "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
        || IsMime(work, "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml")
        || IsMime(work, "application/vnd.ms-word.document.macroEnabled.main+xml")
        || IsMime(work, "application/vnd.ms-word.template.macroEnabledTemplate.main+xml");

    private static bool IsPropertyRevisionName(string localName) => localName is
        "numberingChange" or "pPrChange" or "rPrChange" or "sectPrChange"
        or "tblGridChange" or "tblPrChange" or "tblPrExChange" or "tcPrChange"
        or "trPrChange";

    private static bool IsCustomXmlRevisionRangeStart(string localName) => localName is
        "customXmlInsRangeStart" or "customXmlDelRangeStart"
        or "customXmlMoveFromRangeStart" or "customXmlMoveToRangeStart";

    private static bool IsOfficeRelationshipNamespace(string value) =>
        value is TransitionalOfficeRelationshipNamespace or StrictOfficeRelationshipNamespace;

    private static bool IsPackageRelationshipsNamespace(string value) =>
        value == TransitionalPackageRelationshipsNamespace;

    private static bool IsOfficeRelationshipType(string value, string localType) =>
        value.Equals(TransitionalOfficeRelationshipTypePrefix + localType, StringComparison.Ordinal)
        || value.Equals(StrictOfficeRelationshipTypePrefix + localType, StringComparison.Ordinal);

    private static int CountPositiveDefinitions(IEnumerable<XElement> elements, string localName) =>
        elements.Count(element => element.Name.LocalName == localName
            && int.TryParse(ElementNamespaceAttribute(element, "id"), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var id)
            && id > 0);

    private static bool IsContentType(EntryWork work, string token) =>
        IsMime(work, "application/vnd.openxmlformats-officedocument.wordprocessingml."
            + token + "+xml");

    private static bool IsMime(EntryWork work, string contentType) =>
        work.ContentType is { } declared
        && string.Equals(MediaTypeEssence(declared), contentType,
            StringComparison.OrdinalIgnoreCase);

    private static bool IsMediaTypeFamily(string contentType, string topLevelType) =>
        MediaTypeEssence(contentType).StartsWith(
            topLevelType + "/", StringComparison.OrdinalIgnoreCase);

    private static string MediaTypeEssence(string contentType)
    {
        var parameter = contentType.IndexOf(';');
        return (parameter < 0 ? contentType : contentType[..parameter]).TrimEnd(' ', '\t');
    }

    private static bool IsCommentsPart(EntryWork work) =>
        IsContentType(work, "comments");

    private static bool IsCustomXmlDataPart(EntryWork work)
    {
        if (!work.Uri.StartsWith("/customXml/", StringComparison.OrdinalIgnoreCase))
            return false;
        var fileName = work.Uri[(work.Uri.LastIndexOf('/') + 1)..];
        return fileName.StartsWith("item", StringComparison.OrdinalIgnoreCase)
            && !fileName.StartsWith("itemProps", StringComparison.OrdinalIgnoreCase)
            && fileName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsOn(string? value) => value is not null
        && (value.Equals("1", StringComparison.OrdinalIgnoreCase)
            || value.Equals("true", StringComparison.OrdinalIgnoreCase)
            || value.Equals("on", StringComparison.OrdinalIgnoreCase));

    private static bool IsValidContentTypeExtension(string value, int maximumLength)
    {
        if (value.Length == 0 || value.Length > maximumLength)
            return false;
        foreach (var character in value)
        {
            if (character is '.' or '/' or '\\' or '?' or '#'
                || char.IsControl(character) || char.IsWhiteSpace(character))
                return false;
        }
        return true;
    }

    private static bool IsValidContentTypeMediaType(string value)
    {
        // OPC ContentType values are media types, not arbitrary labels. Keep the parser local and
        // bounded: MIME tokens are ASCII and optional parameters use the RFC quoted-string shape.
        // Leading/trailing whitespace is significant in an XML attribute and is never accepted.
        if (value.Length == 0
            || char.IsWhiteSpace(value[0])
            || char.IsWhiteSpace(value[^1]))
        {
            return false;
        }

        var offset = 0;
        if (!ConsumeMediaTypeToken(value, ref offset)
            || offset >= value.Length || value[offset++] != '/'
            || !ConsumeMediaTypeToken(value, ref offset))
        {
            return false;
        }

        while (offset < value.Length)
        {
            ConsumeOptionalWhitespace(value, ref offset);
            if (offset >= value.Length || value[offset++] != ';')
                return false;
            ConsumeOptionalWhitespace(value, ref offset);
            if (!ConsumeMediaTypeToken(value, ref offset))
                return false;
            if (offset >= value.Length || value[offset++] != '=')
                return false;
            if (offset >= value.Length)
                return false;

            if (value[offset] == '"')
            {
                offset++;
                var closed = false;
                while (offset < value.Length)
                {
                    var character = value[offset++];
                    if (character == '"')
                    {
                        closed = true;
                        break;
                    }
                    if (character == '\\')
                    {
                        if (offset >= value.Length || !IsQuotedMediaTypeCharacter(value[offset++]))
                            return false;
                    }
                    else if (!IsQuotedMediaTypeCharacter(character) || character == '\\')
                    {
                        return false;
                    }
                }
                if (!closed)
                    return false;
            }
            else if (!ConsumeMediaTypeToken(value, ref offset))
            {
                return false;
            }
        }

        return true;
    }

    private static bool ConsumeMediaTypeToken(string value, ref int offset)
    {
        var start = offset;
        while (offset < value.Length && IsMediaTypeTokenCharacter(value[offset]))
            offset++;
        return offset > start;
    }

    private static bool IsMediaTypeTokenCharacter(char character) =>
        character is >= '0' and <= '9'
        or >= 'A' and <= 'Z'
        or >= 'a' and <= 'z'
        or '!' or '#' or '$' or '%' or '&' or '\'' or '*' or '+' or '-' or '.'
        or '^' or '_' or '`' or '|' or '~';

    private static bool IsQuotedMediaTypeCharacter(char character) =>
        character is '\t' or ' '
        || character is >= (char)0x21 and <= (char)0x7e;

    private static void ConsumeOptionalWhitespace(string value, ref int offset)
    {
        while (offset < value.Length && value[offset] is ' ' or '\t')
            offset++;
    }

    private const uint CfbDifSector = 0xfffffffc;
    private const uint CfbFatSector = 0xfffffffd;
    private const uint CfbEndOfChain = 0xfffffffe;
    private const uint CfbFreeSector = 0xffffffff;
    private const uint CfbNoStream = 0xffffffff;

    private static bool ContainsEncryptedOoxmlStreams(byte[] bytes) =>
        TryReadCompoundFileDirectory(bytes, out var hasEncryptedPackage, out var hasEncryptionInfo)
        && hasEncryptedPackage && hasEncryptionInfo;

    private static bool TryReadCompoundFileDirectory(
        ReadOnlySpan<byte> bytes,
        out bool hasEncryptedPackage,
        out bool hasEncryptionInfo)
    {
        hasEncryptedPackage = false;
        hasEncryptionInfo = false;

        if (bytes.Length < 512 || !bytes.StartsWith(OleSignature)
            || BinaryPrimitives.ReadUInt16LittleEndian(bytes[24..]) != 0x003e
            || BinaryPrimitives.ReadUInt16LittleEndian(bytes[28..]) != 0xfffe
            || BinaryPrimitives.ReadUInt16LittleEndian(bytes[32..]) != 6
            || BinaryPrimitives.ReadUInt32LittleEndian(bytes[56..]) != 4096
            || !AllZero(bytes[8..24]) || !AllZero(bytes[34..40]))
        {
            return false;
        }

        var majorVersion = BinaryPrimitives.ReadUInt16LittleEndian(bytes[26..]);
        var sectorShift = BinaryPrimitives.ReadUInt16LittleEndian(bytes[30..]);
        if ((majorVersion != 3 || sectorShift != 9)
            && (majorVersion != 4 || sectorShift != 12))
        {
            return false;
        }

        var sectorSize = 1 << sectorShift;
        if (bytes.Length < sectorSize || bytes.Length % sectorSize != 0)
            return false;
        var sectorCount = bytes.Length / sectorSize - 1;
        if (sectorCount <= 0)
            return false;

        var declaredDirectorySectorCount = BinaryPrimitives.ReadUInt32LittleEndian(bytes[40..]);
        if ((majorVersion == 3 && declaredDirectorySectorCount != 0)
            || declaredDirectorySectorCount > (uint)sectorCount)
        {
            return false;
        }

        var fatSectorCount = BinaryPrimitives.ReadUInt32LittleEndian(bytes[44..]);
        var firstDirectorySector = BinaryPrimitives.ReadUInt32LittleEndian(bytes[48..]);
        var firstMiniFatSector = BinaryPrimitives.ReadUInt32LittleEndian(bytes[60..]);
        var miniFatSectorCount = BinaryPrimitives.ReadUInt32LittleEndian(bytes[64..]);
        var firstDifatSector = BinaryPrimitives.ReadUInt32LittleEndian(bytes[68..]);
        var difatSectorCount = BinaryPrimitives.ReadUInt32LittleEndian(bytes[72..]);
        if (fatSectorCount == 0 || fatSectorCount > (uint)sectorCount
            || !IsRegularCfbSector(firstDirectorySector, sectorCount)
            || miniFatSectorCount > (uint)sectorCount
            || difatSectorCount > (uint)sectorCount
            || (miniFatSectorCount == 0 && firstMiniFatSector != CfbEndOfChain)
            || (miniFatSectorCount != 0 && !IsRegularCfbSector(firstMiniFatSector, sectorCount)))
        {
            return false;
        }

        var fatSectors = new List<uint>((int)fatSectorCount);
        var difatSectors = new List<uint>((int)difatSectorCount);
        var specialSectors = new HashSet<uint>();
        for (var index = 0; index < 109; index++)
        {
            var sector = BinaryPrimitives.ReadUInt32LittleEndian(bytes[(76 + index * 4)..]);
            if ((uint)fatSectors.Count < fatSectorCount)
            {
                if (!TryAddCfbSector(sector, sectorCount, fatSectors, specialSectors))
                    return false;
            }
            else if (sector != CfbFreeSector)
            {
                return false;
            }
        }

        if ((uint)fatSectors.Count == fatSectorCount)
        {
            if (difatSectorCount != 0 || firstDifatSector != CfbEndOfChain)
                return false;
        }
        else
        {
            if (difatSectorCount == 0 || !IsRegularCfbSector(firstDifatSector, sectorCount))
                return false;
            var difatSector = firstDifatSector;
            var entriesPerDifatSector = sectorSize / sizeof(uint) - 1;
            for (uint chainIndex = 0; chainIndex < difatSectorCount; chainIndex++)
            {
                if (!IsRegularCfbSector(difatSector, sectorCount)
                    || !specialSectors.Add(difatSector))
                {
                    return false;
                }
                difatSectors.Add(difatSector);

                var sectorBytes = CompoundFileSector(bytes, sectorSize, difatSector);
                for (var index = 0; index < entriesPerDifatSector; index++)
                {
                    var fatSector = BinaryPrimitives.ReadUInt32LittleEndian(
                        sectorBytes[(index * sizeof(uint))..]);
                    if ((uint)fatSectors.Count < fatSectorCount)
                    {
                        if (!TryAddCfbSector(
                            fatSector, sectorCount, fatSectors, specialSectors))
                        {
                            return false;
                        }
                    }
                    else if (fatSector != CfbFreeSector)
                    {
                        return false;
                    }
                }

                var next = BinaryPrimitives.ReadUInt32LittleEndian(
                    sectorBytes[(entriesPerDifatSector * sizeof(uint))..]);
                if (chainIndex + 1 == difatSectorCount)
                {
                    if (next != CfbEndOfChain)
                        return false;
                }
                else if (!IsRegularCfbSector(next, sectorCount))
                {
                    return false;
                }
                difatSector = next;
            }
            if ((uint)fatSectors.Count != fatSectorCount)
                return false;
        }

        foreach (var fatSector in fatSectors)
        {
            if (!TryReadCompoundFileFatEntry(
                    bytes, sectorSize, sectorCount, fatSectors, fatSector, out var marker)
                || marker != CfbFatSector)
            {
                return false;
            }
        }
        foreach (var difatSector in difatSectors)
        {
            if (!TryReadCompoundFileFatEntry(
                    bytes, sectorSize, sectorCount, fatSectors, difatSector, out var marker)
                || marker != CfbDifSector)
            {
                return false;
            }
        }

        var directorySectors = new HashSet<uint>();
        var currentDirectorySector = firstDirectorySector;
        var directorySectorCount = 0;
        var sawRootEntry = false;
        var directoryEntries = new List<CompoundFileDirectoryEntry?>();
        while (currentDirectorySector != CfbEndOfChain)
        {
            if (!IsRegularCfbSector(currentDirectorySector, sectorCount)
                || specialSectors.Contains(currentDirectorySector)
                || !directorySectors.Add(currentDirectorySector))
            {
                return false;
            }

            directorySectorCount++;
            var sectorBytes = CompoundFileSector(bytes, sectorSize, currentDirectorySector);
            for (var offset = 0; offset < sectorBytes.Length; offset += 128)
            {
                var entry = sectorBytes.Slice(offset, 128);
                var objectType = entry[66];
                if (objectType == 0)
                {
                    directoryEntries.Add(null);
                    continue;
                }
                if (objectType is not (1 or 2 or 5)
                    || entry[67] > 1
                    || !TryReadCompoundFileDirectoryName(entry, out var name))
                {
                    return false;
                }

                var isFirstDirectoryEntry = directorySectorCount == 1 && offset == 0;
                if (isFirstDirectoryEntry)
                {
                    if (objectType != 5 || name != "Root Entry")
                        return false;
                    sawRootEntry = true;
                }
                else if (objectType == 5)
                {
                    return false;
                }

                var streamSize = BinaryPrimitives.ReadUInt64LittleEndian(entry[120..]);
                if (majorVersion == 3 && streamSize > uint.MaxValue)
                    return false;
                var startingSector = BinaryPrimitives.ReadUInt32LittleEndian(entry[116..]);
                if (objectType == 2
                    && ((streamSize == 0 && startingSector != CfbEndOfChain)
                        || (streamSize != 0 && startingSector >= CfbDifSector)))
                {
                    return false;
                }
                directoryEntries.Add(new CompoundFileDirectoryEntry
                {
                    Name = name,
                    ObjectType = objectType,
                    LeftSibling = BinaryPrimitives.ReadUInt32LittleEndian(entry[68..]),
                    RightSibling = BinaryPrimitives.ReadUInt32LittleEndian(entry[72..]),
                    Child = BinaryPrimitives.ReadUInt32LittleEndian(entry[76..]),
                    StartingSector = startingSector,
                    StreamSize = streamSize,
                });
            }

            if (!TryReadCompoundFileFatEntry(bytes, sectorSize, sectorCount, fatSectors,
                    currentDirectorySector, out currentDirectorySector)
                || currentDirectorySector is CfbFreeSector or CfbFatSector or CfbDifSector)
            {
                return false;
            }
        }

        if (!sawRootEntry
            || (majorVersion == 4
                && declaredDirectorySectorCount != (uint)directorySectorCount)
            || directoryEntries.Count == 0 || directoryEntries[0] is not { } root
            || root.LeftSibling != CfbNoStream || root.RightSibling != CfbNoStream)
        {
            return false;
        }

        if (!TryReadCompoundFileMiniStreamAllocation(
                bytes, sectorSize, sectorCount, fatSectors, specialSectors,
                directorySectors, firstMiniFatSector, miniFatSectorCount, root,
                out var miniFatSectors, out var rootMiniStreamSize,
                out var unavailableRegularSectors))
        {
            return false;
        }

        var seenEncryptedPackageEntry = false;
        var seenEncryptionInfoEntry = false;
        var allocatedEncryptionMiniSectors = new HashSet<uint>();
        var allocatedEncryptionRegularSectors = new HashSet<uint>();
        var directoryTree = new Stack<uint>();
        var visitedDirectoryEntries = new HashSet<uint>();
        if (root.Child != CfbNoStream)
            directoryTree.Push(root.Child);
        while (directoryTree.Count != 0)
        {
            var streamId = directoryTree.Pop();
            if (streamId >= (uint)directoryEntries.Count
                || directoryEntries[(int)streamId] is not { } entry
                || !visitedDirectoryEntries.Add(streamId))
            {
                return false;
            }
            if (entry.LeftSibling != CfbNoStream)
                directoryTree.Push(entry.LeftSibling);
            if (entry.RightSibling != CfbNoStream)
                directoryTree.Push(entry.RightSibling);
            if (entry.ObjectType != 2)
                continue;
            if (entry.Child != CfbNoStream)
                return false;

            if (entry.Name.Equals("EncryptedPackage", StringComparison.OrdinalIgnoreCase))
            {
                if (seenEncryptedPackageEntry)
                    return false;
                seenEncryptedPackageEntry = true;
                hasEncryptedPackage = IsValidCompoundFileEncryptionStream(
                    bytes, sectorSize, sectorCount, fatSectors, unavailableRegularSectors,
                    miniFatSectors, rootMiniStreamSize, allocatedEncryptionMiniSectors,
                    allocatedEncryptionRegularSectors, entry);
            }
            else if (entry.Name.Equals("EncryptionInfo", StringComparison.OrdinalIgnoreCase))
            {
                if (seenEncryptionInfoEntry)
                    return false;
                seenEncryptionInfoEntry = true;
                hasEncryptionInfo = IsValidCompoundFileEncryptionStream(
                    bytes, sectorSize, sectorCount, fatSectors, unavailableRegularSectors,
                    miniFatSectors, rootMiniStreamSize, allocatedEncryptionMiniSectors,
                    allocatedEncryptionRegularSectors, entry);
            }
        }

        return true;
    }

    private static bool IsValidCompoundFileEncryptionStream(
        ReadOnlySpan<byte> bytes,
        int sectorSize,
        int sectorCount,
        IReadOnlyList<uint> fatSectors,
        IReadOnlySet<uint> unavailableRegularSectors,
        IReadOnlyList<uint> miniFatSectors,
        ulong rootMiniStreamSize,
        HashSet<uint> allocatedEncryptionMiniSectors,
        HashSet<uint> allocatedEncryptionRegularSectors,
        CompoundFileDirectoryEntry entry)
    {
        if (entry.StreamSize < sizeof(ulong))
            return false;
        if (entry.StreamSize < 4096)
        {
            const ulong miniSectorSize = 64;
            var expectedMiniSectors = entry.StreamSize / miniSectorSize
                + (entry.StreamSize % miniSectorSize == 0 ? 0UL : 1UL);
            var finalMiniSectorBytes = entry.StreamSize % miniSectorSize;
            var currentMiniSector = entry.StartingSector;
            var streamMiniSectors = new HashSet<uint>();
            for (ulong index = 0; index < expectedMiniSectors; index++)
            {
                var bytesInMiniSector = index + 1 == expectedMiniSectors
                    && finalMiniSectorBytes != 0
                    ? finalMiniSectorBytes
                    : miniSectorSize;
                var rootOffset = (ulong)currentMiniSector * miniSectorSize;
                if (rootOffset >= rootMiniStreamSize
                    || bytesInMiniSector > rootMiniStreamSize - rootOffset
                    || !streamMiniSectors.Add(currentMiniSector)
                    || allocatedEncryptionMiniSectors.Contains(currentMiniSector)
                    || !TryReadCompoundFileMiniFatEntry(
                        bytes, sectorSize, miniFatSectors, currentMiniSector,
                        out var nextMiniSector))
                {
                    return false;
                }
                if (index + 1 == expectedMiniSectors)
                {
                    if (nextMiniSector != CfbEndOfChain)
                        return false;
                }
                else
                {
                    currentMiniSector = nextMiniSector;
                }
            }

            allocatedEncryptionMiniSectors.UnionWith(streamMiniSectors);
            return true;
        }

        var expectedSectors = entry.StreamSize / (ulong)sectorSize
            + (entry.StreamSize % (ulong)sectorSize == 0 ? 0UL : 1UL);
        if (!TryReadCompoundFileRegularSectorChain(
                bytes, sectorSize, sectorCount, fatSectors, unavailableRegularSectors,
                entry.StartingSector, expectedSectors, out var streamSectors)
            || streamSectors.Any(allocatedEncryptionRegularSectors.Contains))
        {
            return false;
        }
        allocatedEncryptionRegularSectors.UnionWith(streamSectors);
        return true;
    }

    private static bool TryReadCompoundFileMiniStreamAllocation(
        ReadOnlySpan<byte> bytes,
        int sectorSize,
        int sectorCount,
        IReadOnlyList<uint> fatSectors,
        IReadOnlySet<uint> specialSectors,
        IReadOnlySet<uint> directorySectors,
        uint firstMiniFatSector,
        uint miniFatSectorCount,
        CompoundFileDirectoryEntry root,
        out List<uint> miniFatSectors,
        out ulong rootMiniStreamSize,
        out HashSet<uint> unavailableRegularSectors)
    {
        miniFatSectors = new List<uint>();
        rootMiniStreamSize = 0;
        unavailableRegularSectors = new HashSet<uint>(specialSectors);
        unavailableRegularSectors.UnionWith(directorySectors);

        if (!TryReadCompoundFileRegularSectorChain(
                bytes, sectorSize, sectorCount, fatSectors, unavailableRegularSectors,
                firstMiniFatSector, miniFatSectorCount, out miniFatSectors))
        {
            return false;
        }
        unavailableRegularSectors.UnionWith(miniFatSectors);

        var expectedRootSectors = root.StreamSize / (ulong)sectorSize
            + (root.StreamSize % (ulong)sectorSize == 0 ? 0UL : 1UL);
        if (!TryReadCompoundFileRegularSectorChain(
                bytes, sectorSize, sectorCount, fatSectors, unavailableRegularSectors,
                root.StartingSector, expectedRootSectors, out var rootMiniStreamSectors))
        {
            return false;
        }
        unavailableRegularSectors.UnionWith(rootMiniStreamSectors);

        rootMiniStreamSize = root.StreamSize;
        return true;
    }

    private static bool TryReadCompoundFileRegularSectorChain(
        ReadOnlySpan<byte> bytes,
        int sectorSize,
        int sectorCount,
        IReadOnlyList<uint> fatSectors,
        IReadOnlySet<uint> unavailableSectors,
        uint startingSector,
        ulong expectedSectorCount,
        out List<uint> chain)
    {
        chain = new List<uint>();
        if (expectedSectorCount == 0)
            return startingSector == CfbEndOfChain;
        if (expectedSectorCount > (ulong)sectorCount)
            return false;

        var currentSector = startingSector;
        var visitedSectors = new HashSet<uint>();
        for (ulong index = 0; index < expectedSectorCount; index++)
        {
            if (!IsRegularCfbSector(currentSector, sectorCount)
                || unavailableSectors.Contains(currentSector)
                || !visitedSectors.Add(currentSector)
                || !TryReadCompoundFileFatEntry(bytes, sectorSize, sectorCount, fatSectors,
                    currentSector, out var nextSector))
            {
                return false;
            }
            chain.Add(currentSector);
            if (index + 1 == expectedSectorCount)
                return nextSector == CfbEndOfChain;
            currentSector = nextSector;
        }
        return false;
    }

    private static bool TryReadCompoundFileMiniFatEntry(
        ReadOnlySpan<byte> bytes,
        int sectorSize,
        IReadOnlyList<uint> miniFatSectors,
        uint miniSector,
        out uint value)
    {
        value = CfbFreeSector;
        var entriesPerMiniFatSector = sectorSize / sizeof(uint);
        var miniFatSectorIndex = (ulong)miniSector / (uint)entriesPerMiniFatSector;
        if (miniFatSectorIndex >= (ulong)miniFatSectors.Count)
            return false;
        var entryIndex = (int)(miniSector % (uint)entriesPerMiniFatSector);
        var miniFatBytes = CompoundFileSector(
            bytes, sectorSize, miniFatSectors[(int)miniFatSectorIndex]);
        value = BinaryPrimitives.ReadUInt32LittleEndian(
            miniFatBytes[(entryIndex * sizeof(uint))..]);
        return true;
    }

    private static bool TryAddCfbSector(
        uint sector,
        int sectorCount,
        List<uint> sectors,
        HashSet<uint> usedSectors)
    {
        if (!IsRegularCfbSector(sector, sectorCount) || !usedSectors.Add(sector))
            return false;
        sectors.Add(sector);
        return true;
    }

    private static bool TryReadCompoundFileFatEntry(
        ReadOnlySpan<byte> bytes,
        int sectorSize,
        int sectorCount,
        IReadOnlyList<uint> fatSectors,
        uint sector,
        out uint value)
    {
        value = CfbFreeSector;
        if (!IsRegularCfbSector(sector, sectorCount))
            return false;
        var entriesPerFatSector = sectorSize / sizeof(uint);
        var fatSectorIndex = (int)(sector / (uint)entriesPerFatSector);
        if (fatSectorIndex >= fatSectors.Count)
            return false;
        var entryIndex = (int)(sector % (uint)entriesPerFatSector);
        var fatBytes = CompoundFileSector(bytes, sectorSize, fatSectors[fatSectorIndex]);
        value = BinaryPrimitives.ReadUInt32LittleEndian(fatBytes[(entryIndex * sizeof(uint))..]);
        return true;
    }

    private static bool TryReadCompoundFileDirectoryName(
        ReadOnlySpan<byte> entry,
        out string name)
    {
        name = string.Empty;
        var byteLength = BinaryPrimitives.ReadUInt16LittleEndian(entry[64..]);
        if (byteLength is < 2 or > 64 || (byteLength & 1) != 0
            || BinaryPrimitives.ReadUInt16LittleEndian(entry[(byteLength - 2)..]) != 0)
        {
            return false;
        }
        for (var offset = 0; offset < byteLength - 2; offset += 2)
        {
            if (BinaryPrimitives.ReadUInt16LittleEndian(entry[offset..]) == 0)
                return false;
        }
        name = Encoding.Unicode.GetString(entry[..(byteLength - 2)]);
        return true;
    }

    private static ReadOnlySpan<byte> CompoundFileSector(
        ReadOnlySpan<byte> bytes,
        int sectorSize,
        uint sector)
    {
        var offset = checked((int)(((long)sector + 1) * sectorSize));
        return bytes.Slice(offset, sectorSize);
    }

    private static bool IsRegularCfbSector(uint sector, int sectorCount) =>
        sector < (uint)sectorCount;

    private static bool AllZero(ReadOnlySpan<byte> bytes)
    {
        foreach (var value in bytes)
        {
            if (value != 0)
                return false;
        }
        return true;
    }

    private sealed class CompoundFileDirectoryEntry
    {
        required public string Name { get; init; }
        required public byte ObjectType { get; init; }
        required public uint LeftSibling { get; init; }
        required public uint RightSibling { get; init; }
        required public uint Child { get; init; }
        required public uint StartingSector { get; init; }
        required public ulong StreamSize { get; init; }
    }

    private static void AddFinding(
        List<VerificationFinding> findings,
        string code,
        VerificationFindingSeverity severity,
        string message,
        ChangeLocation? location = null) => findings.Add(new VerificationFinding
        {
            Code = code,
            Severity = severity,
            Message = message,
            Location = location,
        });

    private sealed class EntryWork
    {
        required public ZipArchiveEntry ArchiveEntry { get; init; }
        required public int ArchiveIndex { get; init; }
        required public string Uri { get; init; }
        required public bool IsDirectory { get; init; }
        required public long Size { get; init; }
        required public long CompressedSize { get; init; }
        public bool? IsEncrypted { get; init; }
        public uint? ExpectedCrc32 { get; init; }
        required public bool RatioExceeded { get; init; }
        public bool XmlLimitReported { get; set; }
        public bool XmlUnparsable { get; set; }
        required public bool EntryLimitExceeded { get; init; }
        public int Occurrence { get; set; }
        public long ActualSize { get; set; }
        public string? ContentType { get; set; }
        public string ContentTypeSource { get; set; } = "unresolved";
        public bool IsXml { get; set; }
        public VerificationDigest? RawBytesDigest { get; set; }
        public VerificationDigest? NormalizedXmlDigest { get; set; }
        public XDocument? Xml { get; set; }
        public byte[]? PreloadedBytes { get; set; }
        public bool ReadBlocked { get; set; }
    }

    private readonly record struct ZipCentralEntryMetadata(bool IsEncrypted, uint Crc32);

    private sealed class ActualReadBudget
    {
        private long _remaining;

        public ActualReadBudget(long maximum) => _remaining = maximum;

        public bool Exceeded { get; private set; }

        public bool TryConsume(int count)
        {
            if (count <= _remaining)
            {
                _remaining -= count;
                return true;
            }
            Exceeded = true;
            _remaining = 0;
            return false;
        }
    }

    private sealed class ContentTypeMap
    {
        public static readonly ContentTypeMap Empty = new(
            new Dictionary<string, string>(PartNameComparer),
            new Dictionary<string, string>(PartNameComparer),
            Array.Empty<PackageContentTypeDeclaration>(),
            isAvailable: false);

        private readonly IReadOnlyDictionary<string, string> _defaults;
        private readonly IReadOnlyDictionary<string, string> _overrides;

        private ContentTypeMap(
            IReadOnlyDictionary<string, string> defaults,
            IReadOnlyDictionary<string, string> overrides,
            IReadOnlyList<PackageContentTypeDeclaration> declarations,
            bool isAvailable)
        {
            _defaults = defaults;
            _overrides = overrides;
            Declarations = declarations;
            IsAvailable = isAvailable;
        }

        public IReadOnlyList<PackageContentTypeDeclaration> Declarations { get; }

        /// <summary>Whether <c>[Content_Types].xml</c> was parsed, however few declarations it held.</summary>
        public bool IsAvailable { get; }

        public (string? ContentType, string Source) Resolve(string uri)
        {
            if (PartNameComparer.Equals(uri, ContentTypesUri))
                return (ContentTypesMime, "implicit");
            if (IsRelationshipPart(uri))
                return (RelationshipsMime, "implicit");
            if (_overrides.TryGetValue(uri, out var overridden))
                return (overridden, "override");
            var lastSlash = uri.LastIndexOf('/');
            var lastDot = uri.LastIndexOf('.');
            if (lastDot > lastSlash && lastDot + 1 < uri.Length
                && _defaults.TryGetValue(uri[(lastDot + 1)..], out var defaulted))
                return (defaulted, "default");
            return (null, "unresolved");
        }

        public static ContentTypeMap Parse(
            XDocument document,
            int maximumUriLength,
            List<VerificationFinding> findings)
        {
            if (document.Root?.Name.LocalName != "Types"
                || document.Root.Name.NamespaceName != TransitionalContentTypesNamespace)
            {
                AddFinding(findings, "malformed_content_types", VerificationFindingSeverity.Error,
                    "[Content_Types].xml root element must be Types.",
                    new ChangeLocation { EntryUri = ContentTypesUri });
                return Empty;
            }

            var defaults = new Dictionary<string, string>(PartNameComparer);
            var overrides = new Dictionary<string, string>(PartNameComparer);
            var declarations = new List<PackageContentTypeDeclaration>();
            var occurrences = new Dictionary<string, int>(PartNameComparer);
            foreach (var element in document.Root.Elements())
            {
                if (element.Name.NamespaceName != document.Root.Name.NamespaceName)
                    continue;
                var kind = element.Name.LocalName switch
                {
                    "Default" => "default",
                    "Override" => "override",
                    _ => null,
                };
                if (kind is null)
                    continue;
                var rawKey = kind == "default"
                    ? UnqualifiedAttribute(element, "Extension")
                    : UnqualifiedAttribute(element, "PartName");
                var contentType = UnqualifiedAttribute(element, "ContentType");
                if (string.IsNullOrWhiteSpace(rawKey) || string.IsNullOrWhiteSpace(contentType))
                {
                    AddFinding(findings, "malformed_content_type", VerificationFindingSeverity.Error,
                        "Content-type declaration is missing its key or ContentType.",
                        new ChangeLocation { EntryUri = ContentTypesUri });
                    continue;
                }
                var validKey = true;
                string key;
                if (kind == "default")
                {
                    // Reported as declared: extension matching is case-insensitive through the
                    // dictionaries below, so lower-casing here would only lose the package's own
                    // spelling from `contentTypes`.
                    key = rawKey;
                    if (!IsValidContentTypeExtension(rawKey, maximumUriLength))
                    {
                        validKey = false;
                        AddFinding(findings, "invalid_content_type_extension",
                            VerificationFindingSeverity.Error,
                            "Content-type Default Extension must be a single non-empty file-extension token.",
                            new ChangeLocation { EntryUri = ContentTypesUri,
                                PropertyPath = "default:" + rawKey });
                    }
                }
                else if (!TryCanonicalizePartName(rawKey, out key)
                    || key.Length > maximumUriLength)
                {
                    validKey = false;
                    key = rawKey;
                    AddFinding(findings, "invalid_content_type_part_name",
                        VerificationFindingSeverity.Error,
                        "Content-type Override PartName does not satisfy the OPC part-name grammar.",
                        new ChangeLocation { EntryUri = ContentTypesUri, TargetUri = rawKey });
                }
                var validContentType = IsValidContentTypeMediaType(contentType);
                if (!validContentType)
                {
                    AddFinding(findings, "malformed_content_type",
                        VerificationFindingSeverity.Error,
                        "ContentType must be a syntactically valid MIME media type without edge whitespace.",
                        new ChangeLocation { EntryUri = ContentTypesUri,
                            PropertyPath = kind + ":" + key });
                }
                var occurrenceKey = kind + "\0" + key;
                occurrences.TryGetValue(occurrenceKey, out var occurrence);
                occurrences[occurrenceKey] = occurrence + 1;
                declarations.Add(new PackageContentTypeDeclaration
                {
                    Kind = kind,
                    Key = key,
                    ContentType = contentType,
                    Occurrence = occurrence,
                });

                if (!validKey || !validContentType)
                    continue;
                var target = kind == "default" ? defaults : overrides;
                if (target.TryGetValue(key, out var existing))
                {
                    AddFinding(findings,
                        string.Equals(existing, contentType, StringComparison.Ordinal)
                            ? "duplicate_content_type" : "conflicting_content_type",
                        VerificationFindingSeverity.Error,
                        string.Equals(existing, contentType, StringComparison.Ordinal)
                            ? "Content-type key is declared more than once."
                            : "Content-type key is declared with conflicting MIME types.",
                        new ChangeLocation { EntryUri = ContentTypesUri,
                            PropertyPath = kind + ":" + key });
                }
                else
                {
                    target.Add(key, contentType);
                }
            }

            return new ContentTypeMap(defaults, overrides, declarations
                .OrderBy(declaration => declaration.Kind, StringComparer.Ordinal)
                .ThenBy(declaration => declaration.Key, PartNameComparer)
                .ThenBy(declaration => declaration.ContentType, StringComparer.Ordinal)
                .ThenBy(declaration => declaration.Occurrence)
                .ToList(),
                isAvailable: true);
        }
    }

    private sealed class OwnerIdComparer : IEqualityComparer<(string OwnerUri, string Id)>
    {
        public static readonly OwnerIdComparer Instance = new();

        public bool Equals((string OwnerUri, string Id) x, (string OwnerUri, string Id) y) =>
            PartNameComparer.Equals(x.OwnerUri, y.OwnerUri)
            && StringComparer.Ordinal.Equals(x.Id, y.Id);

        public int GetHashCode((string OwnerUri, string Id) obj) =>
            HashCode.Combine(PartNameComparer.GetHashCode(obj.OwnerUri),
                StringComparer.Ordinal.GetHashCode(obj.Id));
    }

    private sealed class AsciiCaseInsensitiveComparer : IEqualityComparer<string>, IComparer<string>
    {
        public static readonly AsciiCaseInsensitiveComparer Instance = new();

        public bool Equals(string? x, string? y) => Compare(x, y) == 0;

        public int Compare(string? x, string? y)
        {
            if (ReferenceEquals(x, y))
                return 0;
            if (x is null)
                return -1;
            if (y is null)
                return 1;
            var shared = Math.Min(x.Length, y.Length);
            for (var index = 0; index < shared; index++)
            {
                var left = FoldAscii(x[index]);
                var right = FoldAscii(y[index]);
                if (left != right)
                    return left.CompareTo(right);
            }
            return x.Length.CompareTo(y.Length);
        }

        public int GetHashCode(string value)
        {
            ArgumentNullException.ThrowIfNull(value);
            var hash = new HashCode();
            foreach (var character in value)
                hash.Add(FoldAscii(character));
            return hash.ToHashCode();
        }

        private static char FoldAscii(char value) =>
            value is >= 'a' and <= 'z' ? (char)(value - ('a' - 'A')) : value;
    }

    private enum SafetyLimitKind
    {
        EntryExpansion,
        EntrySize,
        TotalExpansion,
        XmlSize,
    }

    private sealed class ManifestSafetyException(SafetyLimitKind kind) : Exception
    {
        public SafetyLimitKind Kind { get; } = kind;
    }
}
