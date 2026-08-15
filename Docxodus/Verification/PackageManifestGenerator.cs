// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

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
    private const string StrictPackageRelationshipsNamespace =
        "http://purl.oclc.org/ooxml/package/relationships";
    private const string TransitionalContentTypesNamespace =
        "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string StrictContentTypesNamespace =
        "http://purl.oclc.org/ooxml/package/content-types";
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
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);
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
    public static PackageManifest Generate(byte[] packageBytes, PackageManifestOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(packageBytes);
        options ??= new PackageManifestOptions();
        options.Validate();

        var rawDigest = Digest(packageBytes);
        var findings = new List<VerificationFinding>();
        if (packageBytes.AsSpan().StartsWith(OleSignature))
            return OleManifest(packageBytes, rawDigest, findings);

        try
        {
            using var stream = new MemoryStream(packageBytes, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            return GenerateZipManifest(archive, packageBytes, rawDigest, options, findings);
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or ArgumentException
            or OverflowException)
        {
            AddFinding(findings, "malformed_package", VerificationFindingSeverity.Error,
                $"The supplied bytes are not a readable ZIP/OPC package ({ex.GetType().Name}).");
            return FinalizeManifest(
                "malformed", rawDigest, null, null,
                Array.Empty<PackageManifestEntry>(),
                Array.Empty<PackageContentTypeDeclaration>(),
                Array.Empty<PackageRelationship>(),
                new PackageManifestFacts(), findings);
        }
    }

    /// <summary>Generate the canonical schema-v1 JSON representation.</summary>
    public static string GenerateJson(
        byte[] packageBytes,
        PackageManifestOptions? options = null,
        bool indented = false) => Generate(packageBytes, options).ToJson(indented);

    private static PackageManifest GenerateZipManifest(
        ZipArchive archive,
        byte[] packageBytes,
        VerificationDigest rawDigest,
        PackageManifestOptions options,
        List<VerificationFinding> findings)
    {
        var archiveEntries = archive.Entries.ToList();
        var encryptedFlags = TryReadEncryptedFlags(packageBytes, archiveEntries.Count);
        if (encryptedFlags is null)
        {
            AddFinding(findings, "zip_encryption_detection_unavailable",
                VerificationFindingSeverity.Error,
                "ZIP central-directory encryption flags could not be parsed authoritatively.",
                new ChangeLocation { PropertyPath = "entries[].isEncrypted" });
        }

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

            bool? encrypted = encryptedFlags is not null && index < encryptedFlags.Count
                ? encryptedFlags[index]
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

            if (isDirectory)
            {
                AddFinding(findings, "directory_entry", VerificationFindingSeverity.Warning,
                    "OPC packages should not contain directory-only ZIP entries.",
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
                RatioExceeded = ratioExceeded,
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
        var readBudget = new ActualReadBudget(options.MaxTotalUncompressedBytes);
        var contentTypeMap = ReadContentTypes(
            works, totalLimitExceeded, options, readBudget, findings);
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
            && works.Any(work => string.Equals(work.Uri, ContentTypesUri, StringComparison.OrdinalIgnoreCase)))
        {
            AddFinding(findings, "content_types_unreadable", VerificationFindingSeverity.Error,
                "[Content_Types].xml is present but could not be used, so no part content type was resolved.",
                new ChangeLocation { EntryUri = ContentTypesUri });
        }
        ValidateContentTypeTargets(contentTypeMap, works, findings);

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
            works, options, payloadsInspected, findings, out var unreadableOwners);
        ValidateRelationshipReferences(works, relationships, unreadableOwners, findings);
        var facts = BuildFacts(works, relationships);

        // A truncated inspection has not seen the whole package, so it cannot state a content
        // identity: two packages differing only past the cut would otherwise compare equal.
        VerificationDigest? orderedContentDigest = null;
        if (!totalLimitExceeded && !entryCountExceeded
            && works.All(work => work.RawBytesDigest is not null
                && work.IsEncrypted == false && !work.RatioExceeded))
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

        var entryModels = works
            .OrderBy(work => work.Uri, StringComparer.Ordinal)
            .ThenBy(work => work.Occurrence)
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

        var hasContentTypes = works.Any(work =>
            string.Equals(work.Uri, ContentTypesUri, StringComparison.OrdinalIgnoreCase));
        var isEncrypted = works.Any(work => work.IsEncrypted == true);
        var packageKind = isEncrypted
            ? "zip-encrypted"
            : hasContentTypes ? "opc" : "zip";
        return FinalizeManifest(packageKind, rawDigest, orderedContentDigest, semanticDigest,
            entryModels, contentTypeMap.Declarations, relationships, facts, findings);
    }

    private static PackageManifest OleManifest(
        byte[] bytes,
        VerificationDigest rawDigest,
        List<VerificationFinding> findings)
    {
        var encrypted = ContainsUtf16Name(bytes, "EncryptedPackage")
            || ContainsUtf16Name(bytes, "EncryptionInfo");
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
        bool totalLimitExceeded,
        PackageManifestOptions options,
        ActualReadBudget readBudget,
        List<VerificationFinding> findings)
    {
        var candidates = works.Where(work =>
                string.Equals(work.Uri, ContentTypesUri, StringComparison.OrdinalIgnoreCase))
            .OrderBy(work => work.ArchiveIndex)
            .ToList();
        if (candidates.Count == 0)
        {
            AddFinding(findings, "missing_content_types", VerificationFindingSeverity.Error,
                "The package has no [Content_Types].xml entry.",
                new ChangeLocation { EntryUri = ContentTypesUri });
            return ContentTypeMap.Empty;
        }
        if (totalLimitExceeded)
            return ContentTypeMap.Empty;

        var selected = candidates[0];
        if (selected.IsEncrypted != false || selected.RatioExceeded)
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
                ExpansionCeiling(selected.CompressedSize, options.MaxCompressionRatio),
                readBudget);
            selected.PreloadedBytes = bytes;
            var document = XmlSemanticNormalizer.Parse(bytes,
                Math.Max(options.MaxXmlPartBytes * 2, 1));
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
        if (work.IsEncrypted != false || work.RatioExceeded)
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
            var actualXmlLimitExceeded = false;
            int read;
            while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                if (read > expansionRemaining)
                    throw new ManifestSafetyException(SafetyLimitKind.EntryExpansion);
                expansionRemaining -= read;
                if (!readBudget.TryConsume(read))
                    throw new ManifestSafetyException(SafetyLimitKind.TotalExpansion);
                if (readTotal > long.MaxValue - read)
                    throw new ManifestSafetyException(SafetyLimitKind.EntryExpansion);
                readTotal += read;
                hash.AppendData(buffer, 0, read);
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
                Math.Max(options.MaxXmlPartBytes * 2, 1));
            work.Xml = document;
            work.NormalizedXmlDigest = XmlSemanticNormalizer.Digest(
                document, work.Uri, IsKnownOoxmlXml(work));
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
        PackageManifestOptions options,
        bool payloadsInspected,
        List<VerificationFinding> findings,
        out HashSet<string> unreadableOwners)
    {
        var entryUris = works.Select(work => work.Uri).ToHashSet(StringComparer.OrdinalIgnoreCase);
        unreadableOwners = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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

                var external = string.Equals(rawMode, "External", StringComparison.OrdinalIgnoreCase);
                if (!string.IsNullOrEmpty(rawMode)
                    && !external
                    && !string.Equals(rawMode, "Internal", StringComparison.OrdinalIgnoreCase))
                {
                    AddFinding(findings, "invalid_target_mode", VerificationFindingSeverity.Error,
                        "Relationship TargetMode must be Internal or External.",
                        new ChangeLocation { EntryUri = work.Uri, OwnerUri = owner,
                            RelationshipId = id, TargetUri = target });
                }

                string? resolved = null;
                bool? targetPresent = null;
                if (!external)
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
            .GroupBy(relationship => relationship.OwnerUri, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key,
                group => group.Select(relationship => relationship.Id)
                    .ToHashSet(StringComparer.Ordinal),
                StringComparer.OrdinalIgnoreCase);
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        foreach (var work in works.Where(work => work.Xml?.Root is not null
                     && !IsRelationshipPart(work.Uri)
                     && !string.Equals(work.Uri, ContentTypesUri, StringComparison.OrdinalIgnoreCase)))
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
                                 "id" or "embed" or "link" or "dm" or "lo" or "qs" or "cs"
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
                string.Equals(work.Uri, uri, StringComparison.OrdinalIgnoreCase)
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
                && string.Equals(work.Uri, mainDocumentUri, StringComparison.OrdinalIgnoreCase))
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
            MediaPartCount = works.Count(work =>
                work.ContentType?.StartsWith("image/", StringComparison.OrdinalIgnoreCase) == true),
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
        foreach (var group in works.GroupBy(work => work.Uri, StringComparer.OrdinalIgnoreCase))
        {
            var occurrence = 0;
            foreach (var work in group
                         .OrderBy(work => work.RawBytesDigest?.Value ?? string.Empty, StringComparer.Ordinal)
                         .ThenBy(work => work.Size)
                         .ThenBy(work => work.ArchiveIndex))
                work.Occurrence = occurrence++;
        }
    }

    // Directory-only entries carry no content, so they stay out of both content identities: a
    // repack that adds or drops folder entries is packaging, not a document change.
    private static IEnumerable<EntryWork> StableEntryOrder(IReadOnlyList<EntryWork> works) =>
        works.Where(work => !work.IsDirectory)
            .OrderBy(work => work.Uri, StringComparer.Ordinal)
            .ThenBy(work => work.Occurrence);

    private static void FindDuplicateEntryNames(
        IReadOnlyList<EntryWork> works,
        List<VerificationFinding> findings)
    {
        foreach (var group in works.GroupBy(work => work.Uri, StringComparer.OrdinalIgnoreCase)
                     .Where(group => group.Count() > 1))
        {
            AddFinding(findings, "duplicate_entry", VerificationFindingSeverity.Error,
                "Package contains multiple ZIP entries for the same canonical URI.",
                new ChangeLocation { EntryUri = group.Key });
        }
    }

    private static void FindConflictingEntries(
        IReadOnlyList<EntryWork> works,
        List<VerificationFinding> findings)
    {
        foreach (var group in works.GroupBy(work => work.Uri, StringComparer.OrdinalIgnoreCase)
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
        IReadOnlyList<EntryWork> works,
        List<VerificationFinding> findings)
    {
        var entries = works.Select(work => work.Uri).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in map.Declarations.Where(declaration => declaration.Kind == "override"))
        {
            if (!entries.Contains(declaration.Key))
            {
                AddFinding(findings, "missing_content_type_target", VerificationFindingSeverity.Error,
                    "Content-type Override names a part that is absent from the package.",
                    new ChangeLocation { EntryUri = ContentTypesUri, TargetUri = declaration.Key });
            }
        }
    }

    private static bool IsXml(string uri, string? contentType) =>
        string.Equals(uri, ContentTypesUri, StringComparison.OrdinalIgnoreCase)
        || IsRelationshipPart(uri)
        || uri.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
        || uri.EndsWith(".vml", StringComparison.OrdinalIgnoreCase)
        || contentType?.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) == true
        || contentType?.Equals("application/xml", StringComparison.OrdinalIgnoreCase) == true
        || contentType?.Equals("text/xml", StringComparison.OrdinalIgnoreCase) == true;

    private static bool IsKnownOoxmlXml(EntryWork work) =>
        string.Equals(work.Uri, ContentTypesUri, StringComparison.OrdinalIgnoreCase)
        || IsRelationshipPart(work.Uri)
        || (work.ContentType is not null
            && KnownOoxmlXmlContentTypes.Contains(work.ContentType));

    private static bool IsRelationshipPart(string uri) =>
        XmlSemanticNormalizer.IsRelationshipPart(uri);

    private static string? RelationshipOwner(string relationshipPartUri)
    {
        if (relationshipPartUri.Equals("/_rels/.rels", StringComparison.OrdinalIgnoreCase))
            return "/";
        var marker = relationshipPartUri.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
        if (marker < 0 || !relationshipPartUri.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            return null;
        var directory = relationshipPartUri[..marker];
        var file = relationshipPartUri[(marker + "/_rels/".Length)..^".rels".Length];
        if (file.Length == 0)
            return null;
        var owner = directory + "/" + file;
        return IsValidDecodedPartName(owner) ? owner : null;
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
        if (!TryDecodeOpcPath(rawPath, requireAbsolute: false, allowRelative: true,
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
        if (!IsValidDecodedPartName(resolved) || resolved.Length > maximumUriLength)
        {
            invalidTarget = true;
            return null;
        }
        return resolved;
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

        // A single trailing forward slash marks a directory-only entry. Those are packaging
        // artifacts, not OPC parts, so the grammar is applied to the path they name and the
        // slash is kept in the canonical URI so a folder can never collide with a real part.
        // A trailing backslash is still a malformed path and is left to fail below.
        var isDirectory = name.EndsWith("/", StringComparison.Ordinal);
        var body = isDirectory ? name[..^1] : name;
        if (body.Length == 0)
            return false;
        if (!TryDecodeOpcPath(body, requireAbsolute: false, allowRelative: true,
                allowDotSegments: false, out var isAbsolute, out var segments)
            || isAbsolute)
            return false;
        var joined = "/" + string.Join('/', segments);
        if (!IsValidDecodedPartName(joined))
            return false;
        canonical = isDirectory ? joined + "/" : joined;
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
        if (!TryDecodeOpcPath(rawName, requireAbsolute: true, allowRelative: false,
                allowDotSegments: false, out _, out var segments))
            return false;
        canonical = "/" + string.Join('/', segments);
        return IsValidDecodedPartName(canonical);
    }

    private static bool TryDecodeOpcPath(
        string rawPath,
        bool requireAbsolute,
        bool allowRelative,
        bool allowDotSegments,
        out bool isAbsolute,
        out List<string> decodedSegments)
    {
        decodedSegments = new List<string>();
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
            if (!TryDecodeOpcSegment(rawSegment, allowDotSegments, out var segment))
                return false;
            decodedSegments.Add(segment);
        }
        return true;
    }

    private static bool TryDecodeOpcSegment(
        string rawSegment,
        bool allowDotSegments,
        out string decoded)
    {
        decoded = string.Empty;
        if (rawSegment.Length == 0)
            return false;
        var bytes = new List<byte>(rawSegment.Length);
        var literalStart = 0;
        try
        {
            for (var index = 0; index < rawSegment.Length; index++)
            {
                if (rawSegment[index] != '%')
                    continue;
                if (index > literalStart)
                    bytes.AddRange(StrictUtf8.GetBytes(rawSegment[literalStart..index]));
                if (index + 2 >= rawSegment.Length
                    || !TryHex(rawSegment[index + 1], out var high)
                    || !TryHex(rawSegment[index + 2], out var low))
                    return false;
                var encoded = (byte)((high << 4) | low);
                if (encoded is (byte)'/' or (byte)'\\' || IsUnreserved(encoded))
                    return false;
                bytes.Add(encoded);
                index += 2;
                literalStart = index + 1;
            }
            if (literalStart < rawSegment.Length)
                bytes.AddRange(StrictUtf8.GetBytes(rawSegment[literalStart..]));
            decoded = StrictUtf8.GetString(bytes.ToArray());
        }
        catch (EncoderFallbackException)
        {
            return false;
        }
        catch (DecoderFallbackException)
        {
            return false;
        }

        if (decoded.Length == 0
            || decoded.Any(character => char.IsControl(character)
                || character is '/' or '\\' or '?' or '#'))
            return false;
        if (decoded is "." or "..")
            return allowDotSegments;
        return !decoded.EndsWith(".", StringComparison.Ordinal);
    }

    private static bool IsValidDecodedPartName(string value)
    {
        if (!value.StartsWith("/", StringComparison.Ordinal)
            || value.StartsWith("//", StringComparison.Ordinal) || value.Length == 1)
            return false;
        return value[1..].Split('/').All(segment =>
            segment.Length > 0 && segment is not "." and not ".."
            && !segment.EndsWith(".", StringComparison.Ordinal)
            && !segment.Any(character => char.IsControl(character)
                || character is '/' or '\\' or '?' or '#'));
    }

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

    private static bool IsUnreserved(byte value) =>
        value is >= (byte)'A' and <= (byte)'Z'
        || value is >= (byte)'a' and <= (byte)'z'
        || value is >= (byte)'0' and <= (byte)'9'
        || value is (byte)'-' or (byte)'.' or (byte)'_' or (byte)'~';

    private static IReadOnlyList<bool>? TryReadEncryptedFlags(byte[] bytes, int expectedCount)
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
        var flags = new List<bool>(count);
        var position = (int)centralOffset;
        var centralEnd = centralOffset + centralSize;
        for (var index = 0; index < count; index++)
        {
            if (position < 0 || position + 46 > bytes.Length
                || (ulong)(position + 46) > centralEnd
                || ReadUInt32(bytes, position) != centralSignature)
                return null;
            flags.Add((ReadUInt16(bytes, position + 8) & 1) != 0);
            var nameLength = ReadUInt16(bytes, position + 28);
            var extraLength = ReadUInt16(bytes, position + 30);
            var commentLength = ReadUInt16(bytes, position + 32);
            var nextPosition = (long)position + 46 + nameLength + extraLength + commentLength;
            if (nextPosition > bytes.Length || nextPosition > int.MaxValue
                || (ulong)nextPosition > centralEnd)
                return null;
            position = (int)nextPosition;
        }
        return flags;
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
        long maximum,
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
            if (!readBudget.TryConsume(read))
                throw new ManifestSafetyException(SafetyLimitKind.TotalExpansion);
            if (read > maximum - total)
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

    private static string SafetyFindingCode(SafetyLimitKind kind) => kind switch
    {
        SafetyLimitKind.EntryExpansion => "compression_ratio_limit_exceeded",
        SafetyLimitKind.XmlSize => "xml_size_limit_exceeded",
        _ => "entry_expansion_limit_exceeded",
    };

    private static string SafetyFindingMessage(SafetyLimitKind kind, string subject) => kind switch
    {
        SafetyLimitKind.EntryExpansion =>
            $"{subject} produced more actual bytes than compressed size × MaxCompressionRatio permits.",
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
            && WordprocessingXmlContentTypes.Contains(work.ContentType);

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

    private static bool IsPackageRelationshipsNamespace(string value) => value is
        TransitionalPackageRelationshipsNamespace or StrictPackageRelationshipsNamespace;

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
        string.Equals(work.ContentType, contentType, StringComparison.OrdinalIgnoreCase);

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

    private static bool ContainsUtf16Name(byte[] bytes, string value)
    {
        var needle = Encoding.Unicode.GetBytes(value);
        return bytes.AsSpan().IndexOf(needle) >= 0;
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
        required public bool RatioExceeded { get; init; }
        public bool XmlLimitReported { get; set; }
        public bool XmlUnparsable { get; set; }
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
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
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
            if (string.Equals(uri, ContentTypesUri, StringComparison.OrdinalIgnoreCase))
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
                || document.Root.Name.NamespaceName is not
                    (TransitionalContentTypesNamespace or StrictContentTypesNamespace))
            {
                AddFinding(findings, "malformed_content_types", VerificationFindingSeverity.Error,
                    "[Content_Types].xml root element must be Types.",
                    new ChangeLocation { EntryUri = ContentTypesUri });
                return Empty;
            }

            var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var declarations = new List<PackageContentTypeDeclaration>();
            var occurrences = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
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

                if (!validKey)
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
                .ThenBy(declaration => declaration.Key, StringComparer.OrdinalIgnoreCase)
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
            StringComparer.OrdinalIgnoreCase.Equals(x.OwnerUri, y.OwnerUri)
            && StringComparer.Ordinal.Equals(x.Id, y.Id);

        public int GetHashCode((string OwnerUri, string Id) obj) =>
            HashCode.Combine(StringComparer.OrdinalIgnoreCase.GetHashCode(obj.OwnerUri),
                StringComparer.Ordinal.GetHashCode(obj.Id));
    }

    private enum SafetyLimitKind
    {
        EntryExpansion,
        TotalExpansion,
        XmlSize,
    }

    private sealed class ManifestSafetyException(SafetyLimitKind kind) : Exception
    {
        public SafetyLimitKind Kind { get; } = kind;
    }
}
