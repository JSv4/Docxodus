// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

public sealed class DeliverableVerifierHardeningTests
{
    [Fact]
    public void Session_prepares_and_verifies_the_exact_normal_clean_save_bytes()
    {
        using var session = new DocxSession(IrTestDocuments.Create("Alpha").DocumentByteArray);
        var target = Assert.Single(session.FindAllByText("Alpha"));
        Assert.True(Assert.Single(session.ReplaceTextRange(target.Anchor.Id, "Alpha", "Beta")).Success);
        var expectedBytes = session.Save(persistAnchorIds: false);
        var expectedDigest = Digest(expectedBytes);

        var prepared = session.PrepareDeliverable(companionArtifacts: new[]
        {
            new DeliverableCompanionArtifactInput
            {
                ArtifactId = "verification.json",
                Role = DeliverableArtifactRole.RenderReport,
                MediaType = "application/json",
                Availability = DeliverableArtifactAvailability.Available,
                Bytes = Encoding.UTF8.GetBytes("{}"),
                SourcePackageDigest = expectedDigest,
            },
        });

        Assert.Equal(expectedBytes, prepared.DeliverableBytes);
        Assert.Equal(expectedBytes, session.Save(persistAnchorIds: false));
        Assert.Equal(expectedDigest.Value,
            prepared.Report.DeliverablePackage.RawPackageBytesDigest.Value);
        Assert.Equal(expectedDigest.Value,
            Assert.Single(prepared.Report.CompanionArtifacts).SourcePackageDigest!.Value);
        Assert.DoesNotContain(prepared.Report.Findings,
            finding => finding.Code == "artifact.source_digest_mismatch");
    }

    [Fact]
    public void Ordinary_relationship_defects_do_not_suppress_independent_bounded_checks()
    {
        var source = IrTestDocuments.FromBodyXml(
            "<w:p><w:bookmarkStart w:id=\"9\" w:name=\"open\"/>"
            + "<w:r><w:t>{{CLIENT}}</w:t></w:r></w:p>").DocumentByteArray;
        var malformed = RewriteEntry(source, "word/_rels/document.xml.rels", xml =>
        {
            var relationships = XDocument.Parse(xml.TrimStart('\uFEFF'));
            var first = relationships.Root!.Elements().First();
            var duplicate = new XElement(first);
            duplicate.SetAttributeValue("Target", "missing-target.xml");
            relationships.Root.Add(duplicate);
            return relationships.ToString(SaveOptions.DisableFormatting);
        });

        var result = DeliverableVerifier.VerifyDeliverable(malformed, malformed);

        Assert.Contains(result.Findings, finding => finding.Code == "package.conflicting_relationship");
        Assert.Contains(result.Findings, finding => finding.Code == "structure.bookmark_pair_invalid");
        Assert.Contains(result.Findings, finding => finding.Code == "workflow.placeholder_remaining");
        Assert.Contains(result.Checks, check => check.Check == "deliverable.wordprocessing_closure"
            && check.Status == DeliverableCheckStatus.Completed);
        Assert.Contains(result.Checks, check => check.Check == "deliverable.workflow_and_revision_registry"
            && check.Status == DeliverableCheckStatus.Completed);
        Assert.Contains(result.Checks, check => check.Check == "package_delta"
            && check.Status == DeliverableCheckStatus.Completed);
    }

    [Fact]
    public void Workflow_identity_uses_structural_subject_not_only_text_offset()
    {
        var baseline = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>{{PARTY_A}}</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>Safe</w:t></w:r></w:p>").DocumentByteArray;
        var movedAndReplaced = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Safe</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>{{PARTY_B}}</w:t></w:r></w:p>").DocumentByteArray;

        var result = DeliverableVerifier.VerifyDeliverable(baseline, movedAndReplaced);
        var current = Assert.Single(result.Findings,
            finding => finding.Code == "workflow.placeholder_remaining");
        var resolved = Assert.Single(result.ResolvedFindings,
            finding => finding.Code == "workflow.placeholder_remaining");

        Assert.Equal(DeliverableFindingDisposition.New, current.Disposition);
        Assert.Equal(DeliverableFindingDisposition.Resolved, resolved.Disposition);
        Assert.NotEqual(current.FindingId, resolved.FindingId);
        Assert.NotEqual(current.Location!.PropertyPath, resolved.Location!.PropertyPath);
    }

    [Fact]
    public void Workflow_defaults_ignore_legal_brackets_and_spanish_todo_but_keep_explicit_tokens()
    {
        var legal = IrTestDocuments.Create("See [Section 4.2] and [1]; todo está completo.");
        var legalResult = DeliverableVerifier.VerifyDeliverable(legal.DocumentByteArray);
        var optedIn = DeliverableVerifier.VerifyDeliverable(legal.DocumentByteArray,
            new DeliverableVerificationOptions { DetectBracketedAlternativeClauses = true });
        var configured = DeliverableVerifier.VerifyDeliverable(
            IrTestDocuments.Create("todo TODO {{CLIENT}}").DocumentByteArray,
            new DeliverableVerificationOptions { EditorialMarkers = new[] { "TODO" } });

        Assert.DoesNotContain(legalResult.Findings,
            finding => finding.Category == DeliverableFindingCategory.Workflow);
        Assert.Equal(DeliverableVerificationDecision.Passed, legalResult.Decision);
        Assert.Contains(optedIn.Findings, finding => finding.Code == "workflow.alternative_clause"
            && !finding.BlocksDelivery);
        Assert.Contains(configured.Findings, finding => finding.Code == "workflow.editorial_marker"
            && finding.BlocksDelivery);
        Assert.Contains(configured.Findings, finding => finding.Code == "workflow.placeholder_remaining"
            && finding.BlocksDelivery);
    }

    [Fact]
    public void High_confidence_tokens_are_found_in_relationship_reachable_stories()
    {
        var header = IrTestDocuments.FromBodyAndHeaderXml(
            "<w:p><w:r><w:t>&lt;&lt;BODY_TOKEN&gt;&gt;</w:t></w:r></w:p>",
            "<w:p><w:r><w:t>{{HEADER_TOKEN}}</w:t></w:r></w:p>");
        var note = IrTestDocuments.FromBodyXmlWithFootnote(
            "<w:p><w:r><w:t>Safe body</w:t></w:r></w:p>", "${NOTE_TOKEN}");

        var headerResult = DeliverableVerifier.VerifyDeliverable(header.DocumentByteArray);
        var noteResult = DeliverableVerifier.VerifyDeliverable(note.DocumentByteArray);

        Assert.Contains(headerResult.Findings, finding => finding.Code == "workflow.placeholder_remaining"
            && finding.OwningPartUri.Contains("header", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(headerResult.Findings, finding => finding.Code == "workflow.placeholder_remaining"
            && finding.OwningPartUri.EndsWith("document.xml", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(noteResult.Findings, finding => finding.Code == "workflow.placeholder_remaining"
            && finding.OwningPartUri.Contains("footnotes", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Valid_pdf_html_and_canonical_pagemap_form_one_closed_artifact_set()
    {
        var package = IrTestDocuments.Create("Rendered").DocumentByteArray;
        var source = Digest(package);
        const string renderer = "fixture/2; dpi=144";
        var map = PageMapBytes(renderer, pages: 1);
        var mapDigest = Digest(map);
        var result = DeliverableVerifier.VerifyDeliverable(new DeliverableVerificationRequest
        {
            DeliverableBytes = package,
            CompanionArtifacts = new[]
            {
                Artifact("map.json", DeliverableArtifactRole.PageMap,
                    "application/vnd.docxodus.pagemap+json", map, source, renderer, 1),
                Artifact("preview.pdf", DeliverableArtifactRole.Pdf, "application/pdf",
                    MinimalPdf(), source, renderer, 1, mapDigest),
                Artifact("preview.html", DeliverableArtifactRole.Html, "text/html; charset=utf-8",
                    Encoding.UTF8.GetBytes("<!doctype html><html><body>Rendered</body></html>"),
                    source, renderer, 1, mapDigest),
            },
        });

        Assert.Equal(DeliverableVerificationDecision.Passed, result.Decision);
        Assert.DoesNotContain(result.Findings,
            finding => finding.Category == DeliverableFindingCategory.Artifact);
        Assert.Equal(mapDigest.Value, Assert.Single(result.CompanionArtifacts,
            artifact => artifact.Role == DeliverableArtifactRole.PageMap).PageMapDigest!.Value);
    }

    [Fact]
    public void Stale_forged_malformed_and_incompletely_bound_artifacts_fail_closed()
    {
        var package = IrTestDocuments.Create("Rendered").DocumentByteArray;
        var source = Digest(package);
        var stale = Digest(Encoding.UTF8.GetBytes("different package"));
        const string renderer = "fixture/2";
        var map = PageMapBytes(renderer, pages: 1);
        var mapDigest = Digest(map);
        var wrongMapDigest = Digest(Encoding.UTF8.GetBytes("wrong map"));
        var artifacts = new[]
        {
            Artifact("map.json", DeliverableArtifactRole.PageMap, "application/json",
                map, source, renderer, 1) with { PageMapDigest = wrongMapDigest },
            Artifact("bad-map.json", DeliverableArtifactRole.PageMap, "application/json",
                Encoding.UTF8.GetBytes("{\"schemaVersion\":1}"), source, renderer, 1),
            Artifact("fake.pdf", DeliverableArtifactRole.Pdf, "application/pdf",
                Encoding.ASCII.GetBytes("%PDF-1.7 test fixture"), source, renderer, 1, mapDigest),
            Artifact("stale.html", DeliverableArtifactRole.Html, "text/html",
                Encoding.UTF8.GetBytes("<html><body>x</body></html>"), stale, renderer, 1, mapDigest),
            Artifact("wrong-map.html", DeliverableArtifactRole.Html, "text/html",
                Encoding.UTF8.GetBytes("<html><body>x</body></html>"), source, renderer, 1, wrongMapDigest),
            Artifact("mismatch.pdf", DeliverableArtifactRole.Pdf, "application/pdf",
                MinimalPdf(), source, "other-renderer", 2, mapDigest),
            new DeliverableCompanionArtifactInput
            {
                ArtifactId = "unbound.html",
                Role = DeliverableArtifactRole.Html,
                MediaType = "text/html",
                Availability = DeliverableArtifactAvailability.Available,
                Bytes = Encoding.UTF8.GetBytes("<html><body>x</body></html>"),
                RendererFingerprint = renderer,
                PageCount = 1,
                PageMapDigest = mapDigest,
            },
        };

        var result = DeliverableVerifier.VerifyDeliverable(new DeliverableVerificationRequest
        {
            DeliverableBytes = package,
            CompanionArtifacts = artifacts,
        });

        Assert.Equal(DeliverableVerificationDecision.Failed, result.Decision);
        AssertCodes(result, "artifact.page_map_malformed", "artifact.page_map_digest_mismatch",
            "artifact.pdf_malformed",
            "artifact.source_digest_mismatch", "artifact.page_map_missing",
            "artifact.page_map_renderer_mismatch", "artifact.page_map_count_mismatch",
            "artifact.source_digest_missing");
    }

    [Fact]
    public void Orphan_definition_parts_cannot_satisfy_story_references_or_sdt_bindings()
    {
        var body = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:commentReference w:id=\"7\"/></w:r></w:p>"
            + "<w:sdt><w:sdtPr><w:dataBinding w:storeItemID=\"{11111111-1111-1111-1111-111111111111}\" "
            + "w:xpath=\"/root\"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>")
            .DocumentByteArray;
        var withOrphans = AddEntries(body,
            ("word/comments.xml", $"<w:comments xmlns:w=\"{IrTestDocuments.W}\">"
                + "<w:comment w:id=\"7\"><w:p/></w:comment></w:comments>"),
            ("customXml/itemProps1.xml",
                "<ds:datastoreItem ds:itemID=\"{11111111-1111-1111-1111-111111111111}\" "
                + "xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\"/>"));
        withOrphans = RewriteContentTypes(withOrphans,
            ("/word/comments.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"),
            ("/customXml/itemProps1.xml",
                "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"));

        var result = DeliverableVerifier.VerifyDeliverable(withOrphans);

        Assert.Contains(result.Findings, finding => finding.Code == "structure.comment_definition_missing");
        Assert.Contains(result.Findings,
            finding => finding.Code == "structure.content_control_store_item_missing");
    }

    [Fact]
    public void Numbering_overrides_are_supported_while_duplicate_overrides_are_ambiguous()
    {
        const string body = "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"8\"/>"
            + "<w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>";
        const string abstractXml = "<w:abstractNum w:abstractNumId=\"1\">"
            + "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/>"
            + "<w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>";
        const string overrideXml = "<w:lvlOverride w:ilvl=\"8\"><w:lvl w:ilvl=\"8\">"
            + "<w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"•\"/>"
            + "</w:lvl></w:lvlOverride>";
        var valid = IrTestDocuments.FromParts(body,
            numberingInnerXml: abstractXml + "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/>"
                + overrideXml + "</w:num>");
        var duplicate = IrTestDocuments.FromParts(body,
            numberingInnerXml: abstractXml + "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/>"
                + overrideXml + overrideXml + "</w:num>");

        var validResult = DeliverableVerifier.VerifyDeliverable(valid.DocumentByteArray);
        var duplicateResult = DeliverableVerifier.VerifyDeliverable(duplicate.DocumentByteArray);

        Assert.DoesNotContain(validResult.Findings,
            finding => finding.Code == "structure.numbering_level_missing");
        Assert.Contains(duplicateResult.Findings,
            finding => finding.Code == "structure.numbering_override_level_duplicate");
    }

    [Fact]
    public void Native_revision_registry_reports_malformed_ambiguous_and_unsupported_groups()
    {
        var package = IrTestDocuments.FromBodyXml(
            "<w:p><w:ins w:author=\"Missing\"><w:r><w:t>Malformed</w:t></w:r></w:ins></w:p>"
            + "<w:p><w:ins w:id=\"7\" w:author=\"First\"><w:r><w:t>One</w:t></w:r></w:ins></w:p>"
            + "<w:p><w:del w:id=\"7\" w:author=\"Second\"><w:r><w:delText>Two</w:delText>"
            + "</w:r></w:del></w:p>"
            + "<w:p><w:customXmlMoveFromRangeStart w:id=\"9\" w:author=\"Third\"/>"
            + "<w:r><w:t>Unsupported</w:t></w:r>"
            + "<w:customXmlMoveFromRangeEnd w:id=\"9\"/></w:p>").DocumentByteArray;

        var result = DeliverableVerifier.VerifyDeliverable(package);

        Assert.Contains(result.Findings,
            finding => finding.Code == "structure.revision_malformed");
        Assert.Contains(result.Findings,
            finding => finding.Code == "structure.revision_ambiguous");
        Assert.Contains(result.Findings,
            finding => finding.Code == "structure.revision_unsupported");
    }

    [Fact]
    public void Shared_detector_budget_stops_large_marker_and_placeholder_scans_deterministically()
    {
        var body = string.Concat(Enumerable.Range(0, 200).Select(index =>
            $"<w:p><w:bookmarkStart w:id=\"{index}\" w:name=\"b{index}\"/>"
            + $"<w:r><w:t>{{{{TOKEN_{index}}}}}</w:t></w:r><w:bookmarkEnd w:id=\"{index}\"/></w:p>"));
        var package = IrTestDocuments.FromBodyXml(body).DocumentByteArray;
        var options = new DeliverableVerificationOptions
        {
            MaxDetectorNodes = 150,
            MaxDetectorSteps = 500,
            MaxDetectorRegexMatches = 5,
        };

        var first = DeliverableVerifier.VerifyDeliverable(package, options);
        var second = DeliverableVerifier.VerifyDeliverable(package, options);
        var regexBounded = DeliverableVerifier.VerifyDeliverable(package,
            new DeliverableVerificationOptions
            {
                MaxDetectorNodes = 100_000,
                MaxDetectorSteps = 100_000,
                MaxDetectorRegexMatches = 5,
            });

        Assert.Contains(first.Findings,
            finding => finding.Code == "verification.resource_budget_exceeded");
        Assert.Contains(first.Checks, check => check.Status == DeliverableCheckStatus.UnavailableEvidence
            && (check.Diagnostic?.Contains("resource budget exceeded", StringComparison.Ordinal) ?? false));
        Assert.Equal(first.ToCanonicalJson(), second.ToCanonicalJson());
        Assert.Contains(regexBounded.Findings, finding =>
            finding.Code == "verification.resource_budget_exceeded"
            && finding.Location?.PropertyPath == "detectorBudget/regex_matches");
    }

    [Fact]
    public void OpenXml_findings_are_locale_independent()
    {
        var package = RewriteEntry(IrTestDocuments.Create("Invalid").DocumentByteArray,
            "word/document.xml", xml => xml.Replace("<w:p>",
                "<w:p w:invalidFixture=\"true\">", StringComparison.Ordinal));
        var originalCulture = CultureInfo.CurrentCulture;
        var originalUi = CultureInfo.CurrentUICulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo("fr-FR");
            var french = DeliverableVerifier.VerifyDeliverable(package);
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("tr-TR");
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo("tr-TR");
            var turkish = DeliverableVerifier.VerifyDeliverable(package);

            Assert.Equal(french.ToCanonicalJson(), turkish.ToCanonicalJson());
            Assert.All(french.Findings.Where(finding =>
                    finding.Category == DeliverableFindingCategory.OpenXml),
                finding => Assert.StartsWith("Open XML validation", finding.Message,
                    StringComparison.Ordinal));
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
            CultureInfo.CurrentUICulture = originalUi;
        }
    }

    private static DeliverableCompanionArtifactInput Artifact(
        string id,
        DeliverableArtifactRole role,
        string mediaType,
        byte[] bytes,
        VerificationDigest source,
        string renderer,
        long pageCount,
        VerificationDigest? pageMap = null) => new()
    {
        ArtifactId = id,
        Role = role,
        MediaType = mediaType,
        Availability = DeliverableArtifactAvailability.Available,
        Bytes = bytes,
        SourcePackageDigest = source,
        RendererFingerprint = renderer,
        PageCount = pageCount,
        PageMapDigest = pageMap,
    };

    private static byte[] PageMapBytes(string renderer, int pages)
    {
        var pageList = Enumerable.Range(1, pages).Select(page => new PageMapPage
        {
            PageNumber = page,
            PageInSection = page,
            Width = 612,
            Height = 792,
            SectionIndex = 0,
            PageName = "letter",
        }).ToArray();
        var map = new PageMap
        {
            Mode = PageMapMode.Paginated,
            Availability = PageMapAvailability.Available,
            DocumentVersion = 0,
            RendererFingerprint = renderer,
            Pages = pageList,
            Fragments = new[]
            {
                new PageMapFragment
                {
                    FragmentId = "anchor:p1:0",
                    AnchorId = "anchor",
                    FragmentIndex = 0,
                    PageNumber = 1,
                    Geometry = new PageMapRect(72, 72, 100, 20),
                    Story = PageMapStory.Body,
                },
            },
        };
        return Encoding.UTF8.GetBytes(DocxSessionJson.SerializePageMap(map));
    }

    private static byte[] MinimalPdf() => Encoding.ASCII.GetBytes(
        "%PDF-1.4\n1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n"
        + "2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj\n"
        + "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >> endobj\n"
        + "xref\n0 4\n0000000000 65535 f \n"
        + "trailer << /Size 4 /Root 1 0 R >>\nstartxref\n0\n%%EOF\n");

    private static VerificationDigest Digest(byte[] bytes) => new()
    {
        Algorithm = "SHA-256",
        Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
    };

    private static byte[] RewriteEntry(byte[] package, string entryName, Func<string, string> rewrite)
    {
        using var output = new MemoryStream();
        using (var destination = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        using (var source = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read))
        {
            foreach (var sourceEntry in source.Entries)
            {
                var destinationEntry = destination.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
                destinationEntry.LastWriteTime = FixtureTimestamp;
                using var input = sourceEntry.Open();
                using var copied = new MemoryStream();
                input.CopyTo(copied);
                var bytes = sourceEntry.FullName == entryName
                    ? Encoding.UTF8.GetBytes(rewrite(Encoding.UTF8.GetString(copied.ToArray())))
                    : copied.ToArray();
                using var entryOutput = destinationEntry.Open();
                entryOutput.Write(bytes);
            }
        }
        return output.ToArray();
    }

    private static byte[] AddEntries(byte[] package, params (string Name, string Xml)[] additions)
    {
        using var output = new MemoryStream();
        using (var destination = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        using (var source = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read))
        {
            foreach (var sourceEntry in source.Entries)
            {
                var destinationEntry = destination.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
                destinationEntry.LastWriteTime = FixtureTimestamp;
                using var input = sourceEntry.Open();
                using var entryOutput = destinationEntry.Open();
                input.CopyTo(entryOutput);
            }
            foreach (var (name, xml) in additions)
            {
                var entry = destination.CreateEntry(name, CompressionLevel.Optimal);
                entry.LastWriteTime = FixtureTimestamp;
                using var entryOutput = entry.Open();
                entryOutput.Write(Encoding.UTF8.GetBytes(xml));
            }
        }
        return output.ToArray();
    }

    private static byte[] RewriteContentTypes(
        byte[] package,
        params (string PartName, string ContentType)[] additions) => RewriteEntry(
        package, "[Content_Types].xml", xml =>
        {
            var contentTypes = XDocument.Parse(xml.TrimStart('\uFEFF'));
            XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
            foreach (var (partName, contentType) in additions)
                contentTypes.Root!.Add(new XElement(ns + "Override",
                    new XAttribute("PartName", partName), new XAttribute("ContentType", contentType)));
            return contentTypes.ToString(SaveOptions.DisableFormatting);
        });

    private static void AssertCodes(DeliverableVerificationResult result, params string[] codes)
    {
        foreach (var code in codes)
            Assert.Contains(result.Findings, finding => finding.Code == code);
    }

    private static readonly DateTimeOffset FixtureTimestamp =
        new(2026, 1, 1, 0, 0, 0, TimeSpan.Zero);
}
