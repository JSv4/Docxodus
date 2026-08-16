// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

public sealed class DeliverableVerifierTests
{
    private const string W = IrTestDocuments.W;

    [Fact]
    public void Valid_complex_package_passes_deterministically_without_mutation()
    {
        var document = ComplexValidDocument();
        var bytes = document.DocumentByteArray.ToArray();
        var before = bytes.ToArray();

        var first = DeliverableVerifier.VerifyDeliverable(bytes);
        var second = DeliverableVerifier.VerifyDeliverable(bytes);

        Assert.True(first.AnalysisCompleted, Diagnostics(first));
        Assert.Equal(DeliverableVerificationDecision.Passed, first.Decision);
        Assert.Empty(first.Findings);
        Assert.Equal(first.ToCanonicalJson(), second.ToCanonicalJson());
        Assert.Equal(before, bytes);
        using var json = JsonDocument.Parse(first.ToCanonicalUtf8Bytes());
        Assert.Equal(DeliverableVerificationResult.SchemaId,
            json.RootElement.GetProperty("schema").GetString());
    }

    [Fact]
    public void Corrupt_input_returns_bounded_structured_failure()
    {
        var bytes = Encoding.UTF8.GetBytes("this is not a DOCX package");

        var result = DeliverableVerifier.VerifyDeliverable(bytes);

        Assert.False(result.AnalysisCompleted);
        Assert.Equal(DeliverableVerificationDecision.Failed, result.Decision);
        Assert.False(result.DeliverablePackage.ManifestValid);
        Assert.Contains(result.Findings, finding => finding.Code.StartsWith("package.", StringComparison.Ordinal));
        Assert.Contains(result.Checks, check =>
            check.Check == "deliverable.open_xml"
            && check.Status == DeliverableCheckStatus.SkippedPrerequisiteFailed);
    }

    [Fact]
    public void Failed_manifest_safety_boundary_skips_unbounded_downstream_readers()
    {
        var document = ComplexValidDocument();
        var options = new DeliverableVerificationOptions
        {
            PackageManifestOptions = new PackageManifestOptions
            {
                MaxEntryUncompressedBytes = 32,
            },
        };

        var result = DeliverableVerifier.VerifyDeliverable(document.DocumentByteArray, options);

        Assert.False(result.AnalysisCompleted);
        Assert.Equal(DeliverableVerificationDecision.Failed, result.Decision);
        Assert.Contains(result.Findings, finding =>
            finding.Code == "package.entry_size_limit_exceeded");
        Assert.Contains(result.Checks, check =>
            check.Check == "deliverable.open_xml"
            && check.Status == DeliverableCheckStatus.SkippedPrerequisiteFailed);
    }

    [Fact]
    public void Duplicate_relationship_ids_are_reported_without_aborting_baseline_comparison()
    {
        var document = IrTestDocuments.Create("Duplicate relationship fixture");
        var malformed = RewriteEntry(
            document.DocumentByteArray,
            "word/_rels/document.xml.rels",
            xml =>
            {
                var relationships = XDocument.Parse(xml.TrimStart('\uFEFF'));
                var original = relationships.Root!.Elements().First();
                var conflicting = new XElement(original);
                conflicting.SetAttributeValue("Target", "missing-fixture.xml");
                relationships.Root.Add(conflicting);
                return relationships.ToString(SaveOptions.DisableFormatting);
            });

        var result = DeliverableVerifier.VerifyDeliverable(malformed, malformed);

        Assert.True(result.BaselineCompared);
        Assert.Equal(DeliverableVerificationDecision.Failed, result.Decision);
        Assert.Contains(result.Findings, finding =>
            finding.Code == "package.conflicting_relationship"
            && finding.Disposition == DeliverableFindingDisposition.PreExisting);
    }

    [Fact]
    public void Standard_grandfathers_only_unchanged_openxml_errors_while_strict_rejects_them()
    {
        var invalid = OpenXmlInvalidDocument();

        var standard = DeliverableVerifier.VerifyDeliverable(invalid, invalid);
        var strict = DeliverableVerifier.VerifyDeliverable(invalid, invalid,
            new DeliverableVerificationOptions { Mode = DeliverableVerificationMode.Strict });

        Assert.True(standard.AnalysisCompleted, Diagnostics(standard));
        Assert.Equal(DeliverableVerificationDecision.PassedWithPreExistingFindings, standard.Decision);
        Assert.Contains(standard.Findings, finding =>
            finding.Category == DeliverableFindingCategory.OpenXml
            && finding.Disposition == DeliverableFindingDisposition.PreExisting
            && !finding.BlocksDelivery);
        Assert.Equal(DeliverableVerificationDecision.Failed, strict.Decision);
        Assert.Contains(strict.Findings, finding =>
            finding.Category == DeliverableFindingCategory.OpenXml
            && finding.Disposition == DeliverableFindingDisposition.PreExisting
            && finding.BlocksDelivery);
    }

    [Fact]
    public void New_openxml_or_cross_part_defect_blocks_standard_delivery()
    {
        var baseline = IrTestDocuments.Create("Safe").DocumentByteArray;
        var invalid = OpenXmlInvalidDocument();
        var brokenStructure = IrTestDocuments.FromBodyXml(
            "<w:p><w:bookmarkStart w:id=\"4\" w:name=\"MissingEnd\"/>" +
            "<w:r><w:t>Text</w:t></w:r></w:p>").DocumentByteArray;

        var schemaResult = DeliverableVerifier.VerifyDeliverable(baseline, invalid);
        var structureResult = DeliverableVerifier.VerifyDeliverable(baseline, brokenStructure);

        Assert.Equal(DeliverableVerificationDecision.Failed, schemaResult.Decision);
        Assert.Contains(schemaResult.Findings, finding =>
            finding.Category == DeliverableFindingCategory.OpenXml
            && finding.Disposition == DeliverableFindingDisposition.New
            && finding.BlocksDelivery);
        Assert.Equal(DeliverableVerificationDecision.Failed, structureResult.Decision);
        Assert.Contains(structureResult.Findings, finding =>
            finding.Code == "structure.bookmark_pair_invalid"
            && finding.Disposition == DeliverableFindingDisposition.New
            && finding.BlocksDelivery);
    }

    [Fact]
    public void Cross_part_registry_reports_comments_notes_lists_fields_and_content_controls()
    {
        var document = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"8\"/><w:numId w:val=\"99\"/>" +
            "</w:numPr></w:pPr><w:r><w:commentReference w:id=\"7\"/></w:r>" +
            "<w:r><w:footnoteReference w:id=\"8\"/></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>" +
            "<w:p><w:moveFromRangeStart w:id=\"12\"/><w:r><w:t>Moved</w:t></w:r>" +
            "<w:moveToRangeStart w:id=\"12\"/><w:moveToRangeEnd w:id=\"12\"/></w:p>" +
            "<w:sdt><w:sdtPr><w:id w:val=\"2\"/><w:showingPlcHdr/></w:sdtPr>" +
            "<w:sdtContent><w:p><w:r><w:t>Placeholder</w:t></w:r></w:p></w:sdtContent></w:sdt>");

        var result = DeliverableVerifier.VerifyDeliverable(document.DocumentByteArray);

        Assert.Contains(result.Findings, finding => finding.Code == "structure.comment_definition_missing");
        Assert.Contains(result.Findings, finding => finding.Code == "structure.footnote_definition_missing");
        Assert.Contains(result.Findings, finding => finding.Code == "structure.numbering_instance_missing");
        Assert.Contains(result.Findings, finding => finding.Code == "structure.field_sequence_invalid");
        Assert.Contains(result.Findings, finding => finding.Code == "structure.move_range_pair_invalid");
        Assert.Contains(result.Findings, finding => finding.Code == "workflow.content_control_placeholder");
    }

    [Fact]
    public void Placeholder_policy_and_report_only_mode_are_explicit()
    {
        var document = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>[_______]</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>{{CLIENT_NAME}}</w:t></w:r></w:p>");

        var standard = DeliverableVerifier.VerifyDeliverable(document.DocumentByteArray);
        var advisory = DeliverableVerifier.VerifyDeliverable(document.DocumentByteArray,
            new DeliverableVerificationOptions { RequireNoPlaceholders = false });
        var reportOnly = DeliverableVerifier.VerifyDeliverable(document.DocumentByteArray,
            new DeliverableVerificationOptions { Mode = DeliverableVerificationMode.ReportOnly });

        Assert.Equal(DeliverableVerificationDecision.Failed, standard.Decision);
        Assert.Contains(standard.Findings, finding =>
            finding.Code == "workflow.placeholder_remaining" && finding.BlocksDelivery);
        Assert.Equal(DeliverableVerificationDecision.Passed, advisory.Decision);
        Assert.All(advisory.Findings, finding => Assert.False(finding.BlocksDelivery));
        Assert.Equal(DeliverableVerificationDecision.NotEvaluated, reportOnly.Decision);
        Assert.All(reportOnly.Findings, finding => Assert.False(finding.BlocksDelivery));
    }

    [Fact]
    public void Expected_semantic_and_package_delta_must_match_exactly()
    {
        var baseline = IrTestDocuments.Create("Alpha");
        var deliverable = IrTestDocuments.Create("Beta");
        var options = new DeliverableVerificationOptions { FailOnUnexpectedChanges = true };

        var unexpected = DeliverableVerifier.VerifyDeliverable(
            baseline.DocumentByteArray, deliverable.DocumentByteArray, options);
        var expectedSemantic = SemanticDiff.Compare(baseline, deliverable,
            new SemanticDiffOptions { IncludePackageChanges = false });
        var expectedPackage = unexpected.PackageChanges.Select(change =>
            new DeliverablePackageChangeExpectation
            {
                Kind = change.Kind,
                Location = change.Location,
                BeforeDigest = change.BeforeDigest,
                AfterDigest = change.AfterDigest,
                BeforeValue = change.BeforeValue,
                AfterValue = change.AfterValue,
            }).ToArray();

        var approved = DeliverableVerifier.VerifyDeliverable(new DeliverableVerificationRequest
        {
            BaselineBytes = baseline.DocumentByteArray,
            DeliverableBytes = deliverable.DocumentByteArray,
            ExpectedSemanticChanges = expectedSemantic,
            ExpectedPackageChanges = expectedPackage,
        }, options);

        Assert.Equal(DeliverableVerificationDecision.Failed, unexpected.Decision);
        Assert.Contains(unexpected.Findings, finding => finding.Code == "delta.semantic_change_unexpected");
        Assert.Contains(unexpected.Findings, finding => finding.Code == "delta.package_change_unexpected");
        Assert.Equal(DeliverableVerificationDecision.Passed, approved.Decision);
        Assert.DoesNotContain(approved.Findings, finding => finding.Category == DeliverableFindingCategory.Delta);
        Assert.NotNull(approved.SemanticDelta);
        Assert.NotEmpty(approved.PackageChanges);
    }

    [Fact]
    public void Companion_artifacts_are_digest_bound_and_renderer_diagnostics_are_structured()
    {
        var document = IrTestDocuments.Create("Rendered");
        var sourceDigest = PackageManifestGenerator.Generate(document.DocumentByteArray).RawPackageBytesDigest;
        var nonCanonicalSourceDigest = new VerificationDigest
        {
            Algorithm = "sha-256",
            Value = sourceDigest.Value.ToUpperInvariant(),
        };
        var request = new DeliverableVerificationRequest
        {
            DeliverableBytes = document.DocumentByteArray,
            CompanionArtifacts = new[]
            {
                new DeliverableCompanionArtifactInput
                {
                    ArtifactId = "preview.pdf",
                    Role = DeliverableArtifactRole.Pdf,
                    MediaType = "application/pdf",
                    Availability = DeliverableArtifactAvailability.Available,
                    Bytes = Encoding.ASCII.GetBytes("%PDF-1.7 test fixture"),
                    PageCount = 1,
                    RendererFingerprint = "fixture-renderer/1",
                    SourcePackageDigest = nonCanonicalSourceDigest,
                    RenderDiagnostics = new[]
                    {
                        new DeliverableRenderDiagnostic
                        {
                            Kind = DeliverableRenderDiagnosticKind.FontSubstitution,
                            Message = "Fixture Sans was substituted.",
                            FontName = "Missing Fixture Sans",
                            SubstitutedFontName = "Fixture Sans",
                        },
                    },
                },
            },
        };

        var standard = DeliverableVerifier.VerifyDeliverable(request);
        var strict = DeliverableVerifier.VerifyDeliverable(request,
            new DeliverableVerificationOptions { Mode = DeliverableVerificationMode.Strict });

        Assert.Equal(DeliverableVerificationDecision.Passed, standard.Decision);
        Assert.Contains(standard.Findings, finding =>
            finding.Code == "render.font_substitution" && !finding.BlocksDelivery);
        var artifact = Assert.Single(standard.CompanionArtifacts);
        Assert.NotNull(artifact.Digest);
        Assert.Equal("SHA-256", artifact.SourcePackageDigest!.Algorithm);
        Assert.Equal(sourceDigest.Value, artifact.SourcePackageDigest.Value);
        Assert.Equal(1, artifact.RenderDiagnosticCount);
        Assert.Equal(DeliverableVerificationDecision.Failed, strict.Decision);
    }

    [Fact]
    public void Session_entry_point_verifies_current_logical_package_against_opening_bytes()
    {
        using var session = new DocxSession(IrTestDocuments.Create("Alpha").DocumentByteArray);
        var target = Assert.Single(session.FindAllByText("Alpha"));
        Assert.True(Assert.Single(session.ReplaceTextRange(target.Anchor.Id, "Alpha", "Beta")).Success);

        var result = session.VerifyDeliverable();

        Assert.True(result.BaselineCompared);
        Assert.NotNull(result.SemanticDelta);
        Assert.NotEmpty(result.SemanticDelta.Changes);
    }

    [Fact]
    public void Finding_identity_and_disposition_are_stable_across_runs()
    {
        var invalid = OpenXmlInvalidDocument();

        var first = DeliverableVerifier.VerifyDeliverable(invalid, invalid);
        var second = DeliverableVerifier.VerifyDeliverable(invalid, invalid);

        Assert.Equal(
            first.Findings.Select(finding => (finding.FindingId, finding.Disposition)),
            second.Findings.Select(finding => (finding.FindingId, finding.Disposition)));
        Assert.All(first.Findings, finding =>
            Assert.Equal(DeliverableFindingDisposition.PreExisting, finding.Disposition));
    }

    [Fact]
    public void Artifact_bundle_is_emitted_when_requested()
    {
        var directory = Environment.GetEnvironmentVariable("DOCXODUS_DELIVERABLE_ARTIFACT_DIR");
        if (string.IsNullOrWhiteSpace(directory)) return;

        Directory.CreateDirectory(directory);
        var valid = ComplexValidDocument().DocumentByteArray;
        var preInvalid = OpenXmlInvalidDocument();
        var corrupt = Encoding.UTF8.GetBytes("not a zip fixture");
        File.WriteAllBytes(Path.Combine(directory, "valid-complex.docx"), valid);
        File.WriteAllBytes(Path.Combine(directory, "pre-invalid.docx"), preInvalid);
        File.WriteAllBytes(Path.Combine(directory, "corrupt.bin"), corrupt);
        File.WriteAllText(Path.Combine(directory, "valid-report.json"),
            DeliverableVerifier.VerifyDeliverable(valid).ToJson());
        File.WriteAllText(Path.Combine(directory, "pre-invalid-report.json"),
            DeliverableVerifier.VerifyDeliverable(preInvalid, preInvalid).ToJson());
        File.WriteAllText(Path.Combine(directory, "corrupt-report.json"),
            DeliverableVerifier.VerifyDeliverable(corrupt).ToJson());

        var digest = PackageManifestGenerator.Generate(valid).RawPackageBytesDigest;
        var renderReport = DeliverableVerifier.VerifyDeliverable(new DeliverableVerificationRequest
        {
            DeliverableBytes = valid,
            CompanionArtifacts = new[]
            {
                new DeliverableCompanionArtifactInput
                {
                    ArtifactId = "fixture.pdf",
                    Role = DeliverableArtifactRole.Pdf,
                    MediaType = "application/pdf",
                    Availability = DeliverableArtifactAvailability.Available,
                    Bytes = Encoding.ASCII.GetBytes("%PDF-1.7 artifact fixture"),
                    PageCount = 2,
                    RendererFingerprint = "artifact-fixture/1",
                    SourcePackageDigest = digest,
                    RenderDiagnostics = new[]
                    {
                        new DeliverableRenderDiagnostic
                        {
                            Kind = DeliverableRenderDiagnosticKind.MissingFont,
                            Message = "Artifact Fixture Font is unavailable.",
                            FontName = "Artifact Fixture Font",
                        },
                    },
                },
            },
        });
        File.WriteAllText(Path.Combine(directory, "render-diagnostics-report.json"),
            renderReport.ToJson());
    }

    private static byte[] OpenXmlInvalidDocument()
    {
        var document = IrTestDocuments.Create("Invalid schema fixture");
        return RewriteEntry(document.DocumentByteArray, "word/document.xml", xml =>
            xml.Replace("<w:p>", "<w:p w:invalidFixtureAttribute=\"true\">",
                StringComparison.Ordinal));
    }

    private static WmlDocument ComplexValidDocument() => IrTestDocuments.FromParts(
        "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/>" +
        "</w:numPr></w:pPr><w:bookmarkStart w:id=\"1\" w:name=\"ClauseOne\"/>" +
        "<w:r><w:t>Complete clause</w:t></w:r><w:bookmarkEnd w:id=\"1\"/>" +
        "<w:fldSimple w:instr=\"PAGE\"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>" +
        "<w:sdt><w:sdtPr><w:id w:val=\"10\"/><w:tag w:val=\"approved\"/></w:sdtPr>" +
        "<w:sdtContent><w:p><w:r><w:t>Approved content</w:t></w:r></w:p></w:sdtContent></w:sdt>",
        numberingInnerXml:
        "<w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\">" +
        "<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/>" +
        "</w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num>");

    private static byte[] RewriteEntry(byte[] package, string entryName, Func<string, string> rewrite)
    {
        using var output = new MemoryStream();
        using (var destination = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        using (var source = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read))
        {
            foreach (var sourceEntry in source.Entries)
            {
                var destinationEntry = destination.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
                destinationEntry.LastWriteTime = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero);
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

    private static string Diagnostics(DeliverableVerificationResult result) => string.Join("\n",
        result.Checks.Select(check => $"{check.Check}: {check.Status} {check.Diagnostic}")
        .Concat(result.Findings.Select(finding =>
            $"{finding.Severity} {finding.Code}: {finding.Message}")));
}
