// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

/// <summary>
/// The semantic surface must behave like its sibling DocxDiff entry points: same input
/// normalization, same compatibility gate, same package safety policy, and no change class the
/// redline engine reports that the audit surface silently drops.
/// </summary>
public class SemanticDiffEngineParityTests
{
    private const string TransitionalMain = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StrictMain = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    private const string TransitionalRels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string StrictRels = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    [Fact]
    public void Strict_conformance_packages_compare_like_the_sibling_engine()
    {
        var left = ToStrict(IrTestDocuments.Create("Hello world."));
        var right = ToStrict(IrTestDocuments.Create("Hello brave world."));

        var result = SemanticDiff.Compare(left, right);

        Assert.Contains(result.Changes, change => change.Family == SemanticChangeFamily.Text);
    }

    [Fact]
    public void Semantic_surface_honors_the_compatibility_gate()
    {
        var left = IrTestDocuments.FromBodyXml(
            $"<w:p><m:oMath xmlns:m=\"{M}\"><m:r><m:t>x</m:t></m:r></m:oMath></w:p>");
        var right = IrTestDocuments.Create("plain");
        var options = new SemanticDiffOptions
        {
            DiffSettings = new DocxDiffSettings { ThrowOnCompatibilityWarning = true },
        };

        var ex = Assert.Throws<DocxDiffCompatibilityException>(() =>
            SemanticDiff.Compare(left, right, options));

        Assert.Contains(ex.Report.Warnings, warning => warning.Feature.Id == "math");
    }

    [Fact]
    public void Compatibility_callback_fires_from_the_semantic_surface()
    {
        var left = IrTestDocuments.FromBodyXml(
            $"<w:p><m:oMath xmlns:m=\"{M}\"><m:r><m:t>x</m:t></m:r></m:oMath></w:p>");
        var right = IrTestDocuments.Create("plain");
        DocxDiffCompatibilityReport? captured = null;
        var options = new SemanticDiffOptions
        {
            DiffSettings = new DocxDiffSettings { OnCompatibilityWarning = r => captured = r },
        };

        DocxDiff.GetSemanticChanges(left, right, options);

        Assert.NotNull(captured);
    }

    [Fact]
    public void Semantic_package_limits_default_to_the_manifest_policy()
    {
        Assert.Equal(new PackageManifestOptions(), new SemanticDiffOptions().PackageOptions);
    }

    [Fact]
    public void Atomic_token_format_changes_survive_carrier_rewrites()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">alpha </w:t></w:r>" +
            "<w:r><w:tab/></w:r><w:r><w:t>beta</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">alpha </w:t></w:r>" +
            "<w:r><w:rPr><w:b/></w:rPr><w:tab/></w:r>" +
            "<w:hyperlink w:anchor=\"target\"><w:r><w:t>beta</w:t></w:r></w:hyperlink></w:p>");

        var result = SemanticDiff.Compare(left, right);

        Assert.Contains(result.Changes, change =>
            change.Family == SemanticChangeFamily.RunFormatting
            && change.Path.Contains("atomic_tokens", StringComparison.Ordinal));
    }

    [Fact]
    public void Header_story_paths_pin_the_kind_vocabulary()
    {
        var left = IrTestDocuments.WithHeaderAndFooter("Old header", "Footer");
        var right = IrTestDocuments.WithHeaderAndFooter("New header", "Footer");

        var result = SemanticDiff.Compare(left, right);

        var story = Assert.Single(result.Changes, change =>
            change.Path.StartsWith("header[section=", StringComparison.Ordinal));
        // The path grammar is v1 wire output: the kind token is pinned to the w:type
        // vocabulary (default/first/even), never an enum member's ToString().
        Assert.Contains("kind=default", story.Path, StringComparison.Ordinal);
    }

    private static WmlDocument ToStrict(WmlDocument doc)
    {
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries.ToList())
            {
                if (!entry.FullName.EndsWith(".xml", StringComparison.Ordinal)
                    && !entry.FullName.EndsWith(".rels", StringComparison.Ordinal))
                    continue;
                string text;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    text = reader.ReadToEnd();
                var rewritten = text
                    .Replace(TransitionalMain, StrictMain)
                    .Replace(TransitionalRels, StrictRels);
                if (rewritten == text) continue;
                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.BaseStream.SetLength(0);
                writer.Write(rewritten);
            }
        }
        return new WmlDocument("strict.docx", ms.ToArray());
    }
}
