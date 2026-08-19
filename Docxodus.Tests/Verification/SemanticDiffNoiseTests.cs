// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

/// <summary>
/// The semantic surface promises "what meaningfully changed": serialization bookkeeping
/// (indentation of Word-owned metadata, rsids, regenerated w14 paragraph ids) must produce
/// zero changes, while genuine vendor facts keep surfacing.
/// </summary>
public class SemanticDiffNoiseTests
{
    private const string AnnotationsCompact =
        "<annotations xmlns=\"http://docxodus.dev/annotations/v1\" version=\"1.0\">" +
        "<annotation id=\"ann-1\" labelId=\"clause\" label=\"Clause\" color=\"#ffff00\">" +
        "<range bookmarkName=\"b1\"/></annotation></annotations>";

    private const string AnnotationsIndented =
        "<annotations xmlns=\"http://docxodus.dev/annotations/v1\" version=\"1.0\">\n" +
        "  <annotation id=\"ann-1\" labelId=\"clause\" label=\"Clause\" color=\"#ffff00\">\n" +
        "    <range bookmarkName=\"b1\"/>\n" +
        "  </annotation>\n" +
        "</annotations>\n";

    [Fact]
    public void Reindenting_the_annotations_part_is_not_a_semantic_change()
    {
        var left = WithCustomXmlPayload(IrTestDocuments.Create("Alpha"), AnnotationsCompact);
        var right = RewriteEntry(left, AnnotationsEntryName(left), AnnotationsIndented);

        var result = SemanticDiff.Compare(left, right);

        Assert.Empty(result.Changes);
    }

    [Fact]
    public void Annotation_attribute_changes_still_surface_after_whitespace_normalization()
    {
        var left = WithCustomXmlPayload(IrTestDocuments.Create("Alpha"), AnnotationsCompact);
        var right = RewriteEntry(left, AnnotationsEntryName(left),
            AnnotationsIndented.Replace("label=\"Clause\"", "label=\"Definition\""));

        var result = SemanticDiff.Compare(left, right);

        var change = Assert.Single(result.Changes);
        Assert.Equal(SemanticChangeFamily.Annotation, change.Family);
    }

    [Fact]
    public void Rsid_only_differences_are_not_semantic_changes()
    {
        const string bodyTemplate =
            "<w:p><w:r><w:t>Stable</w:t></w:r></w:p>" +
            "<w:p><w:ins w:id=\"11\" w:author=\"Ann\" w:date=\"2026-01-01T00:00:00Z\">" +
            "<w:r w:rsidR=\"{0}\"><w:t>Added</w:t></w:r></w:ins></w:p>";
        const string settingsTemplate =
            "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
            "<w:rsids><w:rsidRoot w:val=\"00AAAAAA\"/><w:rsid w:val=\"00AAAAAA\"/>{0}</w:rsids>" +
            "</w:settings>";
        var left = RewriteEntry(
            IrTestDocuments.FromBodyXml(string.Format(bodyTemplate, "00AAAAAA")),
            "word/settings.xml",
            string.Format(settingsTemplate, string.Empty));
        var right = RewriteEntry(
            IrTestDocuments.FromBodyXml(string.Format(bodyTemplate, "00BBBBBB")),
            "word/settings.xml",
            string.Format(settingsTemplate, "<w:rsid w:val=\"00BBBBBB\"/>"));

        var result = SemanticDiff.Compare(left, right);

        Assert.Empty(result.Changes);
    }

    [Fact]
    public void Editing_paragraph_text_does_not_flip_story_extensions()
    {
        const string template =
            "<w:p xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" " +
            "w14:paraId=\"11111111\" w14:textId=\"11111111\"><w:r><w:t>{0}</w:t></w:r></w:p>" +
            "<w:p xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" " +
            "w14:paraId=\"22222222\" w14:textId=\"22222222\"><w:r><w:t>Constant</w:t></w:r></w:p>";
        var left = IrTestDocuments.FromBodyXml(string.Format(template, "hello world"));
        var right = IrTestDocuments.FromBodyXml(string.Format(template, "hello brave world"));

        var result = SemanticDiff.Compare(left, right);

        Assert.Contains(result.Changes, change => change.Family == SemanticChangeFamily.Text);
        Assert.DoesNotContain(result.Changes, change =>
            change.Path.StartsWith("story.extensions", StringComparison.Ordinal));
    }

    [Fact]
    public void Vendor_extension_attribute_changes_still_surface()
    {
        const string template =
            "<w:p xmlns:v=\"urn:example:vendor\" v:flag=\"{0}\">" +
            "<w:r><w:t>Constant</w:t></w:r></w:p>";
        var left = IrTestDocuments.FromBodyXml(string.Format(template, "alpha"));
        var right = IrTestDocuments.FromBodyXml(string.Format(template, "beta"));

        var result = SemanticDiff.Compare(left, right);

        var change = Assert.Single(result.Changes);
        Assert.Equal("story.extensions.package", change.Path);
    }

    [Fact]
    public void Editing_text_around_a_vendor_extension_does_not_flip_story_extensions()
    {
        const string template =
            "<w:p xmlns:v=\"urn:example:vendor\" v:flag=\"pinned\">" +
            "<w:r><w:t>{0}</w:t></w:r></w:p>";
        var left = IrTestDocuments.FromBodyXml(string.Format(template, "hello world"));
        var right = IrTestDocuments.FromBodyXml(string.Format(template, "hello brave world"));

        var result = SemanticDiff.Compare(left, right);

        Assert.Contains(result.Changes, change => change.Family == SemanticChangeFamily.Text);
        Assert.DoesNotContain(result.Changes, change =>
            change.Path.StartsWith("story.extensions", StringComparison.Ordinal));
    }

    private static WmlDocument WithCustomXmlPayload(WmlDocument source, string payload)
    {
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var partStream = part.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(partStream, Encoding.UTF8, 1024, leaveOpen: false);
            writer.Write(payload);
        }
        return new WmlDocument(source.FileName, stream.ToArray());
    }

    private static string AnnotationsEntryName(WmlDocument source)
    {
        using var stream = new MemoryStream(source.DocumentByteArray, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.Entries.First(entry =>
            entry.FullName.StartsWith("customXml/item", StringComparison.Ordinal)
            && !entry.FullName.Contains("itemProps", StringComparison.Ordinal)).FullName;
    }

    private static WmlDocument RewriteEntry(WmlDocument source, string entryName, string xml)
    {
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry(entryName)
                ?? throw new InvalidOperationException($"missing entry {entryName}");
            using var target = entry.Open();
            target.SetLength(0);
            var bytes = Encoding.UTF8.GetBytes(xml);
            target.Write(bytes);
        }
        return new WmlDocument(source.FileName, stream.ToArray());
    }
}
