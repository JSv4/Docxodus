// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

public class SemanticDiffProjectionRegressionTests
{
    private const string RelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Fact]
    public void Equal_text_with_changed_run_boundary_reports_character_range_formatting()
    {
        var left = IrTestDocuments.FromBodyXmlWithHyperlinks(
            $"<w:p><w:hyperlink xmlns:r=\"{RelationshipNamespace}\" r:id=\"rIdLink\">" +
            "<w:r><w:rPr><w:b/></w:rPr><w:t>AB</w:t></w:r>" +
            "</w:hyperlink></w:p>",
            ("rIdLink", "https://example.test/same"));
        var right = IrTestDocuments.FromBodyXmlWithHyperlinks(
            $"<w:p><w:hyperlink xmlns:r=\"{RelationshipNamespace}\" r:id=\"rIdLink\">" +
            "<w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r>" +
            "<w:r><w:rPr><w:i/></w:rPr><w:t>B</w:t></w:r>" +
            "</w:hyperlink></w:p>",
            ("rIdLink", "https://example.test/same"));

        var result = SemanticDiff.Compare(left, right,
            new SemanticDiffOptions { IncludePackageChanges = false });

        var formatting = Assert.Single(result.Changes, change =>
            change.Family == SemanticChangeFamily.RunFormatting);
        Assert.Equal(SemanticChangeOperation.Modify, formatting.Operation);
        Assert.Contains("paragraph.characters[1:2]", formatting.Path);
        Assert.True(BooleanProperty(formatting.Before, "bold"));
        Assert.True(BooleanProperty(formatting.After, "italic"));
        Assert.DoesNotContain(result.Changes, change => change.Family == SemanticChangeFamily.Text);
        AssertModifyValuesDiffer(result);
    }

    [Fact]
    public void Coordinated_note_id_churn_has_no_note_envelope_but_nested_edits_survive()
    {
        var left = IrTestDocuments.WithFootnoteAndEndnote("Same footnote", "Same endnote");
        var renumbered = RenumberNotes(
            IrTestDocuments.WithFootnoteAndEndnote("Same footnote", "Same endnote"),
            footnoteId: 41,
            endnoteId: 73);

        var idOnly = SemanticDiff.Compare(left, renumbered);

        Assert.DoesNotContain(idOnly.Changes, change =>
            change.Family is SemanticChangeFamily.Footnote or SemanticChangeFamily.Endnote);

        var editedAndRenumbered = RenumberNotes(
            IrTestDocuments.WithFootnoteAndEndnote("Edited footnote", "Same endnote"),
            footnoteId: 41,
            endnoteId: 73);
        var edited = SemanticDiff.Compare(left, editedAndRenumbered);

        Assert.Contains(edited.Changes, change =>
            change.Family == SemanticChangeFamily.Text
            && change.PartUri == "/word/footnotes.xml");
        Assert.DoesNotContain(edited.Changes, change =>
            change.Family == SemanticChangeFamily.Endnote);
        AssertModifyValuesDiffer(edited);
    }

    [Fact]
    public void Header_and_footer_envelopes_carry_side_specific_content_and_topology()
    {
        var left = IrTestDocuments.WithHeaderAndFooter("Old header", "Old footer");
        var right = IrTestDocuments.WithHeaderAndFooter("New header", "New footer");

        var result = SemanticDiff.Compare(left, right);
        var envelopes = result.Changes.Where(change =>
            change.Family is SemanticChangeFamily.Header or SemanticChangeFamily.Footer).ToArray();

        Assert.Equal(2, envelopes.Length);
        Assert.All(envelopes, change =>
        {
            Assert.Equal(SemanticChangeOperation.Modify, change.Operation);
            Assert.Contains(change.Before.Properties, property => property.Name == "blocks");
            Assert.Contains(change.After.Properties, property => property.Name == "blocks");
            Assert.NotEqual(ValueFingerprint(change.Before), ValueFingerprint(change.After));
        });
        Assert.Contains(result.Changes, change =>
            change.Family == SemanticChangeFamily.Text
            && change.PartUri.StartsWith("/word/header", System.StringComparison.Ordinal));
        Assert.Contains(result.Changes, change =>
            change.Family == SemanticChangeFamily.Text
            && change.PartUri.StartsWith("/word/footer", System.StringComparison.Ordinal));
        AssertModifyValuesDiffer(result);
    }

    [Fact]
    public void Body_changes_use_the_actual_nonstandard_main_part_uri()
    {
        var left = RenameMainPart(IrTestDocuments.Create("Before"));
        var right = RenameMainPart(IrTestDocuments.Create("After"));

        var result = SemanticDiff.Compare(left, right,
            new SemanticDiffOptions { IncludePackageChanges = false });

        Assert.Contains(result.Changes, change =>
            change.Family == SemanticChangeFamily.Text
            && change.PartUri == "/word/main.xml");
        Assert.DoesNotContain(result.Changes, change =>
            change.PartUri == "/word/document.xml");
        AssertModifyValuesDiffer(result);
    }

    [Fact]
    public void Registry_changes_use_the_actual_nonstandard_part_uri()
    {
        const string body = "<w:p><w:r><w:t>Same</w:t></w:r></w:p>";
        var left = RenameStylesPart(IrTestDocuments.FromParts(body,
            "<w:style w:type=\"paragraph\" w:styleId=\"Clause\">" +
            "<w:name w:val=\"Old\"/></w:style>"));
        var right = RenameStylesPart(IrTestDocuments.FromParts(body,
            "<w:style w:type=\"paragraph\" w:styleId=\"Clause\">" +
            "<w:name w:val=\"New\"/></w:style>"));

        var result = SemanticDiff.Compare(left, right,
            new SemanticDiffOptions { IncludePackageChanges = false });

        Assert.Contains(result.Changes, change =>
            change.Family == SemanticChangeFamily.Style
            && change.PartUri == "/word/config/styles-alt.xml");
        Assert.DoesNotContain(result.Changes, change =>
            change.PartUri == "/word/styles.xml");
        AssertModifyValuesDiffer(result);
    }

    private static bool? BooleanProperty(SemanticValue value, string name) =>
        value.Properties.Single(property => property.Name == name).Value.BooleanValue;

    private static void AssertModifyValuesDiffer(SemanticChangeSet result) =>
        Assert.All(result.Changes.Where(change => change.Operation == SemanticChangeOperation.Modify),
            change => Assert.NotEqual(
                ValueFingerprint(change.Before),
                ValueFingerprint(change.After)));

    private static string ValueFingerprint(SemanticValue value) =>
        JsonSerializer.Serialize(value);

    private static WmlDocument RenumberNotes(
        WmlDocument source,
        int footnoteId,
        int endnoteId)
    {
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            foreach (var reference in main.Document.Descendants<FootnoteReference>())
                reference.Id = footnoteId;
            foreach (var reference in main.Document.Descendants<EndnoteReference>())
                reference.Id = endnoteId;

            var footnote = main.FootnotesPart!.Footnotes!
                .Elements<Footnote>()
                .Single(note => note.Id?.Value == 1);
            footnote.Id = footnoteId;
            var endnote = main.EndnotesPart!.Endnotes!
                .Elements<Endnote>()
                .Single(note => note.Id?.Value == 1);
            endnote.Id = endnoteId;

            main.Document.Save();
            main.FootnotesPart.Footnotes.Save();
            main.EndnotesPart.Endnotes.Save();
        }

        return new WmlDocument("semantic-projection-renumbered.docx", stream.ToArray());
    }

    private static WmlDocument RenameMainPart(WmlDocument source)
    {
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            RenameEntry(archive, "word/document.xml", "word/main.xml");
            RenameEntry(archive, "word/_rels/document.xml.rels", "word/_rels/main.xml.rels");
            RewriteEntry(archive, "[Content_Types].xml",
                xml => xml.Replace("/word/document.xml", "/word/main.xml",
                    System.StringComparison.Ordinal));
            RewriteEntry(archive, "_rels/.rels",
                xml => xml.Replace("word/document.xml", "word/main.xml",
                    System.StringComparison.Ordinal));
        }

        return new WmlDocument("semantic-projection-main-part.docx", stream.ToArray());
    }

    private static WmlDocument RenameStylesPart(WmlDocument source)
    {
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            RenameEntry(archive, "word/styles.xml", "word/config/styles-alt.xml");
            RewriteEntry(archive, "[Content_Types].xml",
                xml => xml.Replace("/word/styles.xml", "/word/config/styles-alt.xml",
                    System.StringComparison.Ordinal));
            RewriteEntry(archive, "word/_rels/document.xml.rels",
                xml => xml.Replace("styles.xml", "config/styles-alt.xml",
                    System.StringComparison.Ordinal));
        }

        return new WmlDocument("semantic-projection-styles-part.docx", stream.ToArray());
    }

    private static void RenameEntry(ZipArchive archive, string sourceName, string targetName)
    {
        var source = archive.GetEntry(sourceName)!;
        using var bytes = new MemoryStream();
        using (var input = source.Open())
            input.CopyTo(bytes);
        source.Delete();

        var target = archive.CreateEntry(targetName, CompressionLevel.Optimal);
        using var output = target.Open();
        bytes.Position = 0;
        bytes.CopyTo(output);
    }

    private static void RewriteEntry(
        ZipArchive archive,
        string entryName,
        System.Func<string, string> rewrite)
    {
        var entry = archive.GetEntry(entryName)!;
        string xml;
        using (var reader = new StreamReader(entry.Open(), Encoding.UTF8, true))
            xml = reader.ReadToEnd();
        entry.Delete();

        var replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var writer = new StreamWriter(
            replacement.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(rewrite(xml));
    }
}
