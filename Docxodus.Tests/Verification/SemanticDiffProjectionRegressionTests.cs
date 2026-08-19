// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
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

    [Fact]
    public void Out_of_range_document_integers_project_as_lossless_strings()
    {
        // wp:extent/@cx and w:gridCol/@w parse as unbounded longs, so a crafted package can carry a
        // value above the v1 safe integer range. Projecting it must not turn the verification
        // surface into a crash oracle, and the two sides must stay distinguishable.
        const long aboveSafeRange = 9_007_199_254_740_992L;

        var imageLeft = IrTestDocuments.FromBodyXmlWithImageParts(
            InlineImageXml("rIdImage", aboveSafeRange), ("rIdImage", IrTestDocuments.TinyPng));
        var imageRight = IrTestDocuments.FromBodyXmlWithImageParts(
            InlineImageXml("rIdImage", aboveSafeRange + 1), ("rIdImage", IrTestDocuments.TinyPng));

        var images = SemanticDiff.Compare(imageLeft, imageRight,
            new SemanticDiffOptions { IncludePackageChanges = false });
        var image = Assert.Single(images.Changes, change => change.Family == SemanticChangeFamily.Image);
        Assert.Equal("9007199254740992", StringProperty(image.Before, "widthEmu"));
        Assert.Equal("9007199254740993", StringProperty(image.After, "widthEmu"));

        var gridLeft = IrTestDocuments.FromBodyXml(SingleColumnTableXml(aboveSafeRange));
        var gridRight = IrTestDocuments.FromBodyXml(SingleColumnTableXml(aboveSafeRange + 1));

        var grids = SemanticDiff.Compare(gridLeft, gridRight,
            new SemanticDiffOptions { IncludePackageChanges = false });
        var grid = Assert.Single(grids.Changes, change => change.Path == "table.grid");
        Assert.Equal(new[] { "9007199254740992" }, GridColumns(grid.Before));
        Assert.Equal(new[] { "9007199254740993" }, GridColumns(grid.After));
    }

    [Fact]
    public void Nested_table_properties_never_project_onto_the_outer_table()
    {
        // The outer table's own grid and borders change; the nested table's style, width, and grid
        // change independently. Every record the outer anchor carries must describe the outer table.
        var left = IrTestDocuments.FromBodyXml(
            NestedTableXml(outerGridTwips: 1000, borderSize: 4, nestedStyle: "NestedA", nestedWidth: 800));
        var right = IrTestDocuments.FromBodyXml(
            NestedTableXml(outerGridTwips: 1200, borderSize: 8, nestedStyle: "NestedB", nestedWidth: 900));

        var result = SemanticDiff.Compare(left, right,
            new SemanticDiffOptions { IncludePackageChanges = false });

        // Only the outer w:tblGrid changed, so exactly one grid record exists and it identifies the
        // outer table. Its columns are the outer three, not the outer three plus the nested two.
        var grid = Assert.Single(result.Changes, change => change.Path == "table.grid");
        Assert.Equal(new[] { "1000", "1000", "1000" }, GridColumns(grid.Before));
        Assert.Equal(new[] { "1200", "1200", "1200" }, GridColumns(grid.After));

        // The outer w:tblPr declares no w:tblStyle and no w:tblW; a descendant sweep used to borrow
        // the nested table's values and file them under the outer table's anchor.
        var outerAnchor = grid.LeftAnchor;
        Assert.Equal(
            new[] { "table", "table.grid", "table.properties" },
            result.Changes
                .Where(change => change.LeftAnchor == outerAnchor || change.RightAnchor == outerAnchor)
                .Select(change => change.Path)
                .Distinct()
                .OrderBy(path => path, System.StringComparer.Ordinal)
                .ToArray());

        // The nested table still reports its own style and width against its own anchor.
        var nestedStyle = Assert.Single(result.Changes, change => change.Path == "table.style");
        Assert.NotEqual(outerAnchor, nestedStyle.LeftAnchor);
        Assert.Equal("NestedA", nestedStyle.Before.StringValue);
        Assert.Equal("NestedB", nestedStyle.After.StringValue);
        var nestedWidth = Assert.Single(result.Changes, change => change.Path == "table.width");
        Assert.NotEqual(outerAnchor, nestedWidth.LeftAnchor);
        Assert.Equal("800", StringProperty(nestedWidth.Before, "w"));
        Assert.Equal("900", StringProperty(nestedWidth.After, "w"));
        AssertModifyValuesDiffer(result);
    }

    private static string InlineImageXml(string relId, long widthEmu) =>
        "<w:p><w:r><w:drawing>" +
        "<wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" +
        $"<wp:extent cx=\"{widthEmu}\" cy=\"95250\"/><wp:docPr id=\"1\" name=\"Image\"/>" +
        "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
        "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "<pic:nvPicPr><pic:cNvPr id=\"1\" name=\"Image\"/><pic:cNvPicPr/></pic:nvPicPr>" +
        $"<pic:blipFill><a:blip xmlns:r=\"{RelationshipNamespace}\" r:embed=\"{relId}\"/>" +
        "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
        "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"95250\" cy=\"95250\"/></a:xfrm>" +
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>" +
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    private static string SingleColumnTableXml(long columnTwips) =>
        "<w:tbl><w:tblPr/>" +
        $"<w:tblGrid><w:gridCol w:w=\"{columnTwips}\"/></w:tblGrid>" +
        "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";

    private static string NestedTableXml(
        int outerGridTwips,
        int borderSize,
        string nestedStyle,
        int nestedWidth) =>
        "<w:tbl>" +
        $"<w:tblPr><w:tblBorders><w:top w:val=\"single\" w:sz=\"{borderSize}\"/></w:tblBorders></w:tblPr>" +
        "<w:tblGrid>" +
        string.Concat(Enumerable.Repeat($"<w:gridCol w:w=\"{outerGridTwips}\"/>", 3)) +
        "</w:tblGrid>" +
        "<w:tr><w:tc><w:tcPr/>" +
        "<w:tbl>" +
        $"<w:tblPr><w:tblStyle w:val=\"{nestedStyle}\"/>" +
        $"<w:tblW w:w=\"{nestedWidth}\" w:type=\"dxa\"/></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"500\"/><w:gridCol w:w=\"500\"/></w:tblGrid>" +
        $"<w:tr><w:tc><w:tcPr><w:tcW w:w=\"{nestedWidth}\" w:type=\"dxa\"/></w:tcPr>" +
        "<w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:tc>" +
        "<w:tc><w:tcPr><w:tcW w:w=\"500\" w:type=\"dxa\"/></w:tcPr>" +
        "<w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>" +
        "</w:tbl>" +
        "<w:p><w:r><w:t>Outer</w:t></w:r></w:p></w:tc>" +
        "<w:tc><w:tcPr/><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>" +
        "<w:tc><w:tcPr/><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr>" +
        "</w:tbl>";

    private static string?[] GridColumns(SemanticValue value) =>
        value.Properties.Single(property => property.Name == "columnsTwips").Value.Items
            .Select(item => item.StringValue ?? item.IntegerValue?.ToString(CultureInfo.InvariantCulture))
            .ToArray();

    private static string? StringProperty(SemanticValue value, string name) =>
        value.Properties.Single(property => property.Name == name).Value.StringValue;

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
