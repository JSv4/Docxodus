// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.Tests.Ir;
using Docxodus.Tests.Ir.Diff;
using Docxodus.Verification;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Verification;

public class SemanticDiffTests
{
    private const string W = IrTestDocuments.W;
    private readonly ITestOutputHelper _output;

    public SemanticDiffTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public void Identical_package_has_no_semantic_changes()
    {
        var document = IrTestDocuments.Create("Alpha", "Beta");

        var result = SemanticDiff.Compare(document, new WmlDocument(document));

        Assert.Equal(SemanticChangeSet.CurrentSchema, result.Schema);
        Assert.Equal(SemanticChangeSet.CurrentSchemaVersion, result.SchemaVersion);
        Assert.Empty(result.Changes);
        Assert.Equal(result.ToCanonicalJson(), result.ToJson(indented: false));
    }

    [Fact]
    public void Text_run_paragraph_style_list_and_numbering_are_typed()
    {
        var left = IrTestDocuments.FromParts(
            "<w:p><w:pPr><w:pStyle w:val=\"Clause\"/><w:numPr><w:ilvl w:val=\"0\"/>" +
            "<w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>Alpha text</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Formatting</w:t></w:r></w:p>",
            "<w:style w:type=\"paragraph\" w:styleId=\"Clause\"><w:name w:val=\"Clause\"/></w:style>",
            Numbering("decimal", 1));
        var right = IrTestDocuments.FromParts(
            "<w:p><w:pPr><w:pStyle w:val=\"Clause2\"/><w:jc w:val=\"center\"/>" +
            "<w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr>" +
            "<w:r><w:t>Revised text</w:t></w:r></w:p>" +
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Formatting</w:t></w:r></w:p>",
            "<w:style w:type=\"paragraph\" w:styleId=\"Clause2\"><w:name w:val=\"Clause Two\"/>" +
            "<w:pPr><w:keepNext/></w:pPr></w:style>",
            Numbering("lowerLetter", 2));

        var result = SemanticDiff.Compare(left, right);

        AssertFamilies(result,
            SemanticChangeFamily.Text,
            SemanticChangeFamily.RunFormatting,
            SemanticChangeFamily.ParagraphFormatting,
            SemanticChangeFamily.Style,
            SemanticChangeFamily.List,
            SemanticChangeFamily.Numbering);
        Assert.All(result.Changes, AssertCompleteLocationAndValues);
        var digests = result.Changes
            .SelectMany(change => DescendantValues(change.Before).Concat(DescendantValues(change.After)))
            .Where(value => value.Kind == SemanticValueKind.Digest)
            .ToArray();
        Assert.NotEmpty(digests);
        Assert.All(digests, digest =>
        {
            Assert.Equal("SHA-256", digest.DigestAlgorithm);
            Assert.False(string.IsNullOrWhiteSpace(digest.DigestProfile));
        });
    }

    [Fact]
    public void Table_geometry_styles_rows_cells_and_page_setup_are_typed()
    {
        var left = IrTestDocuments.FromBodyXml(
            Table("A", "Grid", 1, 1000, includeSecondRow: false, rowHeight: 200) +
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:left=\"1440\"/></w:sectPr>");
        var right = IrTestDocuments.FromBodyXml(
            Table("B", "ColorfulGrid", 2, 2400, includeSecondRow: true, rowHeight: 400) +
            "<w:sectPr><w:pgSz w:w=\"15840\" w:h=\"12240\" w:orient=\"landscape\"/>" +
            "<w:pgMar w:left=\"720\"/></w:sectPr>");

        var result = SemanticDiff.Compare(left, right);

        AssertFamilies(result,
            SemanticChangeFamily.Table,
            SemanticChangeFamily.TableRow,
            SemanticChangeFamily.TableCell,
            SemanticChangeFamily.TableSpan,
            SemanticChangeFamily.TableWidth,
            SemanticChangeFamily.TableStyle,
            SemanticChangeFamily.Section,
            SemanticChangeFamily.PageSetup);
    }

    [Fact]
    public void Mixed_story_note_and_comment_parts_keep_owning_part_and_scope()
    {
        var storyLeft = IrTestDocuments.WithHeaderAndFooter("Old header", "Old footer");
        var storyRight = IrTestDocuments.WithHeaderAndFooter("New header", "New footer");
        var stories = SemanticDiff.Compare(storyLeft, storyRight);
        AssertFamilies(stories, SemanticChangeFamily.Header, SemanticChangeFamily.Footer);
        Assert.Contains(stories.Changes, change => change.PartUri.StartsWith("/word/header", StringComparison.Ordinal));
        Assert.Contains(stories.Changes, change => change.PartUri.StartsWith("/word/footer", StringComparison.Ordinal));

        var noteLeft = IrTestDocuments.WithFootnoteAndEndnote("Old footnote", "Old endnote");
        var noteRight = IrTestDocuments.WithFootnoteAndEndnote("New footnote", "New endnote");
        var notes = SemanticDiff.Compare(noteLeft, noteRight);
        AssertFamilies(notes, SemanticChangeFamily.Footnote, SemanticChangeFamily.Endnote);
        Assert.Contains(notes.Changes, change => change.PartUri == "/word/footnotes.xml");
        Assert.Contains(notes.Changes, change => change.PartUri == "/word/endnotes.xml");

        const string body = "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t>Target</w:t></w:r>" +
            "<w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>";
        var commentLeft = IrTestDocuments.WithComment("Alice", "A", "2026-01-01T00:00:00Z", "Old", body);
        var commentRight = IrTestDocuments.WithComment("Bob", "B", "2026-01-02T00:00:00Z", "New", body);
        var comments = SemanticDiff.Compare(commentLeft, commentRight);
        AssertFamilies(comments, SemanticChangeFamily.Comment);
        Assert.Contains(comments.Changes, change => change.PartUri == "/word/comments.xml");
    }

    [Fact]
    public void Fields_links_bookmarks_content_controls_images_media_and_relationships_are_visible()
    {
        var complexLeft = IrTestDocuments.FromBodyXmlWithHyperlinks(
            "<w:p><w:hyperlink r:id=\"rIdLink\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>Link</w:t></w:r></w:hyperlink>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"OldBookmark\"/><w:bookmarkEnd w:id=\"1\"/>" +
            "<w:fldSimple w:instr=\"PAGE\"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>" +
            "<w:sdt><w:sdtPr><w:tag w:val=\"old\"/></w:sdtPr><w:sdtContent>" +
            "<w:p><w:r><w:t>Controlled</w:t></w:r></w:p></w:sdtContent></w:sdt>",
            ("rIdLink", "https://example.test/old"));
        var complexRight = IrTestDocuments.FromBodyXmlWithHyperlinks(
            "<w:p><w:hyperlink r:id=\"rIdLink\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>Link</w:t></w:r></w:hyperlink>" +
            "<w:bookmarkStart w:id=\"77\" w:name=\"NewBookmark\"/><w:bookmarkEnd w:id=\"77\"/>" +
            "<w:fldSimple w:instr=\"NUMPAGES\"><w:r><w:t>2</w:t></w:r></w:fldSimple></w:p>" +
            "<w:sdt><w:sdtPr><w:tag w:val=\"new\"/></w:sdtPr><w:sdtContent>" +
            "<w:p><w:r><w:t>Controlled</w:t></w:r></w:p></w:sdtContent></w:sdt>",
            ("rIdLink", "https://example.test/new"));
        var complex = SemanticDiff.Compare(complexLeft, complexRight);
        AssertFamilies(complex,
            SemanticChangeFamily.Field,
            SemanticChangeFamily.Hyperlink,
            SemanticChangeFamily.Bookmark,
            SemanticChangeFamily.ContentControl,
            SemanticChangeFamily.Relationship);

        var imageXml = HeaderFooterFixtures.ImageParagraphXml("rIdImage");
        var imageLeft = IrTestDocuments.FromBodyXmlWithImageParts(imageXml,
            ("rIdImage", IrTestDocuments.TinyPng));
        var changedPng = IrTestDocuments.TinyPng.Concat(new byte[] { 1, 2, 3 }).ToArray();
        var imageRight = IrTestDocuments.FromBodyXmlWithImageParts(imageXml,
            ("rIdImage", changedPng));
        var images = SemanticDiff.Compare(imageLeft, imageRight);
        AssertFamilies(images, SemanticChangeFamily.Image, SemanticChangeFamily.Media);
    }

    [Fact]
    public void Revision_annotation_and_unknown_parts_never_disappear()
    {
        var plain = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Text</w:t></w:r></w:p>");
        var revised = IrTestDocuments.FromBodyXml(
            "<w:p><w:ins w:id=\"4\" w:author=\"Reviewer\" w:date=\"2026-01-01T00:00:00Z\">" +
            "<w:r><w:t>Text</w:t></w:r></w:ins></w:p>");
        AssertFamilies(SemanticDiff.Compare(plain, revised), SemanticChangeFamily.Revision);

        var annotated = WithCustomXml(plain, annotation: true);
        AssertFamilies(SemanticDiff.Compare(plain, annotated),
            SemanticChangeFamily.Annotation, SemanticChangeFamily.Relationship);

        var opaque = WithCustomXml(plain, annotation: false);
        AssertFamilies(SemanticDiff.Compare(plain, opaque),
            SemanticChangeFamily.OpaquePackagePart, SemanticChangeFamily.Relationship);
    }

    [Fact]
    public void Serialization_only_xml_and_relationship_id_churn_are_suppressed()
    {
        var left = IrTestDocuments.FromBodyXmlWithHyperlinks(
            "<w:p><w:hyperlink r:id=\"rIdA\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>Same</w:t></w:r></w:hyperlink></w:p>",
            ("rIdA", "https://example.test/same"));
        var right = IrTestDocuments.FromBodyXmlWithHyperlinks(
            "<w:p><w:hyperlink r:id=\"rIdZ\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>Same</w:t></w:r></w:hyperlink></w:p>",
            ("rIdZ", "https://example.test/same"));
        right = RewriteEntry(right, "word/settings.xml",
            $"<x:settings xmlns:x=\"{W}\">\n</x:settings>");

        var result = SemanticDiff.Compare(left, right);

        Assert.Empty(result.Changes);
    }

    [Fact]
    public void Opaque_custom_xml_preserves_whitespace_as_semantic_data()
    {
        var basis = IrTestDocuments.Create("Same body");
        var compact = WithCustomXmlPayload(basis,
            "<vendorData xmlns=\"urn:example:vendor\"><value>A</value><value>B</value></vendorData>");
        var spaced = WithCustomXmlPayload(basis,
            "<vendorData xmlns=\"urn:example:vendor\"><value>A</value> <value>B</value></vendorData>");

        var result = SemanticDiff.Compare(compact, spaced);

        Assert.Contains(result.Changes, change =>
            change.Family == SemanticChangeFamily.OpaquePackagePart);
    }

    [Fact]
    public void Opaque_part_uses_declared_content_type_as_semantic_data()
    {
        var basis = IrTestDocuments.Create("Same body");
        var left = WithCustomXmlPayload(basis,
            "<vendorData xmlns=\"urn:example:vendor\"><value>A</value></vendorData>");
        var right = RewriteCustomXmlContentType(left, "application/vnd.example.semantic+xml");

        var change = Assert.Single(SemanticDiff.Compare(left, right).Changes.Where(item =>
            item.Family == SemanticChangeFamily.OpaquePackagePart
            && item.Operation == SemanticChangeOperation.Modify));

        var beforeType = Assert.Single(change.Before.Properties,
            property => property.Name == "contentType").Value.StringValue;
        var afterType = Assert.Single(change.After.Properties,
            property => property.Name == "contentType").Value.StringValue;
        Assert.NotEqual(beforeType, afterType);
        Assert.Equal("application/vnd.example.semantic+xml", afterType);
    }

    [Fact]
    public void Manifest_preflight_applies_when_package_supplement_is_disabled()
    {
        var document = IrTestDocuments.Create("Same");
        var invalid = new WmlDocument(document);
        Array.Clear(invalid.DocumentByteArray);

        Assert.Throws<InvalidDataException>(() => SemanticDiff.Compare(
            invalid,
            document,
            new SemanticDiffOptions { IncludePackageChanges = false }));
    }

    [Fact]
    public void Internal_relationship_target_spellings_are_owner_resolved_before_comparison()
    {
        var basis = IrTestDocuments.Create("Same");
        var relative = WithCustomInternalRelationship(
            basis, "itemProps1.xml");
        var absolute = WithCustomInternalRelationship(
            basis, "/customXml/itemProps1.xml");

        Assert.Empty(SemanticDiff.Compare(relative, absolute).Changes);
    }

    [Fact]
    public void Canonical_order_ids_and_json_are_independent_of_input_order()
    {
        var a = Change("caller-z", "/word/z.xml", SemanticChangeFamily.Text, "z");
        var b = Change("caller-a", "/word/a.xml", SemanticChangeFamily.Style, "a");

        var forward = new SemanticChangeSet(new[] { a, b });
        var reverse = new SemanticChangeSet(new[] { b, a });

        Assert.Equal(forward.ToCanonicalJson(), reverse.ToCanonicalJson());
        Assert.Equal(new[] { "chg-000001", "chg-000002" }, forward.Changes.Select(change => change.Id));
        Assert.Equal(forward.ToCanonicalUtf8Bytes(), reverse.ToCanonicalUtf8Bytes());
    }

    [Fact]
    public void Operation_schema_distinguishes_insert_delete_move_and_modify()
    {
        var baseDocument = IrTestDocuments.Create("Alpha one two three.", "Bravo one two three.");
        var insertedDocument = IrTestDocuments.Create(
            "Alpha one two three.", "Inserted one two three.", "Bravo one two three.");
        var inserted = SemanticDiff.Compare(baseDocument, insertedDocument,
            new SemanticDiffOptions { IncludePackageChanges = false });
        var deleted = SemanticDiff.Compare(insertedDocument, baseDocument,
            new SemanticDiffOptions { IncludePackageChanges = false });
        Assert.Contains(inserted.Changes, change => change.Operation == SemanticChangeOperation.Insert);
        Assert.Contains(deleted.Changes, change => change.Operation == SemanticChangeOperation.Delete);

        var formatLeft = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Formatting only</w:t></w:r></w:p>");
        var formatRight = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Formatting only</w:t></w:r></w:p>");
        var modified = SemanticDiff.Compare(formatLeft, formatRight,
            new SemanticDiffOptions { IncludePackageChanges = false });
        Assert.Contains(modified.Changes, change =>
            change.Operation == SemanticChangeOperation.Modify
            && change.Family == SemanticChangeFamily.RunFormatting);

        var moveLeft = IrTestDocuments.Create(
            "Opening paragraph remains in the document.",
            "This complete paragraph moves with several distinctive words.",
            "Closing paragraph remains in the document.");
        var moveRight = IrTestDocuments.Create(
            "This complete paragraph moves with several distinctive words.",
            "Opening paragraph remains in the document.",
            "Closing paragraph remains in the document.");
        var first = SemanticDiff.Compare(moveLeft, moveRight,
            new SemanticDiffOptions { IncludePackageChanges = false });
        var second = SemanticDiff.Compare(moveLeft, moveRight,
            new SemanticDiffOptions { IncludePackageChanges = false });
        var moved = Assert.Single(first.Changes.Where(change =>
            change.Operation == SemanticChangeOperation.Move));
        Assert.False(string.IsNullOrWhiteSpace(moved.MoveId));
        Assert.NotNull(moved.LeftAnchor);
        Assert.NotNull(moved.RightAnchor);
        Assert.Equal(first.ToCanonicalJson(), second.ToCanonicalJson());
        Assert.Equal(moved.MoveId, Assert.Single(second.Changes.Where(change =>
            change.Operation == SemanticChangeOperation.Move)).MoveId);
    }

    [Fact]
    public void Formatting_only_change_does_not_become_text_or_package_noise()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr><w:r><w:t>Same text</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:jc w:val=\"right\"/></w:pPr><w:r><w:t>Same text</w:t></w:r></w:p>");

        var result = SemanticDiff.Compare(left, right);

        Assert.Contains(result.Changes, change =>
            change.Family == SemanticChangeFamily.ParagraphFormatting
            && change.Operation == SemanticChangeOperation.Modify);
        Assert.DoesNotContain(result.Changes, change => change.Family == SemanticChangeFamily.Text);
        Assert.DoesNotContain(result.Changes, change =>
            change.Family == SemanticChangeFamily.OpaquePackagePart);
    }

    [Fact]
    public void Relationship_only_change_is_isolated_from_content_families()
    {
        var basis = IrTestDocuments.Create("Unchanged body");
        var left = WithExternalRelationship(basis, "https://example.test/old");
        var right = WithExternalRelationship(basis, "https://example.test/new");

        var result = SemanticDiff.Compare(left, right);

        Assert.NotEmpty(result.Changes);
        Assert.All(result.Changes, change => Assert.Equal(SemanticChangeFamily.Relationship, change.Family));
        Assert.Contains(result.Changes, change => change.Operation == SemanticChangeOperation.Modify);
    }

    [Fact]
    public void Relationship_binding_swap_is_visible_but_coordinated_rid_renumber_is_not()
    {
        var basis = IrTestDocuments.Create("Unchanged body");
        var left = WithCustomRelationshipGraph(
            basis,
            "rId1",
            ("rId1", "https://example.test/a"),
            ("rId2", "https://example.test/b"));
        var swapped = WithCustomRelationshipGraph(
            basis,
            "rId1",
            ("rId1", "https://example.test/b"),
            ("rId2", "https://example.test/a"));

        var swapChanges = SemanticDiff.Compare(left, swapped);

        Assert.Contains(swapChanges.Changes, change =>
            change.Family == SemanticChangeFamily.Relationship
            && change.Operation == SemanticChangeOperation.Modify
            && change.Path.StartsWith("relationship.binding", StringComparison.Ordinal));

        var renumbered = WithCustomRelationshipGraph(
            basis,
            "rId9",
            ("rId9", "https://example.test/a"),
            ("rId8", "https://example.test/b"));
        Assert.Empty(SemanticDiff.Compare(left, renumbered).Changes);
    }

    [Fact]
    public void Hyperlink_target_change_is_relationship_and_hyperlink_not_text()
    {
        const string leftId = "rIdLeft";
        const string rightId = "rIdRight";
        var left = IrTestDocuments.FromBodyXmlWithHyperlinks(
            $"<w:p><w:hyperlink r:id=\"{leftId}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>Same label</w:t></w:r></w:hyperlink></w:p>",
            (leftId, "https://example.test/a"));
        var right = IrTestDocuments.FromBodyXmlWithHyperlinks(
            $"<w:p><w:hyperlink r:id=\"{rightId}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>Same label</w:t></w:r></w:hyperlink></w:p>",
            (rightId, "https://example.test/b"));

        var result = SemanticDiff.Compare(left, right);

        AssertFamilies(result, SemanticChangeFamily.Hyperlink, SemanticChangeFamily.Relationship);
        Assert.DoesNotContain(result.Changes, change => change.Family == SemanticChangeFamily.Text);
    }

    [Fact]
    public void Bookmark_and_revision_location_changes_are_not_suppressed()
    {
        var bookmarkLeft = IrTestDocuments.FromBodyXml(
            "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"Clause\"/><w:r><w:t>Alpha</w:t></w:r>" +
            "<w:bookmarkEnd w:id=\"1\"/></w:p><w:p><w:r><w:t>Beta</w:t></w:r></w:p>");
        var bookmarkRight = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:bookmarkStart w:id=\"77\" w:name=\"Clause\"/>" +
            "<w:r><w:t>Beta</w:t></w:r><w:bookmarkEnd w:id=\"77\"/></w:p>");

        Assert.Contains(SemanticDiff.Compare(bookmarkLeft, bookmarkRight).Changes, change =>
            change.Family == SemanticChangeFamily.Bookmark
            && change.Operation == SemanticChangeOperation.Move
            && change.LeftAnchor is not null
            && change.RightAnchor is not null);

        var revisionLeft = IrTestDocuments.FromBodyXml(
            "<w:p><w:ins w:id=\"1\" w:author=\"A\"><w:r><w:t>Tracked</w:t></w:r></w:ins></w:p>" +
            "<w:p><w:r><w:t>Destination</w:t></w:r></w:p>");
        var revisionRight = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Destination</w:t></w:r></w:p>" +
            "<w:p><w:ins w:id=\"1\" w:author=\"A\"><w:r><w:t>Tracked</w:t></w:r></w:ins></w:p>");
        Assert.Contains(SemanticDiff.Compare(revisionLeft, revisionRight).Changes, change =>
            change.Family == SemanticChangeFamily.Revision
            && change.Operation == SemanticChangeOperation.Move);
    }

    [Fact]
    public void Registry_residuals_cover_doc_defaults_numbering_and_theme_details()
    {
        const string body = "<w:p><w:r><w:t>Same</w:t></w:r></w:p>";
        var styleLeft = IrTestDocuments.FromParts(body,
            "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>");
        var styleRight = IrTestDocuments.FromParts(body,
            "<w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault></w:docDefaults>");
        Assert.Contains(SemanticDiff.Compare(styleLeft, styleRight).Changes, change =>
            change.Family == SemanticChangeFamily.Style
            && change.Path == "style.registry.package");

        var numberingLeft = IrTestDocuments.FromParts(body, numberingInnerXml:
            Numbering("decimal", 1));
        var numberingRight = IrTestDocuments.FromParts(body, numberingInnerXml:
            Numbering("decimal", 1).Replace("</w:lvl>", "<w:suff w:val=\"space\"/></w:lvl>", StringComparison.Ordinal));
        Assert.Contains(SemanticDiff.Compare(numberingLeft, numberingRight).Changes, change =>
            change.Family == SemanticChangeFamily.Numbering
            && change.Path == "numbering.registry.package");

        const string latin = "<a:majorFont><a:latin typeface=\"Major\"/></a:majorFont>" +
            "<a:minorFont><a:latin typeface=\"Minor\"/></a:minorFont>";
        const string latinAndJapanese =
            "<a:majorFont><a:latin typeface=\"Major\"/>" +
            "<a:font script=\"Jpan\" typeface=\"Yu Mincho\"/></a:majorFont>" +
            "<a:minorFont><a:latin typeface=\"Minor\"/></a:minorFont>";
        var themeLeft = IrTestDocuments.FromParts(body, themeFontSchemeInnerXml: latin);
        var themeRight = IrTestDocuments.FromParts(body, themeFontSchemeInnerXml: latinAndJapanese);
        Assert.Contains(SemanticDiff.Compare(themeLeft, themeRight).Changes, change =>
            change.Family == SemanticChangeFamily.Style
            && change.Path == "theme.registry.package");
    }

    [Fact]
    public void Unknown_xml_comments_and_processing_instructions_are_semantic()
    {
        var basis = IrTestDocuments.Create("Same body");
        var left = WithCustomXmlPayload(basis,
            "<?audit before?><vendorData xmlns=\"urn:example:vendor\"><!--left--><value>A</value></vendorData>");
        var right = WithCustomXmlPayload(basis,
            "<?audit after?><vendorData xmlns=\"urn:example:vendor\"><!--right--><value>A</value></vendorData>");

        Assert.Contains(SemanticDiff.Compare(left, right).Changes, change =>
            change.Family == SemanticChangeFamily.OpaquePackagePart);
    }

    [Fact]
    public void Inserting_a_comment_does_not_cascade_modifications_to_existing_comments()
    {
        const string oldBody =
            "<w:p><w:r><w:t>First target</w:t></w:r></w:p>" +
            "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t>Existing target</w:t></w:r>" +
            "<w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>";
        var left = IrTestDocuments.WithComment(
            "Alice", "A", "2026-01-01T00:00:00Z", "Existing", oldBody);
        const string newBody =
            "<w:p><w:commentRangeStart w:id=\"1\"/><w:r><w:t>First target</w:t></w:r>" +
            "<w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r></w:p>" +
            "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t>Existing target</w:t></w:r>" +
            "<w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>";
        var right = RewriteEntry(left, "word/document.xml",
            $"<w:document xmlns:w=\"{W}\"><w:body>{newBody}</w:body></w:document>");
        right = RewriteEntry(right, "word/comments.xml",
            $"<w:comments xmlns:w=\"{W}\">" +
            "<w:comment w:id=\"1\" w:author=\"Bob\" w:initials=\"B\" w:date=\"2026-01-02T00:00:00Z\">" +
            "<w:p><w:r><w:t>Inserted</w:t></w:r></w:p></w:comment>" +
            "<w:comment w:id=\"0\" w:author=\"Alice\" w:initials=\"A\" w:date=\"2026-01-01T00:00:00Z\">" +
            "<w:p><w:r><w:t xml:space=\"preserve\">Existing</w:t></w:r></w:p></w:comment></w:comments>");

        var result = SemanticDiff.Compare(left, right);
        var commentChanges = result.Changes
            .Where(change => change.Family == SemanticChangeFamily.Comment)
            .ToArray();

        var inserted = Assert.Single(commentChanges);
        Assert.Equal(SemanticChangeOperation.Insert, inserted.Operation);
    }

    [Fact]
    public void Isolated_table_row_and_cell_formatting_emit_typed_families()
    {
        var rowLeft = IrTestDocuments.FromBodyXml(Table(
            "Same", "Grid", 1, 1000, includeSecondRow: false, rowHeight: 200));
        var rowRight = IrTestDocuments.FromBodyXml(Table(
            "Same", "Grid", 1, 1000, includeSecondRow: false, rowHeight: 400));
        Assert.Contains(SemanticDiff.Compare(rowLeft, rowRight).Changes, change =>
            change.Family == SemanticChangeFamily.TableRow);

        var cellLeft = IrTestDocuments.FromBodyXml(Table(
            "Same", "Grid", 1, 1000, includeSecondRow: false, rowHeight: 200));
        var cellRight = IrTestDocuments.FromBodyXml(Table(
            "Same", "Grid", 2, 2400, includeSecondRow: false, rowHeight: 200));
        var cellChanges = SemanticDiff.Compare(cellLeft, cellRight);
        AssertFamilies(cellChanges, SemanticChangeFamily.TableSpan, SemanticChangeFamily.TableWidth);
    }

    [Fact]
    public void Section_type_only_does_not_claim_page_setup_changed()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Same</w:t></w:r></w:p><w:sectPr><w:type w:val=\"continuous\"/></w:sectPr>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Same</w:t></w:r></w:p><w:sectPr><w:type w:val=\"nextPage\"/></w:sectPr>");

        var result = SemanticDiff.Compare(left, right);

        Assert.Contains(result.Changes, change => change.Family == SemanticChangeFamily.Section);
        Assert.DoesNotContain(result.Changes, change => change.Family == SemanticChangeFamily.PageSetup);
    }

    [Fact]
    public void Package_limits_are_enforced_before_ir_expansion()
    {
        var document = IrTestDocuments.Create("Same");

        Assert.Throws<InvalidDataException>(() => SemanticDiff.Compare(document, document,
            new SemanticDiffOptions
            {
                PackageOptions = new PackageManifestOptions
                {
                    MaxTotalUncompressedBytes = 1,
                },
            }));
        Assert.Throws<InvalidDataException>(() => SemanticDiff.Compare(document, document,
            new SemanticDiffOptions
            {
                PackageOptions = new PackageManifestOptions
                {
                    MaxUriLength = 4,
                },
            }));
        Assert.Throws<InvalidDataException>(() => SemanticDiff.Compare(document, document,
            new SemanticDiffOptions
            {
                PackageOptions = new PackageManifestOptions
                {
                    MaxCompressionRatio = 1,
                },
            }));

        var duplicate = WithDuplicateEntry(document, "word/settings.xml", "<settings/>");
        Assert.Throws<InvalidDataException>(() => new OpcSemanticPackageChangeDetector().Compare(
            duplicate,
            document.DocumentByteArray,
            new SemanticDiffOptions()));
    }

    [Fact]
    public void Public_and_shared_byte_facades_emit_the_same_canonical_schema()
    {
        var left = IrTestDocuments.Create("Before");
        var right = IrTestDocuments.Create("After");

        var publicJson = DocxDiff.GetSemanticChanges(left, right).ToCanonicalJson();
        var wireJson = DocxDiffOps.GetSemanticChangesJson(
            left.DocumentByteArray, right.DocumentByteArray, settingsJson: null);

        Assert.Equal(publicJson, wireJson);
    }

    [Fact]
    public void Published_v1_json_schema_tracks_the_wire_vocabulary()
    {
        var schemaPath = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "../../../../docs/schemas/semantic-changes-v1.schema.json"));
        using var schema = JsonDocument.Parse(File.ReadAllBytes(schemaPath));
        var root = schema.RootElement;

        Assert.Equal("https://json-schema.org/draft/2020-12/schema",
            root.GetProperty("$schema").GetString());
        var properties = root.GetProperty("properties");
        Assert.Equal(SemanticChangeSet.CurrentSchema,
            properties.GetProperty("schema").GetProperty("const").GetString());
        Assert.Equal(SemanticChangeSet.CurrentSchemaVersion,
            properties.GetProperty("schemaVersion").GetProperty("const").GetInt32());

        var changeProperties = root.GetProperty("$defs")
            .GetProperty("change")
            .GetProperty("properties");
        var schemaOperations = changeProperties.GetProperty("operation")
            .GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString())
            .ToArray();
        var schemaFamilies = changeProperties.GetProperty("family")
            .GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString())
            .ToArray();

        Assert.Equal(
            Enum.GetValues<SemanticChangeOperation>().Select(SemanticChangeSet.OperationName),
            schemaOperations);
        Assert.Equal(
            Enum.GetValues<SemanticChangeFamily>().Select(SemanticChangeSet.FamilyName),
            schemaFamilies);
        Assert.Equal(Enum.GetValues<SemanticValueKind>().Length,
            root.GetProperty("$defs").GetProperty("value").GetProperty("oneOf").GetArrayLength());
    }

    [Fact]
    public void Session_semantic_changes_compare_against_the_opening_package()
    {
        var document = IrTestDocuments.Create("Opening text", "Delete this paragraph");
        using var session = new DocxSession(document.DocumentByteArray);
        Assert.Empty(session.GetSemanticChanges().Changes);

        var target = session.Project().AnchorIndex.Values
            .First(anchor => anchor.Anchor.Scope == "body" && anchor.TextPreview.Contains("Delete", StringComparison.Ordinal));
        Assert.True(session.DeleteBlock(target.Anchor.Id).Success);

        var result = session.GetSemanticChanges();
        Assert.Contains(result.Changes, change =>
            change.Operation == SemanticChangeOperation.Delete
            && change.Family is SemanticChangeFamily.Text or SemanticChangeFamily.BlockStructure);

        using var noBaseline = new DocxSession(document.DocumentByteArray,
            new DocxSessionSettings { CaptureInitialProjection = false });
        Assert.Throws<InvalidOperationException>(() => noBaseline.GetSemanticChanges());
    }

    [Fact]
    public void Untouched_blank_session_suppresses_checkpoint_xml_reserialization()
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());

        Assert.Empty(session.GetSemanticChanges().Changes);
    }

    [Fact]
    [Trait("Category", "Performance")]
    public void Thousand_paragraph_single_edit_has_bounded_time_and_output()
    {
        var paragraphs = Enumerable.Range(0, 1_000).Select(index => $"Clause {index:D4}").ToArray();
        var changed = paragraphs.ToArray();
        changed[500] = "Clause 0500 revised";
        var left = IrTestDocuments.Create(paragraphs);
        var right = IrTestDocuments.Create(changed);
        var timer = Stopwatch.StartNew();

        var result = SemanticDiff.Compare(left, right);
        timer.Stop();
        var outputBytes = result.ToCanonicalUtf8Bytes().Length;
        _output.WriteLine(
            $"paragraphs=1000 elapsedMs={timer.Elapsed.TotalMilliseconds:F1} " +
            $"changes={result.ChangeCount} canonicalBytes={outputBytes}");

        Assert.Contains(result.Changes, change => change.Family == SemanticChangeFamily.Text);
        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(30),
            $"Semantic diff took {timer.Elapsed.TotalSeconds:F2}s; 30s regression bound.");
        Assert.True(outputBytes < 1_000_000,
            $"Semantic output was {outputBytes} bytes; 1MB regression bound.");
    }

    [Fact]
    [Trait("Category", "Performance")]
    public void Dense_table_document_single_edit_has_bounded_time_and_output()
    {
        var tables = Enumerable.Range(0, 200)
            .Select(index => Table(
                $"Dense table row {index:D3}",
                "Grid",
                1,
                1800,
                includeSecondRow: true,
                rowHeight: 240))
            .ToArray();
        var changed = tables.ToArray();
        changed[100] = Table(
            "Dense table row 100 revised",
            "Grid",
            2,
            2200,
            includeSecondRow: true,
            rowHeight: 360);
        var left = IrTestDocuments.FromBodyXml(string.Concat(tables));
        var right = IrTestDocuments.FromBodyXml(string.Concat(changed));
        var timer = Stopwatch.StartNew();

        var result = SemanticDiff.Compare(left, right);
        timer.Stop();
        int outputBytes = result.ToCanonicalUtf8Bytes().Length;
        _output.WriteLine(
            $"tables=200 rows=400 elapsedMs={timer.Elapsed.TotalMilliseconds:F1} " +
            $"changes={result.ChangeCount} canonicalBytes={outputBytes}");

        AssertFamilies(result,
            SemanticChangeFamily.Table,
            SemanticChangeFamily.TableRow,
            SemanticChangeFamily.TableSpan,
            SemanticChangeFamily.TableWidth);
        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(30),
            $"Dense semantic diff took {timer.Elapsed.TotalSeconds:F2}s; 30s regression bound.");
        Assert.True(outputBytes < 1_000_000,
            $"Dense semantic output was {outputBytes} bytes; 1MB regression bound.");
    }

    private static void AssertFamilies(SemanticChangeSet result, params SemanticChangeFamily[] families)
    {
        foreach (var family in families)
            Assert.Contains(result.Changes, change => change.Family == family);
    }

    private static void AssertCompleteLocationAndValues(SemanticChange change)
    {
        Assert.StartsWith("chg-", change.Id, StringComparison.Ordinal);
        Assert.StartsWith("/", change.PartUri, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(change.Path));
        Assert.NotNull(change.Before);
        Assert.NotNull(change.After);
        Assert.True(change.LeftScope is not null || change.RightScope is not null
            || change.LeftAnchor is null && change.RightAnchor is null);
    }

    private static IEnumerable<SemanticValue> DescendantValues(SemanticValue value)
    {
        yield return value;
        foreach (var property in value.Properties)
        foreach (var nested in DescendantValues(property.Value))
            yield return nested;
        foreach (var item in value.Items)
        foreach (var nested in DescendantValues(item))
            yield return nested;
    }

    private static string Numbering(string format, int numId) =>
        $"<w:abstractNum w:abstractNumId=\"{numId}\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/>" +
        $"<w:numFmt w:val=\"{format}\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>" +
        $"<w:num w:numId=\"{numId}\"><w:abstractNumId w:val=\"{numId}\"/></w:num>";

    private static string Table(
        string text,
        string style,
        int span,
        int width,
        bool includeSecondRow,
        int rowHeight) =>
        "<w:tbl><w:tblPr>" +
        $"<w:tblStyle w:val=\"{style}\"/><w:tblW w:w=\"{width}\" w:type=\"dxa\"/>" +
        "</w:tblPr><w:tblGrid><w:gridCol w:w=\"1200\"/><w:gridCol w:w=\"1200\"/></w:tblGrid>" +
        $"<w:tr><w:trPr><w:trHeight w:val=\"{rowHeight}\"/></w:trPr><w:tc><w:tcPr>" +
        $"<w:tcW w:w=\"{width}\" w:type=\"dxa\"/><w:gridSpan w:val=\"{span}\"/>" +
        $"</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc></w:tr>" +
        (includeSecondRow
            ? "<w:tr><w:tc><w:p><w:r><w:t>Added row</w:t></w:r></w:p></w:tc></w:tr>"
            : string.Empty) +
        "</w:tbl>";

    private static WmlDocument WithCustomXml(WmlDocument source, bool annotation)
        => WithCustomXmlPayload(source, annotation
            ? "<annotations xmlns=\"http://docxodus.dev/annotations/v1\" version=\"1.0\">" +
              "<annotation id=\"ann-1\" labelId=\"clause\" label=\"Clause\" color=\"#ffff00\">" +
              "<range bookmarkName=\"docxodus_ann_1\"/></annotation></annotations>"
            : "<vendorData xmlns=\"urn:example:vendor\"><value>changed</value></vendorData>");

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

    private static WmlDocument RewriteEntry(WmlDocument source, string entryName, string xml)
    {
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry(entryName)!;
            using var target = entry.Open();
            target.SetLength(0);
            var bytes = Encoding.UTF8.GetBytes(xml);
            target.Write(bytes);
        }
        return new WmlDocument(source.FileName, stream.ToArray());
    }

    private static WmlDocument RewriteCustomXmlContentType(WmlDocument source, string contentType)
    {
        XDocument types;
        string customPartName;
        using (var stream = new MemoryStream(source.DocumentByteArray, writable: false))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
        {
            var entry = archive.GetEntry("[Content_Types].xml")!;
            using var input = entry.Open();
            types = XDocument.Load(input);
            customPartName = "/" + archive.Entries.First(item =>
                item.FullName.StartsWith("customXml/item", StringComparison.Ordinal)
                && item.FullName.EndsWith(".xml", StringComparison.Ordinal)
                && !item.FullName.Contains("itemProps", StringComparison.Ordinal)).FullName;
        }

        var root = types.Root!;
        var declaration = root.Elements().FirstOrDefault(element =>
            element.Name.LocalName == "Override"
            && (string?)element.Attribute("PartName") == customPartName);
        if (declaration is null)
        {
            declaration = new XElement(root.Name.Namespace + "Override",
                new XAttribute("PartName", customPartName));
            root.Add(declaration);
        }
        declaration.SetAttributeValue("ContentType", contentType);
        return RewriteEntry(source, "[Content_Types].xml",
            types.ToString(SaveOptions.DisableFormatting));
    }

    private static WmlDocument WithExternalRelationship(WmlDocument source, string target)
    {
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            document.MainDocumentPart!.AddExternalRelationship(
                "urn:docxodus:test-relationship",
                new Uri(target),
                "rIdSemanticOnly");
        }
        return new WmlDocument(source.FileName, stream.ToArray());
    }

    private static WmlDocument WithCustomRelationshipGraph(
        WmlDocument source,
        string referencedId,
        params (string Id, string Target)[] relationships) =>
        WithCustomRelationshipGraph(
            source,
            referencedId,
            "External",
            "urn:docxodus:test-binding",
            relationships);

    private static WmlDocument WithCustomInternalRelationship(
        WmlDocument source,
        string target) => WithCustomRelationshipGraph(
            source,
            "rIdInternal",
            "Internal",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps",
            new[] { ("rIdInternal", target) });

    private static WmlDocument WithCustomRelationshipGraph(
        WmlDocument source,
        string referencedId,
        string targetMode,
        string relationshipType,
        IReadOnlyList<(string Id, string Target)> relationships)
    {
        const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var document = WithCustomXmlPayload(source,
            $"<vendorData xmlns=\"urn:example:vendor\" xmlns:r=\"{R}\" r:ref=\"{referencedId}\"/>");
        if (relationshipType.EndsWith("/customXmlProps", StringComparison.Ordinal))
        {
            using var package = new MemoryStream();
            package.Write(document.DocumentByteArray);
            using (var wordDocument = WordprocessingDocument.Open(package, true))
            {
                var ownerPart = wordDocument.MainDocumentPart!.CustomXmlParts.Single();
                var propertiesPart = ownerPart.AddNewPart<CustomXmlPropertiesPart>("rIdInternal");
                using var properties = propertiesPart.GetStream(FileMode.Create, FileAccess.Write);
                var payload = Encoding.UTF8.GetBytes(
                    "<ds:datastoreItem xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\" ds:itemID=\"{00000000-0000-0000-0000-000000000001}\"><ds:schemaRefs/></ds:datastoreItem>");
                properties.Write(payload);
            }
            document = new WmlDocument(source.FileName, package.ToArray());
        }
        using var stream = new MemoryStream();
        stream.Write(document.DocumentByteArray);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var owner = archive.Entries.First(entry =>
                entry.FullName.StartsWith("customXml/item", StringComparison.Ordinal)
                && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)
                && !entry.FullName.Contains("itemProps", StringComparison.Ordinal));
            var directory = Path.GetDirectoryName(owner.FullName)!.Replace('\\', '/');
            var relationshipName = $"{directory}/_rels/{Path.GetFileName(owner.FullName)}.rels";
            archive.GetEntry(relationshipName)?.Delete();
            var relationshipPart = archive.CreateEntry(relationshipName);
            using var target = relationshipPart.Open();
            var xml = new XElement(
                XName.Get("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships"),
                relationships.Select(relationship => new XElement(
                    XName.Get("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships"),
                    new XAttribute("Id", relationship.Id),
                    new XAttribute("Type", relationshipType),
                    new XAttribute("Target", relationship.Target),
                    new XAttribute("TargetMode", targetMode))));
            var bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
            target.Write(bytes);
        }
        return new WmlDocument(source.FileName, stream.ToArray());
    }

    private static byte[] WithDuplicateEntry(
        WmlDocument source,
        string entryName,
        string payload)
    {
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var duplicate = archive.CreateEntry(entryName);
            using var target = duplicate.Open();
            var bytes = Encoding.UTF8.GetBytes(payload);
            target.Write(bytes);
        }
        return stream.ToArray();
    }

    private static SemanticChange Change(
        string id,
        string partUri,
        SemanticChangeFamily family,
        string value) => new()
    {
        Id = id,
        Operation = SemanticChangeOperation.Modify,
        Family = family,
        PartUri = partUri,
        Path = "test",
        Before = SemanticValue.String(value + "-before"),
        After = SemanticValue.String(value + "-after"),
    };
}
