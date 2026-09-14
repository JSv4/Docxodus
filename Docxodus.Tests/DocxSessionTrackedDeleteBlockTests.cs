// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace Docxodus.Tests;

public class DocxSessionTrackedDeleteBlockTests
{
    public static IEnumerable<object[]> RoundTripCases =>
        from kind in new[] { "p", "h", "li", "empty", "tbl", "nested-table" }
        from resolver in new[] { "session", "diff", "individual" }
        select new object[] { kind, resolver };

    [Theory]
    [MemberData(nameof(RoundTripCases))]
    public void DeleteBlock_AcceptsToCleanDeletionAndRejectsToOriginal(string kind, string resolver)
    {
        var original = BuildDocument(Paragraph("before"), TargetBlock(kind), NumberedParagraph("after"));
        byte[] clean;
        using (var session = new DocxSession(original))
        {
            Assert.True(session.DeleteBlock(TargetAnchor(session)).Success);
            clean = session.Save();
        }

        byte[] tracked;
        using (var session = new DocxSession(original))
        {
            // Match the public repro: recording and author can be switched after opening.
            session.SetTrackedChanges(TrackedChangeMode.RenderInline);
            session.SetRevisionAuthor("Reviewer");
            var anchor = TargetAnchor(session);
            Assert.Equal(kind is "empty" ? "p" : kind is "nested-table" ? "tbl" : kind,
                session.GetAnchorInfo(anchor)!.Kind);
            var anchorsBefore = session.Project().AnchorIndex.Keys.ToHashSet();

            var result = session.DeleteBlock(anchor);

            Assert.True(result.Success, result.Error?.Message);
            Assert.Equal(anchor, Assert.Single(result.Modified).Id);
            Assert.Empty(result.Removed);
            Assert.Empty(result.Created);
            Assert.True(anchorsBefore.SetEquals(session.Project().AnchorIndex.Keys));
            Assert.Equal(1, session.UndoCount);
            tracked = session.Save();
        }

        AssertSchemaValid(tracked);
        var target = Body(tracked).Elements().ElementAt(1);
        Assert.All(target.DescendantsAndSelf(W.p), paragraph =>
            Assert.NotNull(paragraph.Element(W.pPr)?.Element(W.rPr)?.Element(W.del)));
        Assert.All(target.Descendants(W.tr), row =>
            Assert.NotNull(row.Element(W.trPr)?.Element(W.del)));
        Assert.NotEmpty(target.Descendants(W.del));
        Assert.All(target.Descendants(W.del), deletion =>
        {
            Assert.Equal("Reviewer", (string?)deletion.Attribute(W.author));
            Assert.NotNull(deletion.Attribute(W.id));
            Assert.NotNull(deletion.Attribute(W.date));
        });

        var accepted = Resolve(tracked, resolver, accept: true);
        var rejected = Resolve(tracked, resolver, accept: false);
        // Whole-body XML includes paragraph/table counts and numbering, which a text-only
        // comparison misses when acceptance strands an empty numbered paragraph (#784).
        AssertSameBody(clean, accepted);
        AssertSameBody(original, rejected);
        AssertSchemaValid(accepted);
        AssertSchemaValid(rejected);
    }

    [Theory]
    [InlineData("p")]
    [InlineData("tbl")]
    public void DeleteBlock_LastSibling_TracksOnlyTheTargetAndSupportsUndoRedo(string kind)
    {
        var original = BuildDocument(Paragraph("before"), TargetBlock(kind));
        using var session = OpenTracked(original);
        Assert.True(session.DeleteBlock(TargetAnchor(session)).Success);
        Assert.NotEmpty(session.ListRevisions());
        var tracked = session.Save();

        Assert.True(session.Undo());
        AssertSameBody(original, session.Save());
        Assert.Empty(session.ListRevisions());
        Assert.True(session.Redo());
        AssertSameBody(tracked, session.Save());
        Assert.True(session.AcceptAllRevisions().Success);
        Assert.Equal("before", Assert.Single(Body(session.Save()).Elements(W.p)).Value);
        Assert.Empty(Body(session.Save()).Descendants(W.tbl));
    }

    [Theory]
    [InlineData("p")]
    [InlineData("tbl")]
    public void DeleteBlock_PreservesReferencedBookmarksOrRefusesBeforeRecording(string kind)
    {
        var target = TargetBlock(kind);
        var paragraph = target is Paragraph p ? p : target.Descendants<Paragraph>().First();
        paragraph.PrependChild(new BookmarkStart { Id = "1", Name = "TableBookmark" });
        paragraph.AppendChild(new BookmarkEnd { Id = "1" });
        var original = BuildDocument(
            Paragraph("before"), target,
            new Paragraph(new Hyperlink(new Run(new Text("see table"))) { Anchor = "TableBookmark" }));

        using (var untracked = new DocxSession(original))
        {
            var refused = untracked.DeleteBlock(TargetAnchor(untracked));
            Assert.Equal(EditErrorCode.BookmarkInUse, refused.Error?.Code);
            Assert.Equal(0, untracked.UndoCount);
        }

        using var session = OpenTracked(original);
        var result = session.DeleteBlock(TargetAnchor(session));
        if (kind == "tbl")
        {
            Assert.Equal(EditErrorCode.BookmarkInUse, result.Error?.Code);
            Assert.Equal(0, session.UndoCount);
            Assert.Empty(session.ListRevisions());
        }
        else Assert.True(result.Success, result.Error?.Message);
        Assert.Empty(result.Removed);
        Assert.Equal("TableBookmark", Assert.Single(session.ListBookmarks()).Name);
        Assert.Equal(kind == "tbl" ? 1 : 0, Body(session.Save()).Descendants(W.tbl).Count());
        Assert.True(session.RejectAllRevisions().Success);
        AssertSameBody(original, session.Save());
        AssertSchemaValid(session.Save());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DeleteBlock_InlineCustomXml_RefusesBeforeMutation(bool inTable)
    {
        var paragraph = new Paragraph(new CustomXmlRun(
            new CustomXmlProperties(), new Run(new Text("unsupported")))
        {
            Uri = "urn:docxodus:test",
            Element = "inline",
        });
        OpenXmlElement target = paragraph;
        if (inTable)
        {
            var table = Table();
            var first = table.Descendants<Paragraph>().First();
            first.Parent!.ReplaceChild(paragraph, first);
            target = table;
        }
        using var session = OpenTracked(BuildDocument(Paragraph("before"), target, Paragraph("after")));
        var anchor = TargetAnchor(session);
        var before = session.Save();

        var result = session.DeleteBlock(anchor);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.IncompatibleElementType, result.Error?.Code);
        Assert.Contains("run-level w:customXml", result.Error?.Message);
        Assert.Equal(0, session.UndoCount);
        Assert.Empty(session.ListRevisions());
        AssertSameBody(before, session.Save());
    }

    private static DocxSession OpenTracked(byte[] bytes) => new(bytes, new DocxSessionSettings
    {
        TrackedChanges = TrackedChangeMode.RenderInline,
        RevisionAuthor = "Reviewer",
    });

    private static string TargetAnchor(DocxSession session) => session.Project().AnchorIndex.Values
        .Where(target => target.Anchor.Kind is "p" or "h" or "li" or "tbl")
        .Skip(1).First().Anchor.Id;

    private static byte[] Resolve(byte[] tracked, string resolver, bool accept)
    {
        if (resolver == "diff")
            return accept ? DocxDiffOps.AcceptRevisions(tracked) : DocxDiffOps.RejectRevisions(tracked);

        using var session = new DocxSession(tracked);
        Assert.NotEmpty(session.ListRevisions());
        if (resolver == "session")
        {
            var result = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
            Assert.True(result.Success, result.Error?.Message);
        }
        else
        {
            for (var revisions = session.ListRevisions(); revisions.Count > 0; revisions = session.ListRevisions())
            {
                var result = accept
                    ? session.AcceptRevision(revisions[0].Id)
                    : session.RejectRevision(revisions[0].Id);
                Assert.True(result.Success, result.Error?.Message);
                Assert.True(session.ListRevisions().Count < revisions.Count);
            }
        }
        Assert.Empty(session.ListRevisions());
        return session.Save();
    }

    private static OpenXmlElement TargetBlock(string kind) => kind switch
    {
        "p" => Paragraph("delete me"),
        "h" => NumberedParagraph("delete me", heading: true),
        "li" => NumberedParagraph("delete me"),
        "empty" => new Paragraph(),
        "tbl" => Table(),
        "nested-table" => Table(nested: true),
        _ => throw new ArgumentOutOfRangeException(nameof(kind)),
    };

    private static Paragraph Paragraph(string text) =>
        new(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    private static Paragraph NumberedParagraph(string text, bool heading = false)
    {
        var paragraph = Paragraph(text);
        var properties = new ParagraphProperties();
        if (heading)
            properties.Append(new ParagraphStyleId { Val = "Heading2" });
        properties.Append(
            new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 }),
            new ParagraphMarkRunProperties(new FontSize { Val = "28" }));
        paragraph.PrependChild(properties);
        return paragraph;
    }

    private static WTable Table(bool nested = false)
    {
        var table = new WTable(
            new TableProperties(), new TableGrid(new GridColumn { Width = "2500" }),
            new WTableRow(new WTableCell(Paragraph("Cell A"))),
            new WTableRow(new WTableCell(Paragraph("Cell B"))));
        if (nested)
        {
            var cell = table.Descendants<WTableCell>().First();
            cell.Append(Table(), new Paragraph());
        }
        return table;
    }

    private static byte[] BuildDocument(params OpenXmlElement[] blocks)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(blocks));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new DocDefaults(),
                new Style(new StyleName { Val = "heading 2" })
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Heading2",
                });
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                new AbstractNum(new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = 0 },
                new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });
            document.Save();
        }
        return stream.ToArray();
    }

    private static XElement Body(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return new XElement(document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!);
    }

    private static void AssertSameBody(byte[] expected, byte[] actual)
    {
        static XElement Normalize(XElement body)
        {
            body.DescendantsAndSelf().Attributes()
                .Where(a => a.IsNamespaceDeclaration || a.Name == PtOpenXml.Unid).Remove();
            // Rejecting a mark can leave empty property containers; retain all actual
            // properties and block nodes, especially empty or numbered paragraphs.
            foreach (var element in body.Descendants().Reverse().ToList())
                if ((element.Name == W.pPr || element.Name == W.rPr || element.Name == W.trPr)
                    && !element.HasElements && !element.HasAttributes)
                    element.Remove();
            return body;
        }

        Assert.Equal(Normalize(Body(expected)).ToString(), Normalize(Body(actual)).ToString());
    }

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(error => error.Description)));
    }
}
