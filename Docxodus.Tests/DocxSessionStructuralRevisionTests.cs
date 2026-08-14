// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionStructuralRevisionTests
{
    private static string Fixture(string relative) =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../TestFiles", relative));

    [Theory]
    [InlineData("RP/RP034-Deleted-Cells.docx", RevisionFamily.CellDelete)]
    [InlineData("RP/RP035-Inserted-Cells.docx", RevisionFamily.CellInsert)]
    [InlineData("RP/RP036-Vert-Merged-Cells.docx", RevisionFamily.CellMerge)]
    [InlineData("RP/RP016-Deleted-CC.docx", RevisionFamily.ContentControlDelete)]
    [InlineData("RP/RP017-Inserted-CC.docx", RevisionFamily.ContentControlInsert)]
    [InlineData("RP/RP021-Inserted-Numbering-Properties.docx", RevisionFamily.NumberingPropertiesInsert)]
    [InlineData("RP/RP026-NumberingChange.docx", RevisionFamily.NumberingChange)]
    public void DS45501_RealFixture_ListsNewFamilyAsSupported(string relative, RevisionFamily family)
    {
        using var session = new DocxSession(File.ReadAllBytes(Fixture(relative)));
        var matching = session.ListRevisions().Where(r => r.Family == family).ToList();

        Assert.NotEmpty(matching);
        Assert.All(matching, revision =>
        {
            Assert.StartsWith("rev2-", revision.Id);
            Assert.Equal("/word/document.xml", revision.PartUri);
            Assert.Equal("body", revision.Scope);
            Assert.NotEmpty(revision.ConstituentIds);
            Assert.NotEmpty(revision.AffectedAnchors);
            Assert.Equal(RevisionResolutionStatus.Supported, revision.ResolutionStatus);
            Assert.Null(revision.Diagnostic);
        });
    }

    [Theory]
    [InlineData("RP/RP034-Deleted-Cells.docx", true)]
    [InlineData("RP/RP034-Deleted-Cells.docx", false)]
    [InlineData("RP/RP035-Inserted-Cells.docx", true)]
    [InlineData("RP/RP035-Inserted-Cells.docx", false)]
    [InlineData("RP/RP036-Vert-Merged-Cells.docx", true)]
    [InlineData("RP/RP036-Vert-Merged-Cells.docx", false)]
    [InlineData("RP/RP016-Deleted-CC.docx", true)]
    [InlineData("RP/RP016-Deleted-CC.docx", false)]
    [InlineData("RP/RP017-Inserted-CC.docx", true)]
    [InlineData("RP/RP017-Inserted-CC.docx", false)]
    [InlineData("RP/RP021-Inserted-Numbering-Properties.docx", true)]
    [InlineData("RP/RP021-Inserted-Numbering-Properties.docx", false)]
    [InlineData("RP/RP026-NumberingChange.docx", true)]
    [InlineData("RP/RP026-NumberingChange.docx", false)]
    public void DS45502_IndividualAndBulkResolution_MatchProcessorOracle(string relative, bool accept)
    {
        var input = File.ReadAllBytes(Fixture(relative));
        var oracle = accept
            ? RevisionProcessor.AcceptRevisions(new WmlDocument("oracle.docx", input)).DocumentByteArray
            : RevisionProcessor.RejectRevisions(new WmlDocument("oracle.docx", input)).DocumentByteArray;

        byte[] individual;
        using (var session = new DocxSession(input))
        {
            for (int guard = 0; guard < 1000 && session.ListRevisions().Count > 0; guard++)
            {
                var revision = session.ListRevisions()[0];
                var result = accept
                    ? session.AcceptRevision(revision.Id)
                    : session.RejectRevision(revision.Id);
                Assert.True(result.Success, result.Error?.Message);
            }
            Assert.Empty(session.ListRevisions());
            individual = session.Save();
        }

        byte[] bulk;
        using (var session = new DocxSession(input))
        {
            var result = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
            Assert.True(result.Success, result.Error?.Message);
            Assert.Empty(session.ListRevisions());
            bulk = session.Save();

            Assert.True(session.Undo());
            Assert.NotEmpty(session.ListRevisions());
            Assert.True(session.Redo());
            Assert.Empty(session.ListRevisions());
        }

        var oracleRoot = MainRoot(oracle);
        var individualRoot = MainRoot(individual);
        var bulkRoot = MainRoot(bulk);
        Assert.True(XNode.DeepEquals(oracleRoot, individualRoot),
            FirstDifference(oracleRoot, individualRoot));
        Assert.True(XNode.DeepEquals(oracleRoot, bulkRoot),
            FirstDifference(oracleRoot, bulkRoot));

        // A few of the Office-produced fixtures contain extension attributes that the
        // SDK validator does not recognize. Resolution must not introduce any new
        // validation failures beyond those already present in the processor oracle.
        Assert.Equal(ValidationErrors(oracle), ValidationErrors(bulk));
    }

    [Theory]
    [InlineData("insert_row", true, RevisionFamily.RowInsert)]
    [InlineData("insert_row", false, RevisionFamily.RowInsert)]
    [InlineData("delete_row", true, RevisionFamily.RowDelete)]
    [InlineData("delete_row", false, RevisionFamily.RowDelete)]
    [InlineData("insert_column", true, RevisionFamily.CellInsert)]
    [InlineData("insert_column", false, RevisionFamily.CellInsert)]
    public void DS45503_TrackedTableMutation_ResolvesToDirectOrOriginal(
        string operation, bool accept, RevisionFamily family)
    {
        var baseline = BuildTableDocument();
        byte[] expected;
        if (accept)
        {
            using var direct = new DocxSession(baseline);
            var directResult = ApplyTableMutation(direct, operation);
            Assert.True(directResult.Success, directResult.Error?.Message);
            expected = direct.Save();
        }
        else
        {
            expected = baseline;
        }

        byte[] actual;
        using (var tracked = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Structural Reviewer",
        }))
        {
            var edit = ApplyTableMutation(tracked, operation);
            Assert.True(edit.Success, edit.Error?.Message);

            var revision = Assert.Single(tracked.ListRevisions());
            Assert.Equal(family, revision.Family);
            Assert.Equal("Structural Reviewer", revision.Author);
            Assert.Equal(RevisionResolutionStatus.Supported, revision.ResolutionStatus);
            Assert.True(revision.AffectedAnchors.Count >= 2);

            var resolution = accept
                ? tracked.AcceptRevision(revision.Id)
                : tracked.RejectRevision(revision.Id);
            Assert.True(resolution.Success, resolution.Error?.Message);
            Assert.Empty(tracked.ListRevisions());
            actual = tracked.Save();
        }

        Assert.True(XNode.DeepEquals(MainRoot(expected), MainRoot(actual)),
            FirstDifference(MainRoot(expected), MainRoot(actual)));
        Assert.Equal(ValidationErrors(expected), ValidationErrors(actual));
    }

    [Fact]
    public void DS45504_StableId_SurvivesRelistingSaveAndUnrelatedResolution()
    {
        var input = File.ReadAllBytes(Fixture("RP/RP034-Deleted-Cells.docx"));
        using var session = new DocxSession(input);
        var before = session.ListRevisions();
        var cell = Assert.Single(before, r => r.Family == RevisionFamily.CellDelete);
        Assert.Equal(cell.Id, Assert.Single(session.ListRevisions(), r =>
            r.Family == RevisionFamily.CellDelete).Id);

        var unrelated = before.FirstOrDefault(r => r.Id != cell.Id);
        if (unrelated is not null)
            Assert.True(session.AcceptRevision(unrelated.Id).Success);
        Assert.Equal(cell.Id, Assert.Single(session.ListRevisions(), r =>
            r.Family == RevisionFamily.CellDelete).Id);

        using var reopened = new DocxSession(session.Save());
        Assert.Equal(cell.Id, Assert.Single(reopened.ListRevisions(), r =>
            r.Family == RevisionFamily.CellDelete).Id);
    }

    [Theory]
    [InlineData("missing_id", RevisionResolutionStatus.Malformed, EditErrorCode.RevisionMalformed)]
    [InlineData("duplicate_id", RevisionResolutionStatus.Ambiguous, EditErrorCode.RevisionAmbiguous)]
    [InlineData("unsupported_move", RevisionResolutionStatus.Unsupported, EditErrorCode.RevisionUnsupported)]
    public void DS45505_InvalidTopology_IsListedAndFailsClosed(
        string shape, RevisionResolutionStatus status, EditErrorCode errorCode)
    {
        var input = BuildInvalidRevisionDocument(shape);
        using var session = new DocxSession(input);
        var before = MainRoot(session.Save());
        var invalid = session.ListRevisions().Where(r => r.ResolutionStatus == status).ToList();
        Assert.NotEmpty(invalid);
        if (shape == "duplicate_id")
            Assert.Equal(invalid.Count, invalid.Select(r => r.Id).Distinct().Count());
        var revision = invalid[0];
        Assert.NotNull(revision.Diagnostic);

        var result = session.AcceptRevision(revision.Id);
        Assert.False(result.Success);
        Assert.Equal(errorCode, result.Error!.Code);
        Assert.True(XNode.DeepEquals(before, MainRoot(session.Save())));
        Assert.False(session.Undo());
    }

    [Fact]
    public void DS45506_BulkResolution_BlockedEntryIsAtomic()
    {
        var input = BuildInvalidRevisionDocument("unsupported_move");
        using var session = new DocxSession(input);
        var before = MainRoot(session.Save());

        var result = session.AcceptAllRevisions();

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.RevisionUnsupported, result.Error!.Code);
        Assert.True(XNode.DeepEquals(before, MainRoot(session.Save())));
        Assert.False(session.Undo());
    }

    [Theory]
    [InlineData("apply", true, RevisionFamily.NumberingPropertiesInsert)]
    [InlineData("apply", false, RevisionFamily.NumberingPropertiesInsert)]
    [InlineData("remove", true, RevisionFamily.PropertiesChange)]
    [InlineData("remove", false, RevisionFamily.PropertiesChange)]
    [InlineData("level", true, RevisionFamily.PropertiesChange)]
    [InlineData("level", false, RevisionFamily.PropertiesChange)]
    public void DS45507_TrackedListMutation_UsesNativeRevisionAndRoundTrips(
        string operation, bool accept, RevisionFamily family)
    {
        var baseline = BuildListBaseline(operation);
        byte[] expected;
        if (accept)
        {
            using var direct = new DocxSession(baseline);
            Assert.True(ApplyListMutation(direct, operation).Success);
            expected = direct.Save();
        }
        else
        {
            expected = baseline;
        }

        byte[] actual;
        using (var tracked = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "List Reviewer",
        }))
        {
            var edit = ApplyListMutation(tracked, operation);
            Assert.True(edit.Success, edit.Error?.Message);
            var revision = Assert.Single(tracked.ListRevisions());
            Assert.Equal(family, revision.Family);
            Assert.Equal("List Reviewer", revision.Author);
            var resolution = accept
                ? tracked.AcceptRevision(revision.Id)
                : tracked.RejectRevision(revision.Id);
            Assert.True(resolution.Success, resolution.Error?.Message);
            Assert.Empty(tracked.ListRevisions());
            actual = tracked.Save();
        }

        Assert.True(XNode.DeepEquals(MainRoot(expected), MainRoot(actual)),
            FirstDifference(MainRoot(expected), MainRoot(actual)));
    }

    [Fact]
    public void DS45508_UnsupportedTrackedStructuralOperations_DoNotMutateOrCreateHistory()
    {
        var baseline = BuildTableDocument();
        using var session = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
        });
        var before = MainRoot(session.Save());
        var cell = FirstCellAnchor(session);

        var deleteColumn = session.DeleteTableColumn(cell);
        Assert.False(deleteColumn.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, deleteColumn.Error!.Code);

        var merge = session.MergeCells(cell, 1, 2);
        Assert.False(merge.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, merge.Error!.Code);

        var paragraphs = session.Project().AnchorIndex.Values
            .Where(a => a.Anchor.Scope == "body" && a.Anchor.Kind == "p"
                && (a.TextPreview.StartsWith("First", StringComparison.Ordinal)
                    || a.TextPreview.StartsWith("Second", StringComparison.Ordinal)))
            .Select(a => a.Anchor.Id).Take(2).ToArray();
        var listRange = session.ApplyListFormatRange(
            paragraphs[0], paragraphs[1], ListFormat.Decimal);
        Assert.False(listRange.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, listRange.Error!.Code);

        Assert.True(XNode.DeepEquals(before, MainRoot(session.Save())));
        Assert.False(session.Undo());
    }

    [Fact]
    public void DS45509_UnresolvedCellStructure_BlocksFurtherTableMutationWithoutHistory()
    {
        using var session = new DocxSession(
            File.ReadAllBytes(Fixture("RP/RP035-Inserted-Cells.docx")));
        var before = MainRoot(session.Save());

        var result = session.InsertTableRow(FirstCellAnchor(session), Position.After);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnresolvedStructuralRevision, result.Error!.Code);
        Assert.True(XNode.DeepEquals(before, MainRoot(session.Save())));
        Assert.False(session.Undo());
    }

    [Fact]
    public void DS45510_IdenticalNativeIdsInDifferentParts_AreIndependent()
    {
        using var session = new DocxSession(BuildCrossPartDuplicateIdDocument());
        var revisions = session.ListRevisions();
        Assert.Equal(2, revisions.Count);
        Assert.All(revisions, revision =>
        {
            Assert.Equal(new[] { "777" }, revision.ConstituentIds);
            Assert.Equal(RevisionResolutionStatus.Supported, revision.ResolutionStatus);
        });
        Assert.Equal(2, revisions.Select(revision => revision.Id).Distinct().Count());
        Assert.Contains(revisions, revision => revision.PartUri == "/word/document.xml"
            && revision.Scope == "body");
        var header = Assert.Single(revisions, revision => revision.PartUri.StartsWith(
            "/word/header", StringComparison.Ordinal));
        Assert.Equal("hdr1", header.Scope);

        Assert.True(session.AcceptRevision(header.Id).Success);

        var remaining = Assert.Single(session.ListRevisions());
        Assert.Equal("/word/document.xml", remaining.PartUri);
        Assert.Equal("777", Assert.Single(remaining.ConstituentIds));
    }

    private static byte[] BuildTableDocument()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraph = session.Project().AnchorIndex.Values
            .First(a => a.Anchor.Scope == "body" && a.Anchor.Kind == "p").Anchor.Id;
        var result = session.InsertTable(paragraph, Position.After, 2, 2,
            new TableInsertOptions
            {
                CellContents = new[] { "A1", "B1", "A2", "B2" },
                ColumnWidths = new[] { 2400, 3200 },
            });
        Assert.True(result.Success, result.Error?.Message);
        return session.Save();
    }

    private static string FirstCellAnchor(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .First(a => a.Anchor.Scope == "body" && a.Anchor.Kind == "tc").Anchor.Id;

    private static EditResult ApplyTableMutation(DocxSession session, string operation)
    {
        var cell = FirstCellAnchor(session);
        return operation switch
        {
            "insert_row" => session.InsertTableRow(cell, Position.After),
            "delete_row" => session.DeleteTableRow(cell),
            "insert_column" => session.InsertTableColumn(cell, Position.After),
            _ => throw new ArgumentOutOfRangeException(nameof(operation)),
        };
    }

    private static byte[] BuildListBaseline(string operation)
    {
        var input = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        if (operation == "apply") return input;
        using var session = new DocxSession(input);
        var paragraph = FirstBodyTextAnchor(session);
        var result = session.ApplyListFormat(paragraph, ListFormat.Decimal);
        Assert.True(result.Success, result.Error?.Message);
        return session.Save();
    }

    private static EditResult ApplyListMutation(DocxSession session, string operation)
    {
        var anchor = FirstBodyTextAnchor(session);
        return operation switch
        {
            "apply" => session.ApplyListFormat(anchor, ListFormat.Decimal),
            "remove" => session.RemoveListMembership(anchor),
            "level" => session.SetListLevel(anchor, 1),
            _ => throw new ArgumentOutOfRangeException(nameof(operation)),
        };
    }

    private static string FirstBodyTextAnchor(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .First(a => a.Anchor.Scope == "body" && a.Anchor.Kind is "p" or "li"
                && a.TextPreview.StartsWith("First", StringComparison.Ordinal)).Anchor.Id;

    private static byte[] BuildInvalidRevisionDocument(string shape)
    {
        if (shape == "missing_id")
            return MutateMain(File.ReadAllBytes(Fixture("RP/RP035-Inserted-Cells.docx")), root =>
                root.Descendants(W.cellIns).First().Attribute(W.id)?.Remove());

        return MutateMain(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), root =>
        {
            var paragraphs = root.Descendants(W.p).Take(2).ToArray();
            if (shape == "duplicate_id")
            {
                foreach (var paragraph in paragraphs)
                {
                    var run = paragraph.Elements(W.r).First();
                    run.ReplaceWith(new XElement(W.ins,
                        new XAttribute(W.id, "777"),
                        new XAttribute(W.author, "Duplicate"),
                        new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                        new XElement(run)));
                }
                return;
            }

            if (shape == "unsupported_move")
            {
                paragraphs[0].AddFirst(new XElement(W.customXmlMoveFromRangeStart,
                    new XAttribute(W.id, "888"),
                    new XAttribute(W.author, "Mover"),
                    new XAttribute(W.date, "2026-01-01T00:00:00Z")));
                paragraphs[0].Add(new XElement(W.customXmlMoveFromRangeEnd,
                    new XAttribute(W.id, "888")));
                return;
            }

            throw new ArgumentOutOfRangeException(nameof(shape));
        });
    }

    private static byte[] BuildCrossPartDuplicateIdDocument()
    {
        var input = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var stream = new MemoryStream();
        stream.Write(input);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var header = main.AddNewPart<HeaderPart>();
            header.PutXDocument(new XDocument(new XElement(W.hdr,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XElement(W.p,
                    new XElement(W.ins,
                        new XAttribute(W.id, "777"),
                        new XAttribute(W.author, "Header Reviewer"),
                        new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                        new XElement(W.r, new XElement(W.t, "Header revision")))))));

            var mainDocument = main.GetXDocument();
            var firstRun = mainDocument.Root!.Descendants(W.p).First().Elements(W.r).First();
            firstRun.ReplaceWith(new XElement(W.ins,
                new XAttribute(W.id, "777"),
                new XAttribute(W.author, "Body Reviewer"),
                new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                new XElement(firstRun)));
            var body = mainDocument.Root.Element(W.body)!;
            var sectPr = body.Elements(W.sectPr).LastOrDefault();
            if (sectPr is null)
            {
                sectPr = new XElement(W.sectPr);
                body.Add(sectPr);
            }
            XNamespace relationships =
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            sectPr.AddFirst(new XElement(W.headerReference,
                new XAttribute(relationships + "id", main.GetIdOfPart(header)),
                new XAttribute(W.type, "default")));
            main.PutXDocument();
        }
        return stream.ToArray();
    }

    private static byte[] MutateMain(byte[] input, Action<XElement> mutate)
    {
        using var stream = new MemoryStream();
        stream.Write(input);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var xDocument = document.MainDocumentPart!.GetXDocument();
            mutate(xDocument.Root!);
            document.MainDocumentPart.PutXDocument();
        }
        return stream.ToArray();
    }

    private static XElement MainRoot(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var root = new XElement(document.MainDocumentPart!.GetXDocument().Root!);

        // The legacy all-revisions processor removes Word's transient GoBack bookmark
        // while selective resolution intentionally preserves unrelated bookmarks.
        // Exclude that application affordance from semantic parity comparisons.
        var goBackIds = root.Descendants(W.bookmarkStart)
            .Where(e => (string?)e.Attribute(W.name) == "_GoBack")
            .Select(e => (string?)e.Attribute(W.id))
            .Where(id => id != null)
            .ToHashSet(StringComparer.Ordinal);
        foreach (var bookmark in root.Descendants(W.bookmarkStart)
            .Where(e => (string?)e.Attribute(W.name) == "_GoBack")
            .Concat(root.Descendants(W.bookmarkEnd)
                .Where(e => goBackIds.Contains((string?)e.Attribute(W.id)))).ToList())
            bookmark.Remove();
        foreach (var attribute in root.DescendantsAndSelf().Attributes()
            .Where(a => a.Name.Namespace == W.w
                && a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)).ToList())
            attribute.Remove();
        foreach (var whitespace in root.DescendantNodes().OfType<XText>()
            .Where(t => string.IsNullOrWhiteSpace(t.Value)).ToList())
            whitespace.Remove();
        foreach (var husk in root.Descendants()
            .Where(e => (e.Name == W.pPr || e.Name == W.rPr || e.Name == W.trPr)
                && !e.HasElements && !e.HasAttributes).ToList())
            husk.Remove();
        return root;
    }

    private static string[] ValidationErrors(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return new OpenXmlValidator().Validate(document)
            .Select(e => $"{e.Id}|{e.Description}|{e.Path?.XPath}")
            .OrderBy(e => e, StringComparer.Ordinal)
            .ToArray();
    }

    private static string FirstDifference(XElement expected, XElement actual)
    {
        var expectedNodes = expected.DescendantsAndSelf().ToList();
        var actualNodes = actual.DescendantsAndSelf().ToList();
        int count = Math.Min(expectedNodes.Count, actualNodes.Count);
        for (int i = 0; i < count; i++)
        {
            var left = expectedNodes[i];
            var right = actualNodes[i];
            if (left.Name != right.Name || left.Value != right.Value
                || !left.Attributes().OrderBy(a => a.Name.ToString())
                    .Select(a => (a.Name, a.Value))
                    .SequenceEqual(right.Attributes().OrderBy(a => a.Name.ToString())
                        .Select(a => (a.Name, a.Value))))
                return $"first difference at element {i}: expected {left}, actual {right}";
        }
        return $"element counts differ: expected {expectedNodes.Count}, actual {actualNodes.Count}";
    }
}
