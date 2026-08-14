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

    [Fact]
    public void DS45511_TrackedSdtDelete_AcceptMatchesDirectWithoutParagraphHusks()
    {
        var baseline = BuildSdtDeleteBaseline();
        byte[] expected;
        using (var direct = new DocxSession(baseline))
        {
            Assert.True(direct.DeleteRange(AnchorByText(direct, "delete start"),
                AnchorByText(direct, "after")).Success);
            expected = direct.Save();
        }

        byte[] actual;
        using (var tracked = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "SDT Reviewer",
        }))
        {
            Assert.True(tracked.DeleteRange(AnchorByText(tracked, "delete start"),
                AnchorByText(tracked, "after")).Success);
            var resolved = tracked.AcceptAllRevisions();
            Assert.True(resolved.Success, resolved.Error?.Message);
            Assert.Empty(tracked.ListRevisions());
            actual = tracked.Save();
        }

        var expectedRoot = MainRoot(expected);
        var actualRoot = MainRoot(actual);
        Assert.True(XNode.DeepEquals(expectedRoot, actualRoot),
            FirstDifference(expectedRoot, actualRoot));
        Assert.Equal(new[] { "before", "after" },
            actualRoot.Descendants(W.p).Select(paragraph => paragraph.Value).ToArray());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS45512_TrackedInsertParagraph_RoundTrips(bool accept)
    {
        var baseline = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        byte[] expected;
        if (accept)
        {
            using var direct = new DocxSession(baseline);
            Assert.True(direct.InsertParagraph(FirstBodyTextAnchor(direct), Position.After,
                "inserted paragraph").Success);
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
            RevisionAuthor = "Paragraph Reviewer",
        }))
        {
            Assert.True(tracked.InsertParagraph(FirstBodyTextAnchor(tracked), Position.After,
                "inserted paragraph").Success);
            var revision = Assert.Single(tracked.ListRevisions());
            Assert.Equal(RevisionFamily.ContentInsert, revision.Family);
            var resolved = accept
                ? tracked.AcceptRevision(revision.Id)
                : tracked.RejectRevision(revision.Id);
            Assert.True(resolved.Success, resolved.Error?.Message);
            Assert.Empty(tracked.ListRevisions());
            actual = tracked.Save();
        }

        Assert.True(XNode.DeepEquals(MainRoot(expected), MainRoot(actual)),
            FirstDifference(MainRoot(expected), MainRoot(actual)));
    }

    [Theory]
    [InlineData("split")]
    [InlineData("merge")]
    [InlineData("insert_table")]
    public void DS45513_UnsafeTrackedStructuralOperations_FailTypedAndUndoFree(string operation)
    {
        var baseline = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
        });
        var paragraphs = session.Project().AnchorIndex.Values
            .Where(anchor => anchor.Anchor.Scope == "body" && anchor.Anchor.Kind == "p")
            .Select(anchor => anchor.Anchor.Id).Take(2).ToArray();

        var result = operation switch
        {
            "split" => session.SplitParagraph(paragraphs[0], 1),
            "merge" => session.MergeParagraphs(paragraphs[0], paragraphs[1]),
            "insert_table" => session.InsertTable(paragraphs[0], Position.After, 1, 1),
            _ => throw new ArgumentOutOfRangeException(nameof(operation)),
        };

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, result.Error!.Code);
        Assert.True(XNode.DeepEquals(MainRoot(baseline), MainRoot(session.Save())));
        Assert.False(session.Undo());
    }

    [Theory]
    [InlineData("paragraph_style", true)]
    [InlineData("paragraph_style", false)]
    [InlineData("paragraph_format", true)]
    [InlineData("paragraph_format", false)]
    [InlineData("column_widths", true)]
    [InlineData("column_widths", false)]
    [InlineData("table_borders", true)]
    [InlineData("table_borders", false)]
    [InlineData("cell_shading", true)]
    [InlineData("cell_shading", false)]
    [InlineData("row_options", true)]
    [InlineData("row_options", false)]
    public void DS45514_TrackedPropertyMutation_RoundTrips(string operation, bool accept)
    {
        var baseline = operation.StartsWith("paragraph", StringComparison.Ordinal)
            ? DocxSessionTests.BuildDS001_SimpleTwoParagraphs()
            : BuildTableDocument();
        byte[] expected;
        if (accept)
        {
            using var direct = new DocxSession(baseline);
            Assert.True(ApplyPropertyMutation(direct, operation).Success);
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
            RevisionAuthor = "Format Reviewer",
        }))
        {
            var edit = ApplyPropertyMutation(tracked, operation);
            Assert.True(edit.Success, edit.Error?.Message);
            var revision = Assert.Single(tracked.ListRevisions());
            Assert.Equal(RevisionFamily.PropertiesChange, revision.Family);
            var resolved = accept
                ? tracked.AcceptRevision(revision.Id)
                : tracked.RejectRevision(revision.Id);
            Assert.True(resolved.Success, resolved.Error?.Message);
            Assert.Empty(tracked.ListRevisions());
            actual = tracked.Save();
        }

        var expectedRoot = MainRoot(expected);
        var actualRoot = MainRoot(actual);
        Assert.True(XNode.DeepEquals(expectedRoot, actualRoot),
            FirstDifference(expectedRoot, actualRoot));
        Assert.Equal(ValidationErrors(expected), ValidationErrors(actual));
    }

    [Fact]
    public void DS45515_StableIdsAndSequentialRows_RemainIndependent()
    {
        using (var content = new DocxSession(BuildSeparatedInsertions()))
        {
            var before = content.ListRevisions();
            var first = Assert.Single(before, revision => revision.ConstituentIds.SequenceEqual(new[] { "1" }));
            var separator = Assert.Single(before, revision => revision.ConstituentIds.SequenceEqual(new[] { "2" }));
            var third = Assert.Single(before, revision => revision.ConstituentIds.SequenceEqual(new[] { "3" }));

            Assert.True(content.AcceptRevision(separator.Id).Success);

            var after = content.ListRevisions();
            Assert.Equal(first.Id, Assert.Single(after,
                revision => revision.ConstituentIds.SequenceEqual(new[] { "1" })).Id);
            Assert.Equal(third.Id, Assert.Single(after,
                revision => revision.ConstituentIds.SequenceEqual(new[] { "3" })).Id);
        }

        using var rows = new DocxSession(BuildTableDocument(), new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Row Reviewer",
        });
        var cell = FirstCellAnchor(rows);
        Assert.True(rows.InsertTableRow(cell, Position.After).Success);
        Assert.True(rows.InsertTableRow(cell, Position.After).Success);
        var rowRevisions = rows.ListRevisions()
            .Where(revision => revision.Family == RevisionFamily.RowInsert).ToList();
        Assert.Equal(2, rowRevisions.Count);
        var survivingId = rowRevisions[1].Id;
        Assert.True(rows.AcceptRevision(rowRevisions[0].Id).Success);
        Assert.Equal(survivingId, Assert.Single(rows.ListRevisions(),
            revision => revision.Family == RevisionFamily.RowInsert).Id);
    }

    [Theory]
    [InlineData("math_ctrlpr")]
    [InlineData("numbering_delete")]
    [InlineData("run_properties_delete")]
    public void DS45516_UnhandledRecognizedFamilies_AreListedAndBlockBulk(string shape)
    {
        var input = BuildUnsupportedRevisionDocument(shape);
        using var session = new DocxSession(input);
        var before = MainRoot(session.Save());
        var revision = Assert.Single(session.ListRevisions());
        Assert.Equal(RevisionFamily.Unsupported, revision.Family);
        Assert.Equal(RevisionResolutionStatus.Unsupported, revision.ResolutionStatus);
        Assert.Equal("unsupported_revision_family", revision.Diagnostic!.Code);

        var result = session.AcceptAllRevisions();

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.RevisionUnsupported, result.Error!.Code);
        Assert.True(XNode.DeepEquals(before, MainRoot(session.Save())));
        Assert.False(session.Undo());
    }

    [Fact]
    public void DS45517_NonnumericNativeId_IsMalformedAndFailsClosed()
    {
        var input = MutateMain(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), root =>
        {
            var run = root.Descendants(W.p).First().Elements(W.r).First();
            run.ReplaceWith(new XElement(W.ins,
                new XAttribute(W.id, "not-an-integer"),
                new XAttribute(W.author, "Bad Producer"),
                new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                new XElement(run)));
        });
        using var session = new DocxSession(input);
        var revision = Assert.Single(session.ListRevisions());
        Assert.Equal(RevisionResolutionStatus.Malformed, revision.ResolutionStatus);
        Assert.Equal("invalid_revision_id", revision.Diagnostic!.Code);

        var result = session.AcceptRevision(revision.Id);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.RevisionMalformed, result.Error!.Code);
        Assert.False(session.Undo());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DS45518_RejectTrackedListCreation_RestoresPackageAndUndoRedo(bool bulk)
    {
        var baseline = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        Assert.False(HasNumberingPart(baseline));
        using var session = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "List Reviewer",
        });
        Assert.True(session.ApplyListFormat(FirstBodyTextAnchor(session), ListFormat.Decimal).Success);
        Assert.True(HasNumberingPart(session.Save()));
        var revision = Assert.Single(session.ListRevisions());

        var rejected = bulk
            ? session.RejectAllRevisions()
            : session.RejectRevision(revision.Id);

        Assert.True(rejected.Success, rejected.Error?.Message);
        Assert.False(HasNumberingPart(session.Save()));
        Assert.True(XNode.DeepEquals(MainRoot(baseline), MainRoot(session.Save())));
        Assert.True(session.Undo());
        Assert.True(HasNumberingPart(session.Save()));
        Assert.Single(session.ListRevisions());
        Assert.True(session.Redo());
        Assert.False(HasNumberingPart(session.Save()));
        Assert.Empty(session.ListRevisions());
    }

    [Theory]
    [InlineData("level")]
    [InlineData("switch_format")]
    public void DS45519_TrackedListMutation_RequiringNumberingSidePartChangeFailsClosed(
        string operation)
    {
        var baseline = BuildSingleLevelBulletDocument();
        using var session = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
        });
        var anchor = FirstBodyTextAnchor(session);
        var operationBaseline = session.Save();
        var beforeNumbering = NumberingPartBytes(operationBaseline);

        var result = operation == "level"
            ? session.SetListLevel(anchor, 1)
            : session.ApplyListFormat(anchor, ListFormat.Decimal);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, result.Error!.Code);
        Assert.Empty(session.ListRevisions());
        var after = session.Save();
        Assert.True(beforeNumbering.SequenceEqual(NumberingPartBytes(after)));
        Assert.True(XNode.DeepEquals(MainRoot(operationBaseline), MainRoot(after)));
        PackageEquivalence.AssertSamePackage(
            new WmlDocument("baseline.docx", operationBaseline),
            new WmlDocument("after.docx", after));
        Assert.False(session.Undo());
    }

    [Fact]
    public void DS45520_TrackedListFormat_UsingExistingDefinitionRejectsWithPackageParity()
    {
        var baseline = BuildBulletAndDecimalDocument();
        using var session = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "List Reviewer",
        });
        var anchor = FirstBodyTextAnchor(session);
        var operationBaseline = session.Save();
        var beforeNumbering = NumberingPartBytes(operationBaseline);

        var edit = session.ApplyListFormat(anchor, ListFormat.Decimal);

        Assert.True(edit.Success, edit.Error?.Message);
        var revision = Assert.Single(session.ListRevisions());
        Assert.Equal(RevisionFamily.PropertiesChange, revision.Family);
        Assert.True(session.RejectRevision(revision.Id).Success);
        var after = session.Save();
        Assert.True(beforeNumbering.SequenceEqual(NumberingPartBytes(after)));
        Assert.True(XNode.DeepEquals(MainRoot(operationBaseline), MainRoot(after)));
        Assert.Equal(PackagePartUris(operationBaseline), PackagePartUris(after));
    }

    [Fact]
    public void DS45521_TrackedParagraphStyleRequiresExistingDefinitionAndStyleTopologyUndoes()
    {
        var baseline = BuildDocumentWithoutStylesPart();
        Assert.False(HasStylesPart(baseline));
        using (var tracked = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
        }))
        {
            var anchor = FirstBodyTextAnchor(tracked);
            var operationBaseline = tracked.Save();
            var result = tracked.SetParagraphStyle(anchor, "Heading2");

            Assert.False(result.Success);
            Assert.Equal(EditErrorCode.TrackedOperationUnsupported, result.Error!.Code);
            Assert.Empty(tracked.ListRevisions());
            var after = tracked.Save();
            Assert.False(HasStylesPart(after));
            Assert.True(XNode.DeepEquals(MainRoot(operationBaseline), MainRoot(after)));
            PackageEquivalence.AssertSamePackage(
                new WmlDocument("baseline.docx", operationBaseline),
                new WmlDocument("after.docx", after));
            Assert.False(tracked.Undo());
        }

        using var direct = new DocxSession(baseline);
        Assert.True(direct.SetParagraphStyle(
            FirstBodyTextAnchor(direct), "Heading2").Success);
        Assert.True(HasStylesPart(direct.Save()));
        Assert.True(direct.Undo());
        Assert.False(HasStylesPart(direct.Save()));
        Assert.True(direct.Redo());
        Assert.True(HasStylesPart(direct.Save()));
    }

    [Fact]
    public void DS45522_TrackedParagraphStyleRejectDoesNotChangeStylesPart()
    {
        var baseline = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
        });
        var anchor = FirstBodyTextAnchor(session);
        var operationBaseline = session.Save();
        var beforeStyles = StylesPartBytes(operationBaseline);

        Assert.True(session.SetParagraphStyle(anchor, "Heading2").Success);
        var revision = Assert.Single(session.ListRevisions());
        Assert.True(session.RejectRevision(revision.Id).Success);

        var after = session.Save();
        Assert.True(beforeStyles.SequenceEqual(StylesPartBytes(after)));
        Assert.True(XNode.DeepEquals(MainRoot(operationBaseline), MainRoot(after)));
        Assert.Equal(PackagePartUris(operationBaseline), PackagePartUris(after));
    }

    [Theory]
    [InlineData("del_text", true)]
    [InlineData("del_text", false)]
    [InlineData("del_instr_text", true)]
    [InlineData("del_instr_text", false)]
    public void DS45523_OrphanDeletedPayload_IsListedAndBlocksBulk(
        string shape, bool accept)
    {
        var input = BuildOrphanDeletedPayloadDocument(shape);
        using var session = new DocxSession(input);
        var before = MainRoot(session.Save());
        var revision = Assert.Single(session.ListRevisions());
        Assert.Equal(RevisionFamily.Unsupported, revision.Family);
        Assert.Equal(RevisionResolutionStatus.Unsupported, revision.ResolutionStatus);
        Assert.Equal("unsupported_revision_family", revision.Diagnostic!.Code);

        var result = accept
            ? session.AcceptAllRevisions()
            : session.RejectAllRevisions();

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.RevisionUnsupported, result.Error!.Code);
        Assert.True(XNode.DeepEquals(before, MainRoot(session.Save())));
        Assert.False(session.Undo());
    }

    [Theory]
    [InlineData("del_text")]
    [InlineData("del_instr_text")]
    public void DS45524_DeletedPayloadInsideClaimedWrapper_IsNotDoubleCounted(string shape)
    {
        using var session = new DocxSession(BuildOrdinaryDeletionDocument(shape));

        var revision = Assert.Single(session.ListRevisions());

        Assert.Equal(RevisionFamily.ContentDelete, revision.Family);
        Assert.Equal(RevisionResolutionStatus.Supported, revision.ResolutionStatus);
    }

    [Theory]
    [InlineData("outer_accept_inner_accept")]
    [InlineData("inner_accept_outer_accept")]
    [InlineData("outer_accept_inner_reject")]
    [InlineData("inner_reject_outer_accept")]
    public void DS45525_NestedIndependentRevisions_ResolveInEitherOrder(string order)
    {
        using var session = new DocxSession(BuildNestedRevisionDocument());
        var initial = session.ListRevisions();
        Assert.Equal(2, initial.Count);
        var outer = Assert.Single(initial, revision =>
            revision.ConstituentIds.SequenceEqual(new[] { "901" }));
        var inner = Assert.Single(initial, revision =>
            revision.ConstituentIds.SequenceEqual(new[] { "902" }));
        bool innerAccepted = order.Contains("inner_accept", StringComparison.Ordinal);

        if (order.StartsWith("outer", StringComparison.Ordinal))
        {
            Assert.True(session.AcceptRevision(outer.Id).Success);
            var remaining = Assert.Single(session.ListRevisions());
            Assert.Equal(inner.Id, remaining.Id);
            Assert.True((innerAccepted
                ? session.AcceptRevision(remaining.Id)
                : session.RejectRevision(remaining.Id)).Success);
        }
        else
        {
            Assert.True((innerAccepted
                ? session.AcceptRevision(inner.Id)
                : session.RejectRevision(inner.Id)).Success);
            var remaining = Assert.Single(session.ListRevisions());
            Assert.Equal(outer.Id, remaining.Id);
            Assert.True(session.AcceptRevision(remaining.Id).Success);
        }

        Assert.Empty(session.ListRevisions());
        Assert.Equal(innerAccepted ? "outer-before outer-after" : "outer-before inner outer-after",
            MainRoot(session.Save()).Descendants(W.p).First().Value);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS45526_NestedIndependentRevisions_BulkMatchesProcessor(bool accept)
    {
        var input = BuildNestedRevisionDocument();
        var expected = accept
            ? RevisionProcessor.AcceptRevisions(new WmlDocument("accept.docx", input)).DocumentByteArray
            : RevisionProcessor.RejectRevisions(new WmlDocument("reject.docx", input)).DocumentByteArray;
        using var session = new DocxSession(input);

        var result = accept
            ? session.AcceptAllRevisions()
            : session.RejectAllRevisions();

        Assert.True(result.Success, result.Error?.Message);
        Assert.Empty(session.ListRevisions());
        var actual = session.Save();
        Assert.True(XNode.DeepEquals(MainRoot(expected), MainRoot(actual)),
            FirstDifference(MainRoot(expected), MainRoot(actual)));
    }

    private static byte[] BuildSingleLevelBulletDocument()
    {
        byte[] listed;
        using (var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs()))
        {
            Assert.True(session.ApplyListFormat(
                FirstBodyTextAnchor(session), ListFormat.Bullet).Success);
            listed = session.Save();
        }

        using var stream = new MemoryStream();
        stream.Write(listed);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var numbering = document.MainDocumentPart!.NumberingDefinitionsPart!;
            var root = numbering.GetXDocument().Root!;
            var abstractNum = Assert.Single(root.Elements(W.abstractNum));
            foreach (var extra in abstractNum.Elements(W.lvl).Skip(1).ToList()) extra.Remove();
            var multiLevelType = abstractNum.Element(W.multiLevelType);
            if (multiLevelType is null)
                abstractNum.AddFirst(new XElement(W.multiLevelType, new XAttribute(W.val, "singleLevel")));
            else
                multiLevelType.SetAttributeValue(W.val, "singleLevel");
            numbering.PutXDocument();
        }
        return stream.ToArray();
    }

    private static byte[] BuildBulletAndDecimalDocument()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchors = session.Project().AnchorIndex.Values
            .Where(anchor => anchor.Anchor.Scope == "body" && anchor.Anchor.Kind == "p")
            .Select(anchor => anchor.Anchor.Id).Take(2).ToArray();
        Assert.True(session.ApplyListFormat(anchors[0], ListFormat.Bullet).Success);
        Assert.True(session.ApplyListFormat(anchors[1], ListFormat.Decimal).Success);
        return session.Save();
    }

    private static byte[] BuildDocumentWithoutStylesPart()
    {
        using var stream = new MemoryStream();
        stream.Write(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            if (main.StyleDefinitionsPart is { } styles) main.DeletePart(styles);
        }
        return stream.ToArray();
    }

    private static byte[] BuildOrphanDeletedPayloadDocument(string shape) =>
        MutateMain(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), root =>
        {
            var name = shape switch
            {
                "del_text" => W.delText,
                "del_instr_text" => W.delInstrText,
                _ => throw new ArgumentOutOfRangeException(nameof(shape)),
            };
            root.Descendants(W.p).First().ReplaceNodes(
                new XElement(W.r, new XElement(name, "orphan deleted payload")));
        });

    private static byte[] BuildOrdinaryDeletionDocument(string shape) =>
        MutateMain(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), root =>
        {
            var name = shape switch
            {
                "del_text" => W.delText,
                "del_instr_text" => W.delInstrText,
                _ => throw new ArgumentOutOfRangeException(nameof(shape)),
            };
            root.Descendants(W.p).First().ReplaceNodes(
                new XElement(W.del,
                    new XAttribute(W.id, "900"),
                    new XAttribute(W.author, "Reviewer"),
                    new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                    new XElement(W.r, new XElement(name, "ordinary deletion"))));
        });

    private static byte[] BuildNestedRevisionDocument() =>
        MutateMain(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), root =>
        {
            root.Descendants(W.p).First().ReplaceNodes(
                new XElement(W.ins,
                    new XAttribute(W.id, "901"),
                    new XAttribute(W.author, "Outer Reviewer"),
                    new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                    new XElement(W.r, new XElement(W.t, "outer-before ")),
                    new XElement(W.del,
                        new XAttribute(W.id, "902"),
                        new XAttribute(W.author, "Inner Reviewer"),
                        new XAttribute(W.date, "2026-01-02T00:00:00Z"),
                        new XElement(W.r, new XElement(W.delText, "inner "))),
                    new XElement(W.r, new XElement(W.t, "outer-after"))));
        });

    private static byte[] NumberingPartBytes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        using var partStream = document.MainDocumentPart!.NumberingDefinitionsPart!.GetStream();
        using var copy = new MemoryStream();
        partStream.CopyTo(copy);
        return copy.ToArray();
    }

    private static bool HasStylesPart(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.StyleDefinitionsPart is not null;
    }

    private static byte[] StylesPartBytes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        using var partStream = document.MainDocumentPart!.StyleDefinitionsPart!.GetStream();
        using var copy = new MemoryStream();
        partStream.CopyTo(copy);
        return copy.ToArray();
    }

    private static string[] PackagePartUris(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.GetPackage().GetParts()
            .Select(part => part.Uri.ToString())
            .OrderBy(uri => uri, StringComparer.Ordinal)
            .ToArray();
    }

    private static EditResult ApplyPropertyMutation(DocxSession session, string operation)
    {
        if (operation.StartsWith("paragraph", StringComparison.Ordinal))
        {
            var paragraph = FirstBodyTextAnchor(session);
            return operation switch
            {
                "paragraph_style" => session.SetParagraphStyle(paragraph, "Heading2"),
                "paragraph_format" => session.SetParagraphFormat(paragraph,
                    new ParagraphFormatOp { Alignment = ParagraphAlignment.Center, SpacingAfter = 240 }),
                _ => throw new ArgumentOutOfRangeException(nameof(operation)),
            };
        }

        var cell = FirstCellAnchor(session);
        return operation switch
        {
            "column_widths" => session.SetColumnWidths(cell, new[] { 2600, 3000 }),
            "table_borders" => session.SetTableBorders(cell,
                new TableBorderSpec { Style = "double", Size = 8, Color = "CC0000" }),
            "cell_shading" => session.SetCellShading(cell, "D9EAF7", TableShadingScope.Row),
            "row_options" => session.SetTableRowOptions(cell,
                new TableRowOptions { RepeatHeader = true, AllowBreakAcrossPages = false, HeightTwips = 480 }),
            _ => throw new ArgumentOutOfRangeException(nameof(operation)),
        };
    }

    private static byte[] BuildSdtDeleteBaseline() =>
        MutateMain(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), root =>
        {
            var body = root.Element(W.body)!;
            var sectPr = body.Element(W.sectPr) is { } section
                ? new XElement(section)
                : null;
            body.ReplaceNodes(
                Paragraph("before"),
                Paragraph("delete start"),
                new XElement(W.sdt,
                    new XElement(W.sdtPr,
                        new XElement(W.tag, new XAttribute(W.val, "controlled"))),
                    new XElement(W.sdtContent, Paragraph("controlled paragraph"))),
                Paragraph("after"),
                sectPr);
        });

    private static byte[] BuildSeparatedInsertions() =>
        MutateMain(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), root =>
        {
            var paragraph = root.Descendants(W.p).First();
            paragraph.ReplaceNodes(
                RevisionWrapper(W.ins, "1", "A", "2026-01-01T00:00:00Z", "one"),
                RevisionWrapper(W.del, "2", "B", "2026-01-01T00:00:00Z", "separator"),
                RevisionWrapper(W.ins, "3", "A", "2026-01-02T00:00:00Z", "three"));
        });

    private static byte[] BuildUnsupportedRevisionDocument(string shape) =>
        MutateMain(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), root =>
        {
            var paragraph = root.Descendants(W.p).First();
            var marker = new XElement(W.del,
                new XAttribute(W.id, "44"),
                new XAttribute(W.author, "Unsupported Reviewer"),
                new XAttribute(W.date, "2026-01-01T00:00:00Z"));
            switch (shape)
            {
                case "math_ctrlpr":
                    paragraph.ReplaceNodes(new XElement(M.oMath,
                        new XElement(M.f,
                            new XElement(M.fPr, new XElement(M.ctrlPr, marker)),
                            new XElement(M.num, new XElement(W.r, new XElement(W.t, "1"))),
                            new XElement(M.den, new XElement(W.r, new XElement(W.t, "2"))))));
                    break;
                case "numbering_delete":
                    paragraph.AddFirst(new XElement(W.pPr,
                        new XElement(W.numPr,
                            new XElement(W.ilvl, new XAttribute(W.val, 0)),
                            new XElement(W.numId, new XAttribute(W.val, 1)),
                            marker)));
                    break;
                case "run_properties_delete":
                    paragraph.Elements(W.r).First().AddFirst(new XElement(W.rPr, marker));
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(shape));
            }
        });

    private static XElement Paragraph(string text) =>
        new(W.p, new XElement(W.r, new XElement(W.t, text)));

    private static XElement RevisionWrapper(
        XName name, string id, string author, string date, string text) =>
        new(name,
            new XAttribute(W.id, id),
            new XAttribute(W.author, author),
            new XAttribute(W.date, date),
            new XElement(W.r,
                new XElement(name == W.del ? W.delText : W.t, text)));

    private static string AnchorByText(DocxSession session, string text) =>
        session.Project().AnchorIndex.Values.Single(anchor =>
            string.Equals(session.GetAnchorInfo(anchor.Anchor.Id)?.TextPreview,
                text, StringComparison.Ordinal)).Anchor.Id;

    private static bool HasNumberingPart(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.NumberingDefinitionsPart is not null;
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
