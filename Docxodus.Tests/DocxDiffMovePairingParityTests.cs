// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

/// <summary>
/// Every move the comparison reports must be whole on both surfaces, and the two surfaces must agree
/// (issue #887): in the redline each <c>w:name</c> names exactly one moved-from range and one moved-to range,
/// in the revision list each Moved group has exactly one source and one destination, and the two count the
/// same moves. The shapes cover relocations across a table boundary, row reorders and moved tables, including
/// the cases where one surface draws a table whole and the paragraphs a redline fuses into one stream.
/// </summary>
public class DocxDiffMovePairingParityTests
{
    private const string A = "Alpha paragraph sets out the parties and the effective date of this agreement.";
    private const string B = "Bravo paragraph describes the services to be performed by the contractor.";
    private const string C = "Charlie paragraph covers payment terms, invoices and late fees in detail.";
    private const string D = "Delta paragraph addresses confidentiality obligations of both parties.";
    private const string E = "Echo paragraph limits liability and excludes consequential damages.";
    private const string F = "Foxtrot paragraph governs termination for convenience and for cause.";
    private const string G = "Golf paragraph handles notices sent by courier, mail and electronic means.";
    private const string H = "Hotel paragraph names the governing law and the exclusive venue for disputes.";
    private const string X = "Xray brand new unrelated sentence about warehouse inventory counts today.";
    private const string Y = "Yankee another unrelated sentence regarding parking spaces and badges.";

    private static string P(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    /// <summary>A paragraph whose text is a pending insertion by another author.</summary>
    private static string PendingInsertion(string text) =>
        $"<w:p><w:ins w:id=\"90\" w:author=\"Zed\" w:date=\"2020-01-01T00:00:00Z\"><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:ins></w:p>";

    private static string Cell(params string[] paragraphs) =>
        "<w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"dxa\"/></w:tcPr>" +
        string.Concat(paragraphs.Select(t => t.StartsWith('<') ? t : P(t))) + "</w:tc>";

    private static string Row(params string[] cells) => $"<w:tr>{string.Concat(cells)}</w:tr>";

    private static string Table(int columns, params string[] rows) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid>" +
        string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"2500\"/>", columns)) + "</w:tblGrid>" +
        string.Concat(rows) + "</w:tbl>";

    private static string OneCell(params string[] paragraphs) => Table(1, Row(Cell(paragraphs)));

    private static string OneColumn(params string[] rows) => Table(1, rows.Select(t => Row(Cell(t))).ToArray());

    public static TheoryData<string, string, string> Shapes => new()
    {
        { "paragraph into a cell", P(A) + P(B) + OneCell(C) + P(D), P(B) + OneCell(C, A) + P(D) },
        { "paragraph out of a cell", P(B) + OneCell(C, A) + P(D), P(B) + OneCell(C) + P(D) + P(A) },
        { "paragraph between two tables", OneCell(C, A) + P(B) + OneCell(D), OneCell(C) + P(B) + OneCell(D, A) },
        { "paragraph between cells of a row", Table(2, Row(Cell(A, B), Cell(C))) + P(D), Table(2, Row(Cell(B), Cell(C, A))) + P(D) },
        { "into a cell, replacing its other paragraph", P(A) + P(B) + OneCell(C, X) + P(D), P(B) + OneCell(C, A) + P(D) },
        { "out of the body end, beside an insertion", P(B) + OneCell(C) + P(D) + P(A), P(B) + OneCell(C, A) + P(D) + P(Y) },
        { "relocated and edited", P(A) + P(B) + OneCell(C) + P(D), P(B) + OneCell(C, A.Replace("parties", "signatories")) + P(D) },
        { "row reorder", OneColumn(A, B, C) + P(D), OneColumn(B, C, A) + P(D) },
        { "two row moves", OneColumn(A, B, C, D, E) + P(F), OneColumn(B, C, A, E, D) + P(F) },
        { "moved table", OneColumn(A, B) + P(C) + P(D), P(C) + P(D) + OneColumn(A, B) },
        { "moved and edited table", OneColumn(A, B) + P(C) + P(D) + P(E), P(C) + P(D) + P(E) + OneColumn(A, B.Replace("services", "consulting services")) },
        { "into a moved and edited table",
            OneColumn(C, D, E, F, G, H) + P(A) + P(B) + P(X) + P(Y),
            P(B) + P(X) + P(Y) + Table(1, Row(Cell(C)), Row(Cell(D)), Row(Cell(E)), Row(Cell(F)), Row(Cell(G)), Row(Cell(H, A))) },
        { "out of a moved and edited table",
            Table(1, Row(Cell(C)), Row(Cell(D)), Row(Cell(E)), Row(Cell(F)), Row(Cell(G)), Row(Cell(H, A))) + P(B) + P(X) + P(Y),
            P(B) + P(X) + P(Y) + P(A) + OneColumn(C, D, E, F, G, H) },
        { "into a row that gains a cell", P(A) + P(B) + Table(2, Row(Cell(C), Cell(D))) + P(E), P(B) + Table(3, Row(Cell(C, A), Cell(D), Cell(X))) + P(E) },
        { "into a table carrying a pending insertion",
            P(A) + P(B) + Table(1, Row(Cell(C)), Row(Cell(D))) + P(E),
            P(B) + Table(1, Row(Cell(C, A)), Row(Cell(D, PendingInsertion(X)))) + P(E) },
        { "beside paragraphs the redline fuses",
            P(G) + P(A) + P(B) + OneCell(C) + P(D),
            P(G.Replace("courier", "fax")) + P(B.Replace("services", "work")) + OneCell(C, A) + P(D) },
    };

    public static IEnumerable<object[]> ShapesUnderSettings()
    {
        var settings = new (string Name, Func<DocxDiffSettings> Make)[]
        {
            ("default", () => new DocxDiffSettings()),
            ("moves not reported", () => new DocxDiffSettings { DetectMoves = false }),
            ("cell insertions not tracked", () => new DocxDiffSettings { TrackCellInsertionsAndDeletions = false }),
            ("input revisions preserved", () => new DocxDiffSettings { PreserveInputRevisions = true }),
            ("no cross-paragraph runs", () => new DocxDiffSettings { CrossParagraphTokenDiff = false }),
            ("WmlComparer-compatible revisions", () => new DocxDiffSettings { RevisionGranularity = DocxDiffRevisionGranularity.WmlComparerCompatible }),
        };
        foreach (var row in Shapes)
            foreach (var (name, make) in settings)
                yield return new object[] { (string)row[0], name, (string)row[1], (string)row[2], make };
    }

    [Theory]
    [MemberData(nameof(ShapesUnderSettings))]
    public void EveryReportedMove_IsWholeAndAgreesAcrossSurfaces(
        string shape, string settingsName, string originalBody, string revisedBody, Func<DocxDiffSettings> settings)
    {
        _ = (shape, settingsName);
        var (left, right) = (IrTestDocuments.FromParts(originalBody), IrTestDocuments.FromParts(revisedBody));
        var applied = DocxCompare.ApplyFrontDoorRevisionPolicy(settings());

        var redline = DocxDiff.Compare(left, right, applied);
        var revisions = DocxDiff.GetRevisions(left, right, applied);

        var body = Body(redline.DocumentByteArray);
        var names = body.Descendants()
            .Where(e => e.Name == W.moveFromRangeStart || e.Name == W.moveToRangeStart)
            .GroupBy(e => (string?)e.Attribute(W.name))
            .ToList();
        Assert.All(names, group =>
        {
            Assert.Single(group, e => e.Name == W.moveFromRangeStart);
            Assert.Single(group, e => e.Name == W.moveToRangeStart);
        });
        var groups = revisions.Where(r => r.Type == DocxDiffRevisionType.Moved).GroupBy(r => r.MoveGroupId).ToList();
        Assert.All(groups, group =>
        {
            Assert.Single(group, r => r.IsMoveSource == true);
            Assert.Single(group, r => r.IsMoveSource == false);
        });
        Assert.Equal(names.Count, groups.Count);

        Assert.Equal(Text(DocxDiffOps.AcceptRevisions(right.DocumentByteArray)), Text(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray)));
        Assert.Equal(Text(DocxDiffOps.AcceptRevisions(left.DocumentByteArray)), Text(DocxDiffOps.RejectRevisions(redline.DocumentByteArray)));
        NoNewValidationErrors(left.DocumentByteArray, redline.DocumentByteArray);
    }

    [Fact]
    public void ParagraphBesideFusedParagraphs_IsNotPairedOnEitherSurface()
    {
        // The redline draws G, A and B as one cross-paragraph stream, where A has no standalone op to draw as
        // a move half; the revision list therefore leaves it unpaired too, rather than claiming a move the
        // redline does not show.
        var left = IrTestDocuments.FromParts(P(G) + P(A) + P(B) + OneCell(C) + P(D));
        var right = IrTestDocuments.FromParts(P(G.Replace("courier", "fax")) + P(B.Replace("services", "work")) + OneCell(C, A) + P(D));

        var redline = DocxCompare.Compare(left, right);
        var revisions = DocxDiff.GetRevisions(left, right, DocxCompare.ApplyFrontDoorRevisionPolicy(null));

        Assert.Empty(Body(redline.DocumentByteArray).Descendants(W.moveFromRangeStart));
        Assert.DoesNotContain(revisions, r => r.Type == DocxDiffRevisionType.Moved);
    }

    [Fact]
    public void MovedAndEditedTable_RevisionListDescribesTheEdit()
    {
        var left = IrTestDocuments.FromParts(OneColumn(A, B) + P(C) + P(D) + P(E));
        var right = IrTestDocuments.FromParts(P(C) + P(D) + P(E) + OneColumn(A, B.Replace("services", "consulting services")));

        var revisions = DocxDiff.GetRevisions(left, right, DocxCompare.ApplyFrontDoorRevisionPolicy(null));

        Assert.Equal(2, revisions.Count(r => r.Type == DocxDiffRevisionType.Moved));
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Inserted && r.Text.Contains("consulting"));
    }

    /// <summary>The data script's record of the paragraphs a fused build draws as runs must be exactly the
    /// paragraphs the fused build does draw as runs, or the two scripts would pair different relocations.</summary>
    [Fact]
    public void RecordedFusedParagraphs_MatchTheFusedBuild()
    {
        var corpus = Directory.GetFiles(Path.Combine(TestFileRoot(), "WC"), "*.docx")
            .GroupBy(f => System.Text.RegularExpressions.Regex.Match(Path.GetFileName(f), @"^(WC\d+|WC-[A-Za-z]+)").Value)
            .Where(g => g.Key.Length > 0)
            .SelectMany(g => g.OrderBy(f => f, StringComparer.Ordinal).Skip(1)
                .Select(other => (new WmlDocument(g.OrderBy(f => f, StringComparer.Ordinal).First()), new WmlDocument(other))))
            .Concat(Shapes.Select(row => (IrTestDocuments.FromParts((string)row[1]), IrTestDocuments.FromParts((string)row[2]))))
            .ToList();
        Assert.NotEmpty(corpus);
        var markup = new DocxDiffSettings().ToIrDiffSettings();
        var data = markup with { CrossParagraphTokenDiff = false };
        int fusedPairs = 0;
        foreach (var (left, right) in corpus)
        {
            var leftIr = IrReader.Read(DocxDiff.PreAccept(DocxCompare.ApplyFrontDoorRevisionPolicy(null), left), DocxDiff.ReadOpts);
            var rightIr = IrReader.Read(DocxDiff.PreAccept(DocxCompare.ApplyFrontDoorRevisionPolicy(null), right), DocxDiff.ReadOpts);
            var recorded = new HashSet<string>(StringComparer.Ordinal);
            IrEditScriptBuilder.Build(leftIr, rightIr, data, recorded);
            var fused = IrEditScriptBuilder.Build(leftIr, rightIr, markup).Operations
                .Where(op => op.Kind == IrEditOpKind.CrossParagraphRunBlock)
                .SelectMany(op => op.CrossParagraphCells!)
                .SelectMany(c => new[] { c.LeftAnchor, c.RightAnchor })
                .OfType<string>()
                .ToHashSet(StringComparer.Ordinal);
            Assert.True(fused.SetEquals(recorded), $"{left.FileName} vs {right.FileName}");
            if (fused.Count > 0)
                fusedPairs++;
        }
        Assert.True(fusedPairs > 0, "no pair exercised a fused run");
    }

    private static string TestFileRoot()
    {
        var dir = AppContext.BaseDirectory;
        while (dir != null && !Directory.Exists(Path.Combine(dir, "TestFiles")))
            dir = Path.GetDirectoryName(dir);
        return Path.Combine(dir!, "TestFiles");
    }

    private static XElement Body(byte[] bytes)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    private static string Text(byte[] bytes) => string.Join("|", Body(bytes).Descendants(W.p)
        .Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value))));
}
