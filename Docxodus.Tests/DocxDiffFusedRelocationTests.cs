// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Relocated text next to paragraphs the redline fuses into one cross-paragraph stream is reported as a move
/// on both surfaces (issue #930). The redline draws a fused stretch of edited paragraphs as one flat word
/// stream, so a relocation half inside it has no standalone paragraph op; it pairs when the stream holds the
/// half as one deleted (or inserted) run in one output paragraph, and both the redline and the revision list
/// then show the same move. A half the stream draws any other way — retained across a pilcrow, or split over
/// two output paragraphs — stays a plain deletion and insertion on both surfaces.
/// </summary>
public class DocxDiffFusedRelocationTests
{
    private const string A = "Alpha paragraph sets out the parties and the effective date of this agreement.";
    private const string B = "Bravo paragraph describes the services to be performed by the contractor.";
    private const string C = "Charlie paragraph covers payment terms, invoices and late fees in detail.";
    private const string D = "Delta paragraph addresses confidentiality obligations of both parties.";
    private const string G = "Golf paragraph handles notices sent by courier, mail and electronic means.";
    private const string M = "The moved sentence travels to the closing clause intact.";

    private static readonly string EditedG = G.Replace("courier", "fax");
    private static readonly string EditedB = B.Replace("services", "work");

    private static string P(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string OneCell(params string[] paragraphs) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"2500\"/></w:tblGrid>" +
        "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"dxa\"/></w:tcPr>" +
        string.Concat(paragraphs.Select(P)) + "</w:tc></w:tr></w:tbl>";

    /// <summary>G, A and B fuse into one stream (G and B edited, A deleted from between them).</summary>
    public static TheoryData<string, string, string, string> FusedMoves => new()
    {
        { "a sentence out of a fused paragraph",
            P($"{G} {M}") + P(A) + P(B) + OneCell(C) + P(D),
            P(EditedG) + P(EditedB) + OneCell(C) + P($"{D} {M}"), M },
        { "a sentence into a fused paragraph",
            P(G) + P(A) + P(B) + OneCell(C) + P($"{D} {M}"),
            P($"{EditedG} {M}") + P(EditedB) + OneCell(C) + P(D), M },
        { "a sentence between two fused paragraphs",
            P($"{M} {G}") + P(A) + P(B) + P(D),
            P(EditedG) + P($"{EditedB} {M}") + P(D), M },
        { "a fused paragraph into a cell",
            P(G) + P(A) + P(B) + OneCell(C) + P(D),
            P(EditedG) + P(EditedB) + OneCell(C, A) + P(D), A },
    };

    [Theory]
    [MemberData(nameof(FusedMoves))]
    public void MoveNextToFusedParagraphs_IsReportedOnBothSurfaces(
        string shape, string originalBody, string revisedBody, string moved)
    {
        _ = shape;
        var left = IrTestDocuments.FromParts(originalBody);
        var right = IrTestDocuments.FromParts(revisedBody);
        var settings = DocxCompare.ApplyFrontDoorRevisionPolicy(null);

        var redline = DocxDiff.Compare(left, right, settings);
        var revisions = DocxDiff.GetRevisions(left, right, settings);

        var body = Body(redline.DocumentByteArray);
        var from = Assert.Single(body.Descendants(W.moveFromRangeStart));
        var to = Assert.Single(body.Descendants(W.moveToRangeStart));
        Assert.Equal((string?)from.Attribute(W.name), (string?)to.Attribute(W.name));
        Assert.Equal(moved, MovedText(body, W.moveFrom).Trim());
        Assert.Equal(moved, MovedText(body, W.moveTo).Trim());

        var moves = revisions.Where(r => r.Type == DocxDiffRevisionType.Moved).ToList();
        Assert.Equal(2, moves.Count);
        Assert.All(moves, r => Assert.Equal(moved, r.Text.Trim()));
        Assert.DoesNotContain(revisions, r => r.Type != DocxDiffRevisionType.Moved && r.Text.Contains("moved sentence"));

        DocxDiffMovePairingParityTests.AssertMovesWholeAndAgreeing(left, right, new DocxDiffSettings());
    }

    public static IEnumerable<object[]> FusedMovesUnderSettings()
    {
        var settings = new (string Name, Func<DocxDiffSettings> Make)[]
        {
            ("default", () => new DocxDiffSettings()),
            ("moves not reported", () => new DocxDiffSettings { DetectMoves = false }),
            ("input revisions preserved", () => new DocxDiffSettings { PreserveInputRevisions = true }),
            ("no cross-paragraph runs", () => new DocxDiffSettings { CrossParagraphTokenDiff = false }),
            ("WmlComparer-compatible revisions", () => new DocxDiffSettings { RevisionGranularity = DocxDiffRevisionGranularity.WmlComparerCompatible }),
        };
        foreach (var row in FusedMoves)
            foreach (var (name, make) in settings)
                yield return new object[] { (string)row[0], name, (string)row[1], (string)row[2], make };
    }

    [Theory]
    [MemberData(nameof(FusedMovesUnderSettings))]
    public void FusedMoves_AgreeAcrossSurfacesUnderEverySetting(
        string shape, string settingsName, string originalBody, string revisedBody, Func<DocxDiffSettings> settings)
    {
        _ = (shape, settingsName);
        DocxDiffMovePairingParityTests.AssertMovesWholeAndAgreeing(
            IrTestDocuments.FromParts(originalBody), IrTestDocuments.FromParts(revisedBody), settings());
    }

    [Fact]
    public void SentenceTheStreamRetainsAcrossAPilcrow_StaysUnpairedOnBothSurfaces()
    {
        // The sentence moves from the end of one edited paragraph to the head of the next. The fused stream
        // keeps it in place and moves the paragraph mark instead, so the redline shows no deletion of it; the
        // revision list, built paragraph by paragraph, must not report a move the redline does not show.
        var left = IrTestDocuments.FromParts(P($"{EditedG} {M}") + P(B) + P(D));
        var right = IrTestDocuments.FromParts(P(G) + P($"{M} {EditedB}") + P(D));
        var settings = DocxCompare.ApplyFrontDoorRevisionPolicy(null);

        var redline = DocxDiff.Compare(left, right, settings);
        var revisions = DocxDiff.GetRevisions(left, right, settings);

        Assert.Empty(Body(redline.DocumentByteArray).Descendants(W.moveFromRangeStart));
        Assert.DoesNotContain(revisions, r => r.Type == DocxDiffRevisionType.Moved);
        DocxDiffMovePairingParityTests.AssertMovesWholeAndAgreeing(left, right, new DocxDiffSettings());
    }

    private static string MovedText(XElement body, XName wrapper) =>
        string.Concat(body.Descendants(wrapper).Descendants()
            .Where(e => e.Name == W.t || e.Name == W.delText)
            .Select(e => e.Value));

    private static XElement Body(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }
}
