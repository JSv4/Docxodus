// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Text moved within a paragraph, between paragraphs, or into or out of a table cell is reported as a move
/// rather than an unrelated deletion and insertion (issue #888): the moved span is wrapped in
/// <c>w:moveFrom</c>/<c>w:moveTo</c> under one move name, and <c>GetRevisions</c> reports a Moved pair. The two
/// texts must match exactly (whitespace runs aside) and carry at least <c>MoveMinimumWordCount</c> words.
/// </summary>
public class DocxDiffTextMoveTests
{
    private const string First = "The first sentence explains the scope of this agreement.";
    private const string Second = "The second sentence covers payment terms.";
    private const string Third = "The third sentence covers termination.";
    private const string Steady = "A steady paragraph that does not change at all in this revision.";
    private const string Wanderer = "The wandering sentence describes the obligations in detail.";

    private static string P(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string OneCellTable(params string[] paragraphs) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>" +
        $"<w:tr><w:tc><w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/></w:tcPr>{string.Concat(paragraphs.Select(P))}</w:tc></w:tr></w:tbl>";

    private static WmlDocument Doc(string body) => IrTestDocuments.FromParts(body);

    public static TheoryData<string, string, string, string> TextMoves => new()
    {
        { "a sentence within its paragraph", P($"{First} {Second} {Third}"), P($"{Second} {Third} {First}"), First },
        { "a sentence to the next paragraph", P($"Alpha paragraph opens here. {Wanderer}") + P("Beta paragraph stays put and says something."),
            P("Alpha paragraph opens here.") + P($"Beta paragraph stays put and says something. {Wanderer}"), Wanderer },
        { "a sentence past an unchanged paragraph", P($"Alpha paragraph opens here. {Wanderer}") + P(Steady) + P("Beta paragraph stays put and says something."),
            P("Alpha paragraph opens here.") + P(Steady) + P($"Beta paragraph stays put and says something. {Wanderer}"), Wanderer },
        { "a paragraph merged into the middle of another", P(Wanderer) + P(Steady) + P("Gamma starts here. Gamma ends here."),
            P(Steady) + P($"Gamma starts here. {Wanderer} Gamma ends here."), Wanderer },
        { "a sentence split out into a paragraph of its own", P($"Alpha paragraph opens here. {Wanderer}") + P(Steady),
            P("Alpha paragraph opens here.") + P(Steady) + P(Wanderer), Wanderer },
        { "a sentence into a cell, replacing the cell's text", P($"Intro holds a sentence. {Wanderer}") + OneCellTable("Old cell words here") + P("Tail stays."),
            P("Intro holds a sentence.") + OneCellTable(Wanderer) + P("Tail stays."), Wanderer },
        { "a sentence out of a cell", P("Intro holds a sentence.") + OneCellTable($"Cell keeps this. {Wanderer}") + P("Tail stays."),
            P($"Intro holds a sentence. {Wanderer}") + OneCellTable("Cell keeps this.") + P("Tail stays."), Wanderer },
    };

    [Theory]
    [MemberData(nameof(TextMoves))]
    public void MovedText_IsAMove(string shape, string originalBody, string revisedBody, string moved)
    {
        _ = shape;
        var (left, right) = (Doc(originalBody), Doc(revisedBody));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline.DocumentByteArray);
        var from = Assert.Single(body.Descendants(W.moveFromRangeStart));
        var to = Assert.Single(body.Descendants(W.moveToRangeStart));
        Assert.Equal((string?)from.Attribute(W.name), (string?)to.Attribute(W.name));
        Assert.Equal(moved, Normalize(body.Descendants(W.moveFrom).SelectMany(m => m.Descendants(W.delText))));
        Assert.Equal(moved, Normalize(body.Descendants(W.moveTo).SelectMany(m => m.Descendants(W.t))));
        // The moved text is not also drawn as a plain deletion or insertion.
        Assert.DoesNotContain(body.Descendants(W.del), d => d.Value.Contains("sentence explains") || d.Value.Contains("wandering"));
        Assert.DoesNotContain(body.Descendants(W.ins), i => i.Value.Contains("sentence explains") || i.Value.Contains("wandering"));
    }

    [Theory]
    [MemberData(nameof(TextMoves))]
    public void MovedText_AgreesWithTheRevisionList(string shape, string originalBody, string revisedBody, string moved)
    {
        _ = shape;
        var (left, right) = (Doc(originalBody), Doc(revisedBody));

        var revisions = DocxDiff.GetRevisions(left, right, DocxCompare.ApplyFrontDoorRevisionPolicy(null));

        var group = Assert.Single(revisions.Where(r => r.Type == DocxDiffRevisionType.Moved).GroupBy(r => r.MoveGroupId));
        Assert.Single(group, r => r.IsMoveSource == true && r.Text.Trim() == moved);
        Assert.Single(group, r => r.IsMoveSource == false && r.Text.Trim() == moved);
        DocxDiffMovePairingParityTests.AssertMovesWholeAndAgreeing(left, right, new DocxDiffSettings());
        DocxDiffMovePairingParityTests.AssertMovesWholeAndAgreeing(left, right, new DocxDiffSettings { CrossParagraphTokenDiff = false });
    }

    [Theory]
    [MemberData(nameof(TextMoves))]
    public void MovedText_WithMovesNotReported_IsADeleteAndAnInsert(string shape, string originalBody, string revisedBody, string moved)
    {
        _ = (shape, moved);
        var (left, right) = (Doc(originalBody), Doc(revisedBody));
        var settings = new DocxDiffSettings { DetectMoves = false };

        var redline = DocxCompare.Compare(left, right, settings);

        Assert.Empty(Body(redline.DocumentByteArray).Descendants(W.moveFrom));
        Assert.Empty(Body(redline.DocumentByteArray).Descendants(W.moveTo));
        DocxDiffMovePairingParityTests.AssertMovesWholeAndAgreeing(left, right, settings);
    }

    [Fact]
    public void TextBelowTheMinimumWordCount_IsNotAMove()
    {
        // "Signed copy" is two words: below the default minimum of three, it stays a deletion and an insertion.
        var (left, right) = (Doc(P($"Signed copy. {Second} {Third}")), Doc(P($"{Second} {Third} Signed copy.")));

        var redline = DocxCompare.Compare(left, right);

        Assert.Empty(Body(redline.DocumentByteArray).Descendants(W.moveFromRangeStart));
    }

    [Fact]
    public void TextBelowTheMinimum_IsAMove_WhenTheMinimumIsLowered()
    {
        var (left, right) = (Doc(P($"Signed copy. {Second} {Third}")), Doc(P($"{Second} {Third} Signed copy.")));

        var redline = DocxCompare.Compare(left, right, new DocxDiffSettings { MoveMinimumWordCount = 2 });

        Assert.Single(Body(redline.DocumentByteArray).Descendants(W.moveFromRangeStart));
    }

    [Fact]
    public void RewrittenText_IsNotAMove()
    {
        // The sentence leaves and a reworded one arrives: not the same text, so not a move.
        var (left, right) = (Doc(P($"{First} {Second} {Third}")),
            Doc(P($"{Second} {Third} The first sentence sets out the scope of this agreement.")));

        var redline = DocxCompare.Compare(left, right);

        Assert.Empty(Body(redline.DocumentByteArray).Descendants(W.moveFromRangeStart));
    }

    [Fact]
    public void MovedText_EditScriptTagsBothSpans()
    {
        var (left, right) = (Doc(P($"{First} {Second} {Third}")), Doc(P($"{Second} {Third} {First}")));

        var json = DocxDiff.GetEditScriptJson(left, right, DocxCompare.ApplyFrontDoorRevisionPolicy(null));

        // The relocation rides as the sixth element of the two token-op arrays and survives a round trip.
        var script = IrEditScriptJson.Read(json);
        Assert.Equal(json, IrEditScriptJson.Write(script));
        var tagged = script.Operations.Single(op => op.TokenDiff is not null).TokenDiff!.Ops
            .Where(o => o.RelocationGroupId is not null).ToList();
        Assert.Equal(2, tagged.Count);
        Assert.Single(tagged, o => o.Kind == IrTokenOpKind.Delete);
        Assert.Single(tagged, o => o.Kind == IrTokenOpKind.Insert);
        Assert.Single(tagged.Select(o => o.RelocationGroupId).Distinct());
    }

    [Fact]
    public void MovedText_InConsolidate_IsADeleteAndAnInsert()
    {
        var original = Doc(P($"{First} {Second} {Third}"));
        var reviewer = new DocxDiffReviewer { Author = "Reviewer", Document = Doc(P($"{Second} {Third} {First}")) };

        var consolidated = DocxDiff.Consolidate(original, new[] { reviewer });

        Assert.Empty(Body(consolidated.DocumentByteArray).Descendants(W.moveFromRangeStart));
        Assert.Equal(Text(reviewer.Document.DocumentByteArray), Text(Docxodus.Internal.DocxDiffOps.AcceptRevisions(consolidated.DocumentByteArray)));
        Assert.Equal(Text(original.DocumentByteArray), Text(Docxodus.Internal.DocxDiffOps.RejectRevisions(consolidated.DocumentByteArray)));
    }

    private static string Normalize(System.Collections.Generic.IEnumerable<XElement> texts) =>
        Regex.Replace(string.Concat(texts.Select(t => t.Value)), @"\s+", " ").Trim();

    private static XElement Body(byte[] bytes)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    private static string Text(byte[] bytes) => string.Join("|", Body(bytes).Descendants(W.p)
        .Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value))));
}
