// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1143: when a block wrapper (a <c>w:customXml</c> element or a block <c>w:sdt</c> content control)
/// keeps its own markup, a comparison diffs the blocks inside one copy of it, as it does a table cell's
/// blocks. Only a change to the wrapper itself (customXml attributes and <c>w:customXmlPr</c>; a control's
/// <c>w:sdtPr</c>/<c>w:sdtEndPr</c>) replaces the whole wrapper.
/// </summary>
public class BlockWrapperNestedDiffTests
{
    public enum Wrapper
    {
        CustomXml,
        ContentControl,
    }

    private const string Alpha = "Alpha clause: the parties agree to these terms.";
    private const string Bravo = "Bravo clause: backups are verified monthly by staff.";
    private const string Charlie = "Charlie clause: restoration is tested once a year.";
    private const string Tail = "Closing paragraph outside the wrapper.";

    private static string P(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string Wrap(Wrapper kind, string inner, string id = "a1") => kind == Wrapper.CustomXml
        ? $"<w:customXml w:uri=\"urn:example\" w:element=\"article\"><w:customXmlPr><w:attr w:name=\"id\" w:val=\"{id}\"/></w:customXmlPr>{inner}</w:customXml>"
        : $"<w:sdt><w:sdtPr><w:alias w:val=\"Article\"/><w:tag w:val=\"{id}\"/><w:id w:val=\"101\"/></w:sdtPr><w:sdtContent>{inner}</w:sdtContent></w:sdt>";

    private static XName WrapperName(Wrapper kind) => kind == Wrapper.CustomXml ? W.customXml : W.sdt;

    private static WmlDocument Doc(string body) => IrTestDocuments.FromBodyXml(body);

    private static string Three(Wrapper kind, string bravo = Bravo, string id = "a1") =>
        Wrap(kind, P(Alpha) + P(bravo) + P(Charlie), id);

    private static readonly string BravoEdited = Bravo.Replace("monthly", "weekly");

    // ---- the acceptance test -------------------------------------------------------------------

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_OneWordEditedInsideAnUnchangedWrapper_MarksOnlyThatWord(Wrapper kind)
    {
        var left = Doc(Three(kind) + P(Tail));
        var right = Doc(Three(kind, BravoEdited) + P(Tail));

        foreach (var redline in new[] { DocxCompare.Compare(left, right), DocxDiff.Compare(left, right, new DocxDiffSettings()) })
        {
            var body = Body(redline);
            AssertSingleWrapperWithoutRangeMarkers(body, kind);
            Assert.Equal("monthly", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
            Assert.Equal("weekly", string.Concat(body.Descendants(W.ins).Descendants(W.t).Select(t => t.Value)));
            AssertUnmarked(body, Alpha, Charlie, Tail);
            AssertRoundTrip(left, right, redline);
        }
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void GetRevisions_OneWordEditedInsideAnUnchangedWrapper_ReportsOnlyThatWord(Wrapper kind)
    {
        var left = Doc(Three(kind) + P(Tail));
        var right = Doc(Three(kind, BravoEdited) + P(Tail));

        var revisions = DocxDiff.GetRevisions(left, right);

        Assert.Collection(revisions.OrderBy(r => r.Type),
            r => { Assert.Equal(DocxDiffRevisionType.Inserted, r.Type); Assert.Equal("weekly", r.Text); },
            r => { Assert.Equal(DocxDiffRevisionType.Deleted, r.Type); Assert.Equal("monthly", r.Text); });
    }

    // ---- the wrapper itself changed ------------------------------------------------------------

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_WrapperPropertiesChanged_StillReplacesTheWholeWrapper(Wrapper kind)
    {
        var left = Doc(Three(kind) + P(Tail));
        var right = Doc(Three(kind, BravoEdited, id: "a2") + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        Assert.Equal(2, body.Descendants(WrapperName(kind)).Count());
        Assert.NotEmpty(body.Descendants(W.customXmlDelRangeStart));
        Assert.NotEmpty(body.Descendants(W.customXmlInsRangeStart));
        AssertRoundTrip(left, right, redline);

        var revisions = DocxDiff.GetRevisions(left, right);
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Deleted && r.Text.Contains(Alpha));
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Inserted && r.Text.Contains(Alpha));
    }

    // ---- nesting -------------------------------------------------------------------------------

    [Theory]
    [InlineData(Wrapper.CustomXml, Wrapper.ContentControl)]
    [InlineData(Wrapper.ContentControl, Wrapper.CustomXml)]
    [InlineData(Wrapper.CustomXml, Wrapper.CustomXml)]
    public void Compare_EditInsideNestedUnchangedWrappers_MarksOnlyTheWord(Wrapper outer, Wrapper inner)
    {
        string Nested(string bravo) => Wrap(outer, P(Alpha) + Wrap(inner, P(bravo) + P(Charlie), id: "inner"), id: "outer");
        var left = Doc(Nested(Bravo) + P(Tail));
        var right = Doc(Nested(BravoEdited) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        Assert.Empty(body.Descendants().Where(IsRangeMarker));
        Assert.Equal(outer == inner ? 2 : 1, body.Descendants(WrapperName(outer)).Count());
        Assert.Equal(outer == inner ? 2 : 1, body.Descendants(WrapperName(inner)).Count());
        Assert.Equal("monthly", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        AssertUnmarked(body, Alpha, Charlie, Tail);
        AssertRoundTrip(left, right, redline);
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_InnerWrapperChangedInsideAnUnchangedOuter_ReplacesOnlyTheInnerWrapper(Wrapper kind)
    {
        string Nested(string innerId) => Wrap(kind, P(Alpha) + Wrap(kind, P(Bravo), id: innerId), id: "outer");
        var left = Doc(Nested("inner-1") + P(Tail));
        var right = Doc(Nested("inner-2") + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        var outerWrappers = body.Elements(WrapperName(kind)).ToList();
        Assert.Single(outerWrappers);
        Assert.NotEmpty(body.Descendants(W.customXmlDelRangeStart));
        AssertUnmarked(body, Alpha, Tail);
        AssertRoundTrip(left, right, redline);
    }

    // ---- block insertions and deletions inside the wrapper ---------------------------------------

    public static IEnumerable<object[]> InsertedAndDeletedChildren()
    {
        const string NewClause = "Delta clause: a brand new obligation appears here.";
        foreach (var kind in new[] { Wrapper.CustomXml, Wrapper.ContentControl })
        {
            // name, left children, right children, paragraphs that must stay unmarked
            yield return new object[] { kind, "insert middle", new[] { Alpha, Charlie }, new[] { Alpha, NewClause, Charlie }, new[] { Alpha, Charlie } };
            yield return new object[] { kind, "insert first", new[] { Alpha, Charlie }, new[] { NewClause, Alpha, Charlie }, new[] { Alpha, Charlie } };
            yield return new object[] { kind, "insert last", new[] { Alpha, Charlie }, new[] { Alpha, Charlie, NewClause }, new[] { Alpha, Charlie } };
            yield return new object[] { kind, "delete middle", new[] { Alpha, Bravo, Charlie }, new[] { Alpha, Charlie }, new[] { Alpha, Charlie } };
            yield return new object[] { kind, "delete first", new[] { Alpha, Bravo, Charlie }, new[] { Bravo, Charlie }, new[] { Charlie } };
            yield return new object[] { kind, "delete last", new[] { Alpha, Bravo, Charlie }, new[] { Alpha, Bravo }, new[] { Alpha } };
        }
    }

    [Theory]
    [MemberData(nameof(InsertedAndDeletedChildren))]
    public void Compare_ParagraphInsertedOrDeletedInsideAnUnchangedWrapper_KeepsOneWrapper(
        Wrapper kind, string name, string[] leftChildren, string[] rightChildren, string[] unmarked)
    {
        var left = Doc(Wrap(kind, string.Concat(leftChildren.Select(P))) + P(Tail));
        var right = Doc(Wrap(kind, string.Concat(rightChildren.Select(P))) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        AssertSingleWrapperWithoutRangeMarkers(body, kind);
        AssertUnmarked(body, unmarked.Append(Tail).ToArray());
        AssertRoundTrip(left, right, redline, name);
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_TableCellEditedInsideAnUnchangedWrapper_DiffsTheCell(Wrapper kind)
    {
        string Table(string cell) =>
            "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4000\"/><w:gridCol w:w=\"4000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>" + P("Left cell stays.") + "</w:tc>"
            + "<w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>" + P(cell) + "</w:tc></w:tr></w:tbl>";
        var left = Doc(Wrap(kind, P(Alpha) + Table(Bravo)) + P(Tail));
        var right = Doc(Wrap(kind, P(Alpha) + Table(BravoEdited)) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        AssertSingleWrapperWithoutRangeMarkers(body, kind);
        Assert.Single(body.Descendants(W.tbl));
        Assert.Equal("monthly", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        AssertUnmarked(body, Alpha, "Left cell stays.", Tail);
        AssertRoundTrip(left, right, redline);
    }

    // ---- the wrapper pair is found even when it is not the lone residue of its gap ----------------

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_ParagraphInsertedBeforeAnEditedWrapper_StillDiffsInsideTheWrapper(Wrapper kind)
    {
        const string Intro = "A new introductory paragraph was added.";
        var left = Doc(Three(kind) + P(Tail));
        var right = Doc(P(Intro) + Three(kind, BravoEdited) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        AssertSingleWrapperWithoutRangeMarkers(body, kind);
        Assert.Equal("monthly", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        AssertUnmarked(body, Alpha, Charlie, Tail);
        AssertRoundTrip(left, right, redline);
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_TwoAdjacentEditedWrappers_DiffInsideEach(Wrapper kind)
    {
        string Two(string first, string second) =>
            Wrap(kind, P(Alpha) + P(first), id: "w1") + Wrap(kind, P(second) + P(Charlie), id: "w2");
        var left = Doc(Two(Bravo, Bravo) + P(Tail));
        var right = Doc(Two(BravoEdited, Bravo.Replace("staff", "auditors")) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        Assert.Equal(2, body.Descendants(WrapperName(kind)).Count());
        Assert.Empty(body.Descendants().Where(IsRangeMarker));
        Assert.Equal("monthlystaff", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        AssertUnmarked(body, Alpha, Charlie, Tail);
        AssertRoundTrip(left, right, redline);
    }

    // ---- consolidate ---------------------------------------------------------------------------

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Consolidate_ReviewerEditInsideAnUnchangedWrapper_MarksOnlyThatWord(Wrapper kind)
    {
        var baseline = Doc(Three(kind) + P(Tail));
        var reviewed = Doc(Three(kind, BravoEdited) + P(Tail));
        var reviewers = new[] { new DocxDiffReviewer { Document = reviewed, Author = "Reviewer" } };

        var redline = DocxCompare.Consolidate(baseline, reviewers);

        var body = Body(redline);
        AssertSingleWrapperWithoutRangeMarkers(body, kind);
        Assert.Equal("monthly", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        AssertUnmarked(body, Alpha, Charlie, Tail);
        AssertRoundTrip(baseline, reviewed, redline);

        var revisions = DocxDiff.GetConsolidatedRevisions(baseline, reviewers)
            .Where(r => r.Type is DocxDiffRevisionType.Inserted or DocxDiffRevisionType.Deleted).ToList();
        Assert.Collection(revisions.OrderBy(r => r.Type),
            r => Assert.Equal("weekly", r.Text),
            r => Assert.Equal("monthly", r.Text));
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Consolidate_TwoReviewersEditingDifferentWrappers_DiffInsideEach(Wrapper kind)
    {
        string Two(string first, string second) =>
            Wrap(kind, P(Alpha) + P(first), id: "w1") + Wrap(kind, P(second) + P(Charlie), id: "w2");
        var baseline = Doc(Two(Bravo, Bravo) + P(Tail));
        var first = Doc(Two(BravoEdited, Bravo) + P(Tail));
        var second = Doc(Two(Bravo, Bravo.Replace("staff", "auditors")) + P(Tail));
        var merged = Doc(Two(BravoEdited, Bravo.Replace("staff", "auditors")) + P(Tail));

        var redline = DocxCompare.Consolidate(baseline,
            [new DocxDiffReviewer { Document = first, Author = "One" }, new DocxDiffReviewer { Document = second, Author = "Two" }]);

        var body = Body(redline);
        Assert.Equal(2, body.Descendants(WrapperName(kind)).Count());
        Assert.Empty(body.Descendants().Where(IsRangeMarker));
        AssertUnmarked(body, Alpha, Charlie, Tail);
        AssertRoundTrip(baseline, merged, redline);
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_ParagraphMovedInsideAnUnchangedWrapper_IsANativeMove(Wrapper kind)
    {
        var left = Doc(Wrap(kind, P(Alpha) + P(Bravo) + P(Charlie)) + P(Tail));
        var right = Doc(Wrap(kind, P(Bravo) + P(Charlie) + P(Alpha)) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        AssertSingleWrapperWithoutRangeMarkers(body, kind);
        Assert.NotEmpty(body.Descendants(W.moveFrom));
        Assert.NotEmpty(body.Descendants(W.moveTo));
        AssertMoveNamesPaired(body);
        AssertRoundTrip(left, right, redline);
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Consolidate_TwoReviewersMovingParagraphsInDifferentWrappers_KeepsEachMovePaired(Wrapper kind)
    {
        string Two(string first, string second) => Wrap(kind, first, id: "w1") + Wrap(kind, second, id: "w2");
        string Abc = P(Alpha) + P(Bravo) + P(Charlie);
        string Moved = P(Bravo) + P(Charlie) + P(Alpha);
        var baseline = Doc(Two(Abc, Abc) + P(Tail));
        var first = Doc(Two(Moved, Abc) + P(Tail));
        var second = Doc(Two(Abc, Moved) + P(Tail));
        var merged = Doc(Two(Moved, Moved) + P(Tail));

        var redline = DocxCompare.Consolidate(baseline,
            [new DocxDiffReviewer { Document = first, Author = "One" }, new DocxDiffReviewer { Document = second, Author = "Two" }]);

        var body = Body(redline);
        Assert.Equal(2, body.Descendants(WrapperName(kind)).Count());
        AssertMoveNamesPaired(body);
        AssertRoundTrip(baseline, merged, redline);
    }

    /// <summary>Every move name labels exactly one source range and one destination range.</summary>
    private static void AssertMoveNamesPaired(XElement body)
    {
        var from = body.Descendants(W.moveFromRangeStart).Select(e => (string?)e.Attribute(W.name)).ToList();
        var to = body.Descendants(W.moveToRangeStart).Select(e => (string?)e.Attribute(W.name)).ToList();
        Assert.Equal(from.Count, from.Distinct().Count());
        Assert.Equal(to.Count, to.Distinct().Count());
        Assert.Equal(from.OrderBy(n => n), to.OrderBy(n => n));
    }

    // ---- edit script ---------------------------------------------------------------------------

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void EditScript_CarriesTheWrapperChildOps_AndRoundTripsThroughJson(Wrapper kind)
    {
        var left = Doc(Three(kind) + P(Tail));
        var right = Doc(Three(kind, BravoEdited) + P(Tail));
        var options = new IrReaderOptions { RetainSources = false, RevisionView = RevisionView.Accept };
        var script = IrEditScriptBuilder.Build(IrReader.Read(left, options), IrReader.Read(right, options), new IrDiffSettings());

        var wrapperOp = Assert.Single(script.Operations, op => op.WrapperOps is not null);
        Assert.Equal(IrEditOpKind.ModifyBlock, wrapperOp.Kind);
        Assert.Equal(
            new[] { IrEditOpKind.EqualBlock, IrEditOpKind.ModifyBlock, IrEditOpKind.EqualBlock },
            wrapperOp.WrapperOps!.Select(op => op.Kind));

        var json = IrEditScriptJson.Write(script);
        Assert.Contains("\"wrapperOps\"", json);
        Assert.Equal(script, IrEditScriptJson.Read(json));
    }

    // ---- helpers -------------------------------------------------------------------------------

    private static bool IsRangeMarker(XElement e) =>
        e.Name.LocalName.StartsWith("customXmlIns", StringComparison.Ordinal)
        || e.Name.LocalName.StartsWith("customXmlDel", StringComparison.Ordinal);

    private static void AssertSingleWrapperWithoutRangeMarkers(XElement body, Wrapper kind)
    {
        Assert.Single(body.Descendants(WrapperName(kind)));
        Assert.Empty(body.Descendants().Where(IsRangeMarker));
    }

    /// <summary>Each paragraph whose text is one of <paramref name="texts"/> carries no revision markup at all:
    /// no run-level <c>w:ins</c>/<c>w:del</c>, no paragraph-mark revision and no property change.</summary>
    private static void AssertUnmarked(XElement body, params string[] texts)
    {
        foreach (var text in texts)
        {
            var paragraph = Assert.Single(body.Descendants(W.p), p => string.Concat(p.Descendants(W.t).Select(t => t.Value)) == text
                && !p.Descendants(W.delText).Any());
            Assert.DoesNotContain(paragraph.Descendants(), e => e.Name == W.ins || e.Name == W.del
                || e.Name == W.moveFrom || e.Name == W.moveTo || e.Name.LocalName.EndsWith("Change", StringComparison.Ordinal));
        }
    }

    /// <summary>accept(redline) equals <paramref name="revised"/> and reject(redline) equals
    /// <paramref name="original"/>, block for block: every wrapper with its exact properties, every table, and
    /// every paragraph's properties and formatted text (adjacent runs with the same properties merged, since
    /// the redline may split a run around a change).</summary>
    private static void AssertRoundTrip(WmlDocument original, WmlDocument revised, WmlDocument redline, string? because = null)
    {
        var accepted = DocxDiffOps.AcceptRevisions(redline.DocumentByteArray);
        var rejected = DocxDiffOps.RejectRevisions(redline.DocumentByteArray);
        Assert.True(Signature(Body(revised.DocumentByteArray)) == Signature(Body(accepted)),
            $"accept ≠ revised ({because}):\n{Signature(Body(revised.DocumentByteArray))}\n---\n{Signature(Body(accepted))}");
        Assert.True(Signature(Body(original.DocumentByteArray)) == Signature(Body(rejected)),
            $"reject ≠ original ({because}):\n{Signature(Body(original.DocumentByteArray))}\n---\n{Signature(Body(rejected))}");
    }

    private static string Signature(XElement container)
    {
        var sb = new StringBuilder();
        AppendBlocks(container, sb, 0);
        return sb.ToString();
    }

    private static void AppendBlocks(XElement container, StringBuilder sb, int depth)
    {
        foreach (var child in container.Elements())
        {
            var indent = new string(' ', depth * 2);
            if (child.Name == W.p)
            {
                sb.Append(indent).Append("P ").Append(Canonical(child.Element(W.pPr))).Append(" |");
                string? lastProps = null;
                foreach (var run in child.Descendants(W.r))
                {
                    var props = Canonical(run.Element(W.rPr));
                    if (props != lastProps)
                        sb.Append(" {").Append(props).Append("} ");
                    lastProps = props;
                    sb.Append(string.Concat(run.Elements().Where(e => e.Name != W.rPr)
                        .Select(e => e.Name == W.t ? e.Value : "<" + e.Name.LocalName + ">")));
                }
                sb.Append('\n');
            }
            else if (child.Name == W.tbl)
            {
                sb.Append(indent).Append("TBL ").Append(Canonical(child.Element(W.tblPr)))
                    .Append(Canonical(child.Element(W.tblGrid))).Append('\n');
                foreach (var row in child.Elements(W.tr))
                {
                    sb.Append(indent).Append(" TR ").Append(Canonical(row.Element(W.trPr))).Append('\n');
                    foreach (var cell in row.Elements(W.tc))
                    {
                        sb.Append(indent).Append("  TC ").Append(Canonical(cell.Element(W.tcPr))).Append('\n');
                        AppendBlocks(cell, sb, depth + 2);
                    }
                }
            }
            else if (child.Name == W.customXml)
            {
                sb.Append(indent).Append("CUSTOMXML ")
                    .Append(Canonical(new XElement(W.customXml, child.Attributes(), child.Element(W.customXmlPr)))).Append('\n');
                AppendBlocks(child, sb, depth + 1);
            }
            else if (child.Name == W.sdt)
            {
                sb.Append(indent).Append("SDT ").Append(Canonical(child.Element(W.sdtPr)))
                    .Append(Canonical(child.Element(W.sdtEndPr))).Append('\n');
                AppendBlocks(child.Element(W.sdtContent)!, sb, depth + 1);
            }
            else if (child.Name != W.sectPr && child.Name != W.customXmlPr && child.Name != W.pPr)
            {
                sb.Append(indent).Append("OTHER ").Append(child.Name.LocalName).Append('\n');
            }
        }
    }

    /// <summary>Canonical XML of a property element: rsid and PowerTools attributes dropped, attributes sorted.</summary>
    private static string Canonical(XElement? element)
    {
        if (element is null)
            return "-";
        static XElement Clean(XElement e) => new(e.Name,
            e.Attributes()
                .Where(a => !a.IsNamespaceDeclaration
                    && !a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)
                    && a.Name.NamespaceName != "http://powertools.codeplex.com/2011")
                .OrderBy(a => a.Name.ToString(), StringComparer.Ordinal)
                .Select(a => new XAttribute(a.Name, a.Value)),
            e.Elements().Where(c => !c.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)).Select(Clean),
            e.HasElements ? null : e.Value);
        return Clean(element).ToString(SaveOptions.DisableFormatting);
    }

    private static XElement Body(WmlDocument document) => Body(document.DocumentByteArray);

    private static XElement Body(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return new XElement(document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!);
    }
}
