#nullable enable

using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

public class IrNodeTests
{
    private static readonly IrHash EmptyDigest = IrHash.Compute(Array.Empty<byte>());

    private static IrRunFormat RunFmt() => new() { UnmodeledDigest = EmptyDigest };
    private static IrParaFormat ParaFmt() => new() { UnmodeledDigest = EmptyDigest };

    private static IrParagraph MakeParagraph(IrAnchor anchor, IReadOnlyList<IrInline> inlines, IrProvenance? source = null) =>
        new()
        {
            Anchor = anchor,
            ContentHash = IrHash.Compute("content"),
            FormatFingerprint = IrHash.Compute("fmt"),
            Format = ParaFmt(),
            Inlines = IrNodeList.From(inlines),
            Source = source ?? new IrProvenance(),
        };

    [Fact]
    public void IrParagraph_Equality_IgnoresProvenance()
    {
        var anchor = new IrAnchor(IrAnchorKind.P, "body", "00000000000000000000000000000001");
        var inlines = new IrInline[] { new IrTextRun("hello", RunFmt()) };

        var elemA = new XElement("a");
        var elemB = new XElement("b");

        var p1 = MakeParagraph(anchor, inlines, new IrProvenance { Element = elemA, PartUri = new Uri("/word/document.xml", UriKind.Relative) });
        var p2 = MakeParagraph(anchor, inlines, new IrProvenance { Element = elemB, PartUri = new Uri("/word/other.xml", UriKind.Relative) });

        Assert.Equal(p1, p2);
        Assert.Equal(p1.GetHashCode(), p2.GetHashCode());
    }

    [Fact]
    public void IrParagraph_Equality_ValueSemanticInlines()
    {
        var anchor = new IrAnchor(IrAnchorKind.P, "body", "00000000000000000000000000000002");

        var p1 = MakeParagraph(anchor, new IrInline[] { new IrTextRun("foo", RunFmt()), new IrTab(RunFmt()) });
        var p2 = MakeParagraph(anchor, new IrInline[] { new IrTextRun("foo", RunFmt()), new IrTab(RunFmt()) });
        var p3 = MakeParagraph(anchor, new IrInline[] { new IrTextRun("bar", RunFmt()), new IrTab(RunFmt()) });

        Assert.Equal(p1, p2);
        Assert.Equal(p1.GetHashCode(), p2.GetHashCode());
        Assert.NotEqual(p1, p3);
    }

    [Fact]
    public void IrTable_NestedCells_Construct()
    {
        var pAnchor = new IrAnchor(IrAnchorKind.P, "body", "00000000000000000000000000000010");
        var para = MakeParagraph(pAnchor, new IrInline[] { new IrTextRun("cell text", RunFmt()) });

        var cell = new IrCell(
            new IrAnchor(IrAnchorKind.Tc, "body", "00000000000000000000000000000011"),
            IrNodeList.From<IrBlock>(new IrBlock[] { para }),
            GridSpan: 1,
            VMerge: IrVMerge.None,
            ContentHash: IrHash.Compute("cell"));

        var row = new IrRow(
            new IrAnchor(IrAnchorKind.Tr, "body", "00000000000000000000000000000012"),
            IrNodeList.From(new[] { cell }),
            IrHash.Compute("row"));

        var table = new IrTable
        {
            Anchor = new IrAnchor(IrAnchorKind.Tbl, "body", "00000000000000000000000000000013"),
            ContentHash = IrHash.Compute("table"),
            FormatFingerprint = IrHash.Compute("tablefmt"),
            Rows = IrNodeList.From(new[] { row }),
            UnmodeledTablePropsDigest = EmptyDigest,
        };

        Assert.Single(table.Rows);
        Assert.Single(table.Rows[0].Cells);
        Assert.Single(table.Rows[0].Cells[0].Blocks);
        var inner = Assert.IsType<IrParagraph>(table.Rows[0].Cells[0].Blocks[0]);
        var run = Assert.IsType<IrTextRun>(inner.Inlines[0]);
        Assert.Equal("cell text", run.Text);
    }

    [Fact]
    public void IrDocument_FindByAnchor_ReturnsBlock()
    {
        var anchor = new IrAnchor(IrAnchorKind.P, "body", "00000000000000000000000000000020");
        var para = MakeParagraph(anchor, new IrInline[] { new IrTextRun("hi", RunFmt()) });

        var body = new IrScope("body", IrNodeList.From<IrBlock>(new IrBlock[] { para }));

        var doc = new IrDocument
        {
            Body = body,
            Footnotes = IrNoteStore.Empty,
            Endnotes = IrNoteStore.Empty,
            Comments = IrCommentStore.Empty,
            Styles = IrStyleRegistry.Empty,
            Numbering = IrNumberingRegistry.Empty,
            ThemeFonts = IrThemeFonts.Empty,
            AnchorIndex = new Dictionary<string, IrBlock> { [anchor.ToString()] = para },
            Sources = new Dictionary<Uri, XDocument>(),
        };

        Assert.Same(para, doc.FindByAnchor(anchor));

        var unknown = new IrAnchor(IrAnchorKind.P, "body", "ffffffffffffffffffffffffffffffff");
        Assert.Null(doc.FindByAnchor(unknown));
    }

    [Fact]
    public void IrNodeList_Equality_BySequence()
    {
        var a = IrNodeList.From(new[] { 1, 2, 3 });
        var b = IrNodeList.From(new[] { 1, 2, 3 });
        var c = IrNodeList.From(new[] { 1, 2, 4 });

        Assert.Equal(a, b);
        Assert.Equal(a.GetHashCode(), b.GetHashCode());
        Assert.NotEqual(a, c);
    }
}
