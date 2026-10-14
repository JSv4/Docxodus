#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Tests for the non-body scopes (headers/footers, footnotes/endnotes, comments) and the N15
/// comment-target record half (rule M1.3, Task 4). Scope naming and ordering must mirror
/// <see cref="WmlToMarkdownConverter"/>'s <c>BuildAnchorIndex</c>; the comment tracker must record
/// per-block char-range targets.
/// </summary>
public class IrScopeTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    // --- header / footer scope naming + projection parity -----------------

    [Fact]
    public void HeaderFooter_ScopeNames_AreHdr1Ftr1()
    {
        var ir = IrReader.Read(IrTestDocuments.WithHeaderAndFooter("My header", "My footer"));

        var header = Assert.Single(ir.Headers);
        Assert.Equal("hdr1", header.ScopeName);
        Assert.Equal(IrHeaderFooterKind.Default, header.Kind);
        Assert.All(header.Scope.Blocks, b => Assert.Equal("hdr1", b.Anchor.Scope));

        var footer = Assert.Single(ir.Footers);
        Assert.Equal("ftr1", footer.ScopeName);
        Assert.All(footer.Scope.Blocks, b => Assert.Equal("ftr1", b.Anchor.Scope));
    }

    [Fact]
    public void HeaderFooter_AnchorsMatch_MarkdownProjection()
    {
        // Parity on a real fixture with multiple header/footer parts: the set of IR anchors per
        // hdr*/ftr* scope must equal the markdown projection's AnchorIndex anchors for that scope.
        var doc = new WmlDocument(Path.Combine(TestFilesDir.FullName, "DB002-Sections-With-Headers.docx"));

        var ir = IrReader.Read(doc);

        var projection = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings
        {
            Scopes = ProjectionScopes.All,
            AnchorIdRendering = AnchorIdRendering.FullUnid,
            // AnchorOnly keeps empty paragraphs in the index so it matches the IR (which indexes all
            // blocks); only Suppress would drop them.
            EmptyParagraphs = EmptyParagraphMode.AnchorOnly,
        });

        // Projection anchor ids for hdr*/ftr* scopes (canonical full-unid form).
        var projectionByScope = projection.AnchorIndex.Values
            .Select(t => t.Anchor)
            .Where(a => a.Scope.StartsWith("hdr") || a.Scope.StartsWith("ftr"))
            .GroupBy(a => a.Scope)
            .ToDictionary(g => g.Key, g => g.Select(a => a.Id).ToHashSet());

        // IR anchor ids for the same scopes — paragraph/heading/list/table-ish block anchors.
        var irByScope = ir.Headers.Concat(ir.Footers)
            .ToDictionary(
                hf => hf.ScopeName,
                hf => CollectBlockAnchors(hf.Scope.Blocks).ToHashSet());

        Assert.NotEmpty(irByScope);

        foreach (var (scope, irAnchors) in irByScope)
        {
            Assert.True(projectionByScope.ContainsKey(scope),
                $"Projection has no anchors for IR scope '{scope}'.");
            // Every IR block anchor must be an anchor the projection also produced (the projection may
            // additionally index sub-block elements the IR does not surface as blocks, so compare ⊆).
            foreach (var a in irAnchors)
                Assert.Contains(a, projectionByScope[scope]);
        }
    }

    private static IEnumerable<string> CollectBlockAnchors(IEnumerable<IrBlock> blocks)
    {
        foreach (var b in blocks)
        {
            // A block SDT is internal IR structure: the markdown projection indexes its descendants,
            // not the wrapper anchor itself.
            if (b is IrSdtBlock sdt)
            {
                foreach (var a in CollectBlockAnchors(sdt.Blocks))
                    yield return a;
            }
            else
            {
                yield return b.Anchor.ToString();
            }
            if (b is IrTable t)
                foreach (var row in t.Rows)
                    foreach (var cell in row.Cells)
                        foreach (var a in CollectBlockAnchors(cell.Blocks))
                            yield return a;
        }
    }

    // --- note store -------------------------------------------------------

    [Fact]
    public void Notes_RealNoteRead_BoilerplateSkipped_ScopeTagged()
    {
        var ir = IrReader.Read(IrTestDocuments.WithFootnoteAndEndnote("Footnote body.", "Endnote body."));

        // Only the real note (id=1) survives; the separator/continuationSeparator notes are skipped.
        Assert.Equal(new[] { "1" }, ir.Footnotes.Notes.Keys.OrderBy(k => k));
        Assert.Equal(new[] { "1" }, ir.Endnotes.Notes.Keys.OrderBy(k => k));

        var fnPara = Assert.IsType<IrParagraph>(Assert.Single(ir.Footnotes.Notes["1"].Blocks));
        Assert.Equal("fn", fnPara.Anchor.Scope);
        Assert.Contains(fnPara.Inlines.OfType<IrTextRun>(), r => r.Text == "Footnote body.");

        var enPara = Assert.IsType<IrParagraph>(Assert.Single(ir.Endnotes.Notes["1"].Blocks));
        Assert.Equal("en", enPara.Anchor.Scope);
        Assert.Contains(enPara.Inlines.OfType<IrTextRun>(), r => r.Text == "Endnote body.");
    }

    [Fact]
    public void Notes_ScopeFlagOff_EmptyStores()
    {
        var doc = IrTestDocuments.WithFootnoteAndEndnote("Footnote body.", "Endnote body.");

        var ir = IrReader.Read(doc, new IrReaderOptions { Scopes = IrScopes.Body });

        Assert.Empty(ir.Footnotes.Notes);
        Assert.Empty(ir.Endnotes.Notes);
    }

    // --- comment store + metadata -----------------------------------------

    [Fact]
    public void Comments_MetadataAndBlocks_Read()
    {
        var body =
            "<w:p>" +
            "<w:commentRangeStart w:id=\"0\"/>" +
            "<w:r><w:t xml:space=\"preserve\">Hello</w:t></w:r>" +
            "<w:commentRangeEnd w:id=\"0\"/>" +
            "<w:r><w:commentReference w:id=\"0\"/></w:r>" +
            "</w:p>";
        var ir = IrReader.Read(IrTestDocuments.WithComment(
            "Eric White", "EW", "2014-10-28T20:22:00Z", "Nice point.", body));

        var comment = Assert.Single(ir.Comments.Comments);
        Assert.Equal("cmt", comment.Anchor.Scope);
        Assert.Equal(IrAnchorKind.Cmt, comment.Anchor.Kind);
        Assert.Equal("Eric White", comment.Author);
        Assert.Equal("EW", comment.Initials);
        Assert.Equal("2014-10-28T20:22:00Z", comment.Date);

        var para = Assert.IsType<IrParagraph>(Assert.Single(comment.Blocks));
        Assert.Equal("cmt", para.Anchor.Scope);
        Assert.Contains(para.Inlines.OfType<IrTextRun>(), r => r.Text == "Nice point.");
    }

    // --- N15 comment targets ----------------------------------------------

    [Fact]
    public void CommentTarget_SingleBlockRange_HasCorrectOffsets()
    {
        // "Hello" then the range covers " world" (offsets 5..11).
        var body =
            "<w:p>" +
            "<w:r><w:t xml:space=\"preserve\">Hello</w:t></w:r>" +
            "<w:commentRangeStart w:id=\"0\"/>" +
            "<w:r><w:t xml:space=\"preserve\"> world</w:t></w:r>" +
            "<w:commentRangeEnd w:id=\"0\"/>" +
            "<w:r><w:commentReference w:id=\"0\"/></w:r>" +
            "</w:p>";
        var ir = IrReader.Read(IrTestDocuments.WithComment("A", "A", "2020-01-01T00:00:00Z", "c", body));

        var comment = Assert.Single(ir.Comments.Comments);
        var target = Assert.Single(comment.Targets);
        Assert.Equal(5, target.StartChar);
        Assert.Equal(11, target.EndChar);
        // The target points at the body paragraph that holds the range.
        var bodyPara = Assert.IsType<IrParagraph>(Assert.Single(ir.Body.Blocks));
        Assert.Equal(bodyPara.Anchor, target.BlockAnchor);
    }

    [Fact]
    public void CommentTarget_CrossBlockRange_OneTargetPerBlock()
    {
        // The range opens mid first paragraph and closes mid second paragraph: two targets, one per
        // touched block. First: offset 5 → end (10, "Hello more"=10). Second: 0 → 4 ("Next").
        var body =
            "<w:p>" +
            "<w:r><w:t xml:space=\"preserve\">Hello</w:t></w:r>" +
            "<w:commentRangeStart w:id=\"0\"/>" +
            "<w:r><w:t xml:space=\"preserve\"> more</w:t></w:r>" +
            "</w:p>" +
            "<w:p>" +
            "<w:r><w:t xml:space=\"preserve\">Next</w:t></w:r>" +
            "<w:commentRangeEnd w:id=\"0\"/>" +
            "<w:r><w:commentReference w:id=\"0\"/></w:r>" +
            "</w:p>";
        var ir = IrReader.Read(IrTestDocuments.WithComment("A", "A", "2020-01-01T00:00:00Z", "c", body));

        var comment = Assert.Single(ir.Comments.Comments);
        Assert.Equal(2, comment.Targets.Count);

        var blocks = ir.Body.Blocks.OfType<IrParagraph>().ToList();
        var first = comment.Targets[0];
        Assert.Equal(blocks[0].Anchor, first.BlockAnchor);
        Assert.Equal(5, first.StartChar);   // after "Hello"
        Assert.Equal(10, first.EndChar);    // end of "Hello more"

        var second = comment.Targets[1];
        Assert.Equal(blocks[1].Anchor, second.BlockAnchor);
        Assert.Equal(0, second.StartChar);  // re-opened at offset 0 of the next block
        Assert.Equal(4, second.EndChar);    // end of "Next"
    }

    [Fact]
    public void CommentTarget_OrphanRangeStart_DiscardedNoThrow()
    {
        // A range-start with no matching end (and no reference) — the comment ends up with no target.
        var body =
            "<w:p>" +
            "<w:r><w:t xml:space=\"preserve\">Hello</w:t></w:r>" +
            "<w:commentRangeStart w:id=\"0\"/>" +
            "<w:r><w:t xml:space=\"preserve\"> world</w:t></w:r>" +
            "</w:p>";
        var ir = IrReader.Read(IrTestDocuments.WithComment("A", "A", "2020-01-01T00:00:00Z", "c", body));

        var comment = Assert.Single(ir.Comments.Comments);
        // An orphan range-start (no matching end, no reference) is discarded silently: its provisional
        // per-block spans are never committed. Totality holds (no throw) and the comment has no target.
        Assert.Empty(comment.Targets);
    }

    [Fact]
    public void CommentTarget_ReferenceWithNoRange_ZeroLengthTarget()
    {
        // Only a commentReference, no range markers: a zero-length target at the reference offset.
        var body =
            "<w:p>" +
            "<w:r><w:t xml:space=\"preserve\">Hello</w:t></w:r>" +
            "<w:r><w:commentReference w:id=\"0\"/></w:r>" +
            "</w:p>";
        var ir = IrReader.Read(IrTestDocuments.WithComment("A", "A", "2020-01-01T00:00:00Z", "c", body));

        var comment = Assert.Single(ir.Comments.Comments);
        var target = Assert.Single(comment.Targets);
        Assert.Equal(target.StartChar, target.EndChar); // zero-length
        Assert.Equal(5, target.StartChar);              // after "Hello"
    }

    [Fact]
    public void Comments_ScopeFlagOff_EmptyStore()
    {
        var body =
            "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t>X</w:t></w:r>" +
            "<w:commentRangeEnd w:id=\"0\"/></w:p>";
        var doc = IrTestDocuments.WithComment("A", "A", "2020-01-01T00:00:00Z", "c", body);

        var ir = IrReader.Read(doc, new IrReaderOptions { Scopes = IrScopes.Body });

        Assert.Empty(ir.Comments.Comments);
    }

    // --- scope flag filtering ---------------------------------------------

    [Fact]
    public void Scopes_BodyOnly_HeadersAndFootersEmpty()
    {
        var doc = IrTestDocuments.WithHeaderAndFooter("H", "F");

        var ir = IrReader.Read(doc, new IrReaderOptions { Scopes = IrScopes.Body });

        Assert.Empty(ir.Headers);
        Assert.Empty(ir.Footers);
        Assert.NotEmpty(ir.Body.Blocks);
    }
}
