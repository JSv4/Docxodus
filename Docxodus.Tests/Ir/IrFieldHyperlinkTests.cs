#nullable enable

using System.Linq;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

/// <summary>
/// M1.2 Task 2 tests: N9 (fields — <c>w:fldSimple</c> + complex <c>w:fldChar</c> machinery) and
/// N14 (<c>w:hyperlink</c>) promotion to typed inlines, plus the field/hyperlink content-hash and
/// format-fingerprint semantics from spec §6.1. Also covers two follow-up <c>w:sym</c> cases
/// (control-char rejection, non-coalescing with plain text).
/// </summary>
public class IrFieldHyperlinkTests
{
    private static IrParagraph Para(string bodyXml) =>
        IrReader.Read(IrTestDocuments.FromBodyXml(bodyXml))
            .Body.Blocks.OfType<IrParagraph>().Single();

    private static IrParagraph ParaWithRels(string bodyXml, params (string, string)[] rels) =>
        IrReader.Read(IrTestDocuments.FromBodyXmlWithHyperlinks(bodyXml, rels))
            .Body.Blocks.OfType<IrParagraph>().Single();

    private static string TextOf(System.Collections.Generic.IEnumerable<IrInline> inlines) =>
        string.Concat(inlines.OfType<IrTextRun>().Select(r => r.Text));

    // --- w:sym follow-ups -------------------------------------------------

    [Fact]
    public void Read_Sym_ControlChar_StaysOpaque()
    {
        // w:char="0000" is a C0 control (U+0000): not legal XML text and a sentinel-collision
        // hazard, so it must fall back to opaque rather than the text path.
        var p = Para("<w:p><w:r><w:sym w:font=\"Wingdings\" w:char=\"0000\"/></w:r></w:p>");

        Assert.Empty(p.Inlines.OfType<IrTextRun>());
        var opaque = Assert.Single(p.Inlines.OfType<IrOpaqueInline>());
        Assert.Equal("sym", opaque.ElementName.LocalName);
    }

    [Fact]
    public void Read_Sym_DoesNotCoalesceWithPlainText()
    {
        // A plain text run followed by a w:sym in the SAME run: the sym carries a different format
        // digest (the glyph-bearing w:sym folds into its run format), so N5 coalescing must not
        // merge them — the sym stays a separate IrTextRun.
        var p = Para("<w:p><w:r><w:t>A</w:t><w:sym w:font=\"Wingdings\" w:char=\"F041\"/></w:r></w:p>");

        var runs = p.Inlines.OfType<IrTextRun>().ToList();
        Assert.Equal(2, runs.Count);
        Assert.Equal("A", runs[0].Text);
        Assert.Equal("", runs[1].Text);
        Assert.NotEqual(runs[0].Format, runs[1].Format);
    }

    // --- N9: fields -------------------------------------------------------

    [Fact]
    public void Read_FldSimple_BecomesFieldRun()
    {
        var p = Para(
            "<w:p><w:fldSimple w:instr=\" PAGE \">" +
            "<w:r><w:t>5</w:t></w:r></w:fldSimple></w:p>");

        var field = Assert.Single(p.Inlines.OfType<IrFieldRun>());
        Assert.Equal(" PAGE ", field.Instruction);
        Assert.Equal("5", TextOf(field.CachedResult));
    }

    [Fact]
    public void Read_ComplexField_BecomesFieldRun()
    {
        var p = Para(
            "<w:p>" +
            "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:r><w:instrText> PAGE </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
            "<w:r><w:t>5</w:t></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
            "</w:p>");

        var field = Assert.Single(p.Inlines.OfType<IrFieldRun>());
        Assert.Equal(" PAGE ", field.Instruction);
        Assert.Equal("5", TextOf(field.CachedResult));
        // No stray fldChar/instrText leaked into the inline stream as opaque.
        Assert.DoesNotContain(p.Inlines, i => i is IrOpaqueInline);
    }

    [Fact]
    public void Read_FieldResult_ContentEqualsLiteralText()
    {
        // A field whose cached result reads "5" is content-equal to a literal "5" (the instruction
        // is unhashed). Fingerprints may differ; ContentHash must match.
        var field = Para(
            "<w:p>" +
            "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:r><w:instrText> PAGE </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
            "<w:r><w:t>5</w:t></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
            "</w:p>");
        var literal = Para("<w:p><w:r><w:t>5</w:t></w:r></w:p>");

        Assert.Equal(literal.ContentHash.ToHex(), field.ContentHash.ToHex());
    }

    [Fact]
    public void Read_InstructionOnlyField_HasEmptyResult()
    {
        // A field with no separate (e.g. some TOC fields) → CachedResult empty, no throw.
        var p = Para(
            "<w:p>" +
            "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:r><w:instrText> TOC \\o </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
            "</w:p>");

        var field = Assert.Single(p.Inlines.OfType<IrFieldRun>());
        Assert.Equal(" TOC \\o ", field.Instruction);
        Assert.Empty(field.CachedResult);
    }

    [Fact]
    public void Read_UnterminatedField_FallsBackToOpaque()
    {
        // begin without a matching end by paragraph close → no IrFieldRun; the captured elements
        // are preserved as opaque inlines so nothing is lost, and it does not throw.
        var p = Para(
            "<w:p>" +
            "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:r><w:instrText> PAGE </w:instrText></w:r>" +
            "<w:r><w:t>partial</w:t></w:r>" +
            "</w:p>");

        Assert.Empty(p.Inlines.OfType<IrFieldRun>());
        Assert.NotEmpty(p.Inlines.OfType<IrOpaqueInline>());
    }

    // --- N14: hyperlinks --------------------------------------------------

    [Fact]
    public void Read_Hyperlink_BecomesTypedInline()
    {
        var p = ParaWithRels(
            "<w:p><w:hyperlink r:id=\"rId99\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>click here</w:t></w:r></w:hyperlink></w:p>",
            ("rId99", "https://example.com/"));

        var link = Assert.Single(p.Inlines.OfType<IrHyperlink>());
        Assert.Equal("https://example.com/", link.Target);
        Assert.Null(link.InternalTarget);
        Assert.Equal("click here", TextOf(link.Inlines));
    }

    [Fact]
    public void Read_Hyperlink_MissingRelation_TargetNull()
    {
        // r:id with no matching relationship tolerates to Target=null (no throw).
        var p = Para(
            "<w:p><w:hyperlink r:id=\"rIdMissing\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>dangling</w:t></w:r></w:hyperlink></w:p>");

        var link = Assert.Single(p.Inlines.OfType<IrHyperlink>());
        Assert.Null(link.Target);
        Assert.Equal("dangling", TextOf(link.Inlines));
    }

    [Fact]
    public void Read_Hyperlink_TargetParticipatesInContentHash()
    {
        var a = ParaWithRels(
            "<w:p><w:hyperlink r:id=\"rId1\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>same</w:t></w:r></w:hyperlink></w:p>",
            ("rId1", "https://a.example/"));
        var b = ParaWithRels(
            "<w:p><w:hyperlink r:id=\"rId1\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>same</w:t></w:r></w:hyperlink></w:p>",
            ("rId1", "https://b.example/"));

        Assert.NotEqual(a.ContentHash.ToHex(), b.ContentHash.ToHex());
    }

    [Fact]
    public void Read_Hyperlink_NotEqualToPlainText()
    {
        var linked = ParaWithRels(
            "<w:p><w:hyperlink r:id=\"rId1\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>foo</w:t></w:r></w:hyperlink></w:p>",
            ("rId1", "https://example.com/"));
        var plain = Para("<w:p><w:r><w:t>foo</w:t></w:r></w:p>");

        Assert.NotEqual(plain.ContentHash.ToHex(), linked.ContentHash.ToHex());
    }

    [Fact]
    public void Read_InternalHyperlink_AnchorConvention()
    {
        var p = Para(
            "<w:p><w:hyperlink w:anchor=\"bm1\">" +
            "<w:r><w:t>go</w:t></w:r></w:hyperlink></w:p>");

        var link = Assert.Single(p.Inlines.OfType<IrHyperlink>());
        Assert.Equal("#bm1", link.Target);
        Assert.Null(link.InternalTarget);
    }

    [Fact]
    public void Read_HyperlinkRuns_ParticipateInFingerprint()
    {
        // Same target, same display text, but bold vs plain link text → different FormatFingerprint
        // (the link's run formats participate in the paragraph's run-format sequence in order).
        var bold = ParaWithRels(
            "<w:p><w:hyperlink r:id=\"rId1\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:rPr><w:b/></w:rPr><w:t>link</w:t></w:r></w:hyperlink></w:p>",
            ("rId1", "https://example.com/"));
        var plain = ParaWithRels(
            "<w:p><w:hyperlink r:id=\"rId1\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>",
            ("rId1", "https://example.com/"));

        Assert.Equal(bold.ContentHash.ToHex(), plain.ContentHash.ToHex());
        Assert.NotEqual(bold.FormatFingerprint.ToHex(), plain.FormatFingerprint.ToHex());
    }
}
