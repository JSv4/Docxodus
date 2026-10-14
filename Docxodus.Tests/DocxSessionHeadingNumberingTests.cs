#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// #572 — markdown-authored headings must not carry a `numId=0` numbering
/// suppressor the target style never needed. The suppressor's purpose is to
/// stop a legal-outline-numbered Heading style from prefixing the inserted
/// text; written unconditionally it made every markdown heading diff as
/// FormatChanged(numId, numLevel) against an identical Style-dropdown heading.
/// The rule: write the suppressor only when the resolved style (or its
/// basedOn chain) actually attaches numbering — and apply the same rule on
/// both authoring paths (InsertParagraph and ReplaceText's declared-style
/// payloads), so all roads produce the same paragraph mark.
/// </summary>
public class DocxSessionHeadingNumberingTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>One paragraph; Heading2 attaches numbering (numId 5), directly or
    /// through a basedOn chain.</summary>
    private static byte[] BuildDocWithNumberedHeading2(bool viaBasedOn)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First paragraph.")))));

            var numbered = new StyleParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 5 }));

            Style heading2;
            var styles = new Styles(
                new Style(new StyleName { Val = "Normal" })
                { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });
            if (viaBasedOn)
            {
                styles.Append(new Style(new StyleName { Val = "Numbered Base" }, numbered)
                { Type = StyleValues.Paragraph, StyleId = "NumberedBase" });
                heading2 = new Style(
                    new StyleName { Val = "heading 2" },
                    new BasedOn { Val = "NumberedBase" })
                { Type = StyleValues.Paragraph, StyleId = "Heading2" };
            }
            else
            {
                heading2 = new Style(new StyleName { Val = "heading 2" }, numbered)
                { Type = StyleValues.Paragraph, StyleId = "Heading2" };
            }
            styles.Append(heading2);
            main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static XElement PPrOfFirstHeading(DocxSession session)
    {
        var anchor = session.FindByKind("h", "body").Single().Anchor;
        var p = XElement.Parse(session.Raw.GetXml(anchor.Id));
        var pPr = p.Element(W + "pPr") ?? new XElement(W + "pPr");
        // Strip PowerTools bookkeeping (Unids differ per document) so the
        // comparison sees only the OOXML the paragraph mark actually carries.
        var clean = new XElement(pPr);
        foreach (var el in clean.DescendantsAndSelf())
            el.Attributes().Where(a => a.Name.Namespace != W && a.Name.Namespace != XNamespace.None
                || a.IsNamespaceDeclaration).Remove();
        return clean;
    }

    [Fact]
    public void DS572a_InsertParagraph_Heading_NoSuppressorWhenStyleHasNoNumbering()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();

        var r = s.InsertParagraph(anchor, Position.After, "## Duties");
        Assert.True(r.Success, r.Error?.Message);

        var pPr = PPrOfFirstHeading(s);
        Assert.Null(pPr.Element(W + "numPr"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DS572b_InsertParagraph_Heading_KeepsSuppressorWhenStyleNumbers(bool viaBasedOn)
    {
        using var s = new DocxSession(BuildDocWithNumberedHeading2(viaBasedOn));
        var anchor = s.Project().AnchorIndex.Keys.First();

        var r = s.InsertParagraph(anchor, Position.After, "## Duties");
        Assert.True(r.Success, r.Error?.Message);

        var numPr = PPrOfFirstHeading(s).Element(W + "numPr");
        Assert.NotNull(numPr);
        Assert.Equal("0", (string?)numPr!.Element(W + "numId")?.Attribute(W + "val"));
    }

    [Fact]
    public void DS572c_ReplaceText_HeadingPayload_SameSuppressorRule()
    {
        // No numbering on the style → no suppressor…
        using (var plain = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs()))
        {
            var anchor = plain.Project().AnchorIndex.Keys.First();
            var r = plain.ReplaceText(anchor, "## Duties");
            Assert.True(r.Success, r.Error?.Message);
            Assert.Null(PPrOfFirstHeading(plain).Element(W + "numPr"));
        }

        // …numbered style → the suppressor, exactly like InsertParagraph.
        using (var numbered = new DocxSession(BuildDocWithNumberedHeading2(viaBasedOn: false)))
        {
            var anchor = numbered.Project().AnchorIndex.Keys.First();
            var r = numbered.ReplaceText(anchor, "## Duties");
            Assert.True(r.Success, r.Error?.Message);
            var numPr = PPrOfFirstHeading(numbered).Element(W + "numPr");
            Assert.NotNull(numPr);
            Assert.Equal("0", (string?)numPr!.Element(W + "numId")?.Attribute(W + "val"));
        }
    }

    [Fact]
    public void DS572d_MarkdownAndDropdownHeadings_ProduceIdenticalParagraphMarks()
    {
        // The #572 complaint end-to-end: the three authoring paths must agree on
        // the paragraph mark in the common (unnumbered Heading style) case.
        using var viaMarkdownInsert = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var a1 = viaMarkdownInsert.Project().AnchorIndex.Keys.First();
        Assert.True(viaMarkdownInsert.InsertParagraph(a1, Position.After, "## Duties").Success);

        using var viaDropdown = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var a2 = viaDropdown.Project().AnchorIndex.Keys.First();
        var styled = viaDropdown.ReplaceText(a2, "Duties");
        Assert.True(viaDropdown.SetParagraphStyle(styled.Modified![0].Id, "Heading2").Success);

        using var viaReplaceText = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var a3 = viaReplaceText.Project().AnchorIndex.Keys.First();
        Assert.True(viaReplaceText.ReplaceText(a3, "## Duties").Success);

        string Mark(DocxSession s) => PPrOfFirstHeading(s).ToString(SaveOptions.DisableFormatting);
        Assert.Equal(Mark(viaDropdown), Mark(viaMarkdownInsert));
        Assert.Equal(Mark(viaDropdown), Mark(viaReplaceText));
    }
}
