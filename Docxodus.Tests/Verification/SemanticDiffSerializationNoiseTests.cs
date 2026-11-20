using System.Collections.Generic;
using System.Linq;
using Docxodus.Tests.Ir;
using Docxodus.Tests.Ir.Diff;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

/// <summary>
/// Insignificant whitespace between property elements and the relationship ids a package gives its header parts
/// are not semantic changes (issue #1793).
/// </summary>
public class SemanticDiffSerializationNoiseTests
{
    private const string Page = "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>";
    private const string Columns = "<w:cols w:num=\"2\" w:equalWidth=\"0\"><w:col w:w=\"4000\"/><w:col w:w=\"4000\"/></w:cols>";
    private const string IndentedColumns = "<w:cols w:num=\"2\" w:equalWidth=\"0\">\n  <w:col w:w=\"4000\"/>\n  <w:col w:w=\"4000\"/>\n</w:cols>";

    private static SemanticChangeFamily[] Families(WmlDocument left, WmlDocument right) =>
        SemanticDiff.Compare(left, right, new SemanticDiffOptions { IncludePackageChanges = false }).Changes
            .Select(change => change.Family).Distinct().ToArray();

    private static SemanticChangeFamily[] Families(string left, string right) =>
        Families(IrTestDocuments.FromBodyXml(left), IrTestDocuments.FromBodyXml(right));

    [Theory]
    // The final section's column definitions, indented.
    [InlineData("<w:p><w:r><w:t>Only.</w:t></w:r></w:p><w:sectPr>" + Page + Columns + "</w:sectPr>",
        "<w:p><w:r><w:t>Only.</w:t></w:r></w:p><w:sectPr>" + Page + IndentedColumns + "</w:sectPr>")]
    // A field result whose run properties hold only w:noProof, indented: the noise element and the indentation both go.
    [InlineData("<w:p><w:fldSimple w:instr=\" PAGE \"><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p>",
        "<w:p><w:fldSimple w:instr=\" PAGE \"><w:r><w:rPr>\n  <w:noProof/>\n</w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p>")]
    // A table's borders, indented.
    [InlineData("<w:tbl><w:tblPr><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/><w:bottom w:val=\"single\" w:sz=\"4\"/></w:tblBorders></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell.</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>",
        "<w:tbl><w:tblPr><w:tblBorders>\n  <w:top w:val=\"single\" w:sz=\"4\"/>\n  <w:bottom w:val=\"single\" w:sz=\"4\"/>\n</w:tblBorders></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell.</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>")]
    public void Indentation_between_property_elements_is_no_change(string left, string right) =>
        Assert.Empty(Families(left, right));

    [Fact]
    public void An_indented_style_is_no_change()
    {
        static WmlDocument Styled(string paragraphProperties) => IrTestDocuments.FromBodyAndStylesXml(
            "<w:p><w:pPr><w:pStyle w:val=\"Body\"/></w:pPr><w:r><w:t>Styled.</w:t></w:r></w:p>",
            "<w:style w:type=\"paragraph\" w:styleId=\"Body\"><w:name w:val=\"Body\"/>" + paragraphProperties + "</w:style>");
        Assert.Empty(Families(Styled("<w:pPr><w:spacing w:after=\"120\"/><w:jc w:val=\"both\"/></w:pPr>"),
            Styled("<w:pPr>\n  <w:spacing w:after=\"120\"/>\n  <w:jc w:val=\"both\"/>\n</w:pPr>")));
    }

    [Fact]
    public void A_header_renumbered_on_the_final_section_is_no_change()
    {
        // The same header story under different relationship ids, as independently saved packages number them.
        static WmlDocument Saved(string id) => HeaderFooterFixtures.Build(
            new[] { new HeaderFooterFixtures.Section(new[] { "Harbor pilots log the tide charts." }, Headers: new[] { ("default", id) }) },
            new Dictionary<string, string[]> { [id] = new[] { "Tide table" } });
        Assert.DoesNotContain(SemanticChangeFamily.Section, Families(Saved("rIdHeader1"), Saved("rIdHeaderZ9")));
    }

    [Theory]
    // Controls: real changes are still reported.
    [InlineData("<w:p><w:r><w:t>Only.</w:t></w:r></w:p><w:sectPr>" + Page + Columns + "</w:sectPr>",
        "<w:p><w:r><w:t>Only.</w:t></w:r></w:p><w:sectPr>" + Page + "<w:cols w:num=\"3\"/></w:sectPr>", SemanticChangeFamily.Section)]
    // A space that is the whole text of a run is content, not indentation.
    [InlineData("<w:p><w:r><w:t>Two</w:t></w:r><w:r><w:t xml:space=\"preserve\"> </w:t></w:r><w:r><w:t>words.</w:t></w:r></w:p>",
        "<w:p><w:r><w:t>Two</w:t></w:r><w:r><w:t>words.</w:t></w:r></w:p>", SemanticChangeFamily.Text)]
    public void A_real_change_beside_the_noise_is_still_reported(string left, string right, SemanticChangeFamily family) =>
        Assert.Contains(family, Families(left, right));

    [Fact]
    public void A_header_reference_changing_kind_is_still_reported()
    {
        // Control: references left the section format, but their stories are still compared as stories.
        static WmlDocument Saved(string kind) => HeaderFooterFixtures.Build(
            new[] { new HeaderFooterFixtures.Section(new[] { "Harbor pilots log the tide charts." }, Headers: new[] { (kind, "rIdHeader1") }) },
            new Dictionary<string, string[]> { ["rIdHeader1"] = new[] { "Tide table" } }, titlePg: true);
        Assert.Contains(SemanticChangeFamily.Header, Families(Saved("default"), Saved("first")));
    }

    [Theory]
    // Whitespace that is content in foreign XML: under xml:space="preserve", and beside other text.
    [InlineData("<v:note xmlns:v=\"urn:vendor\" xml:space=\"preserve\"><v:a>Net</v:a> <v:b>30</v:b></v:note>",
        "<v:note xmlns:v=\"urn:vendor\" xml:space=\"preserve\"><v:a>Net</v:a><v:b>30</v:b></v:note>")]
    [InlineData("<v:note xmlns:v=\"urn:vendor\">Net<!--c--> <v:b>30</v:b></v:note>",
        "<v:note xmlns:v=\"urn:vendor\">Net<!--c--><v:b>30</v:b></v:note>")]
    public void Whitespace_that_is_content_in_foreign_xml_is_still_a_change(string left, string right) =>
        // Control: only indentation between child elements is dropped.
        Assert.NotEmpty(Families($"<w:p><w:r><w:t>Before.</w:t></w:r>{left}</w:p>", $"<w:p><w:r><w:t>Before.</w:t></w:r>{right}</w:p>"));

    [Fact]
    public void Hashing_an_indented_element_with_many_children_is_linear()
    {
        // One pass per element: 16,000 indented children hash in a moment (a per-node scan took tens of seconds).
        static System.Xml.Linq.XElement Content(string indentation)
        {
            System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var content = new System.Xml.Linq.XElement(w + "sdtContent");
            for (int i = 0; i < 16_000; i++)
            {
                if (indentation.Length > 0) content.Add(new System.Xml.Linq.XText(indentation));
                content.Add(new System.Xml.Linq.XElement(w + "p", new System.Xml.Linq.XElement(w + "r", new System.Xml.Linq.XElement(w + "t", "Line " + i))));
            }
            return content;
        }
        var compact = Docxodus.Ir.IrHasher.CanonicalHash(Content(""));
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var indented = Docxodus.Ir.IrHasher.CanonicalHash(Content("\n  "));
        watch.Stop();
        Assert.Equal(compact, indented);
        Assert.True(watch.Elapsed < System.TimeSpan.FromSeconds(5), $"hashing took {watch.Elapsed}");
    }
}
