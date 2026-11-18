using System.Linq;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

/// <summary>
/// The edit script pairs paragraphs on their modeled properties, so a paragraph it calls unchanged can still end a
/// section whose unmodeled properties changed. The semantic change set reports that change on a section break inside
/// the body as it does on the document's final section, and only that change (issue #1785).
/// </summary>
public class SemanticDiffUnmodeledPropertyTests
{
    private const string Page = "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>";

    private static SemanticChangeFamily[] Families(string left, string right) =>
        SemanticDiff.Compare(IrTestDocuments.FromBodyXml(left), IrTestDocuments.FromBodyXml(right)).Changes
            .Select(change => change.Family).Distinct().ToArray();

    private static string InlineSection(string properties, string mark = "") =>
        $"<w:p><w:pPr>{mark}<w:sectPr>{properties}</w:sectPr></w:pPr><w:r><w:t>First section.</w:t></w:r></w:p>" +
        $"<w:p><w:r><w:t>Second.</w:t></w:r></w:p><w:sectPr>{Page}</w:sectPr>";

    private static string FinalSection(string properties) =>
        $"<w:p><w:r><w:t>Only section.</w:t></w:r></w:p><w:sectPr>{properties}</w:sectPr>";

    [Theory]
    [InlineData(Page + "<w:titlePg/>", Page)]
    [InlineData(Page, Page + "<w:cols w:num=\"2\"/>")]
    public void An_unmodeled_section_change_is_reported_on_an_inline_break_as_on_the_final_section(string left, string right)
    {
        Assert.Equal(new[] { SemanticChangeFamily.Section }, Families(FinalSection(left), FinalSection(right)));
        Assert.Equal(new[] { SemanticChangeFamily.Section }, Families(InlineSection(left), InlineSection(right)));
    }

    [Theory]
    // Indentation inside the section's properties.
    [InlineData(Page + "<w:cols w:num=\"2\" w:equalWidth=\"0\"><w:col w:w=\"4000\"/><w:col w:w=\"4000\"/></w:cols>",
        Page + "<w:cols w:num=\"2\" w:equalWidth=\"0\">\n  <w:col w:w=\"4000\"/>\n  <w:col w:w=\"4000\"/>\n</w:cols>", "")]
    // A default stated on one side only.
    [InlineData(Page + "<w:titlePg w:val=\"0\"/>", Page, "")]
    // A paragraph-mark language, which the default comparison policy does not model.
    [InlineData(Page, Page, "<w:rPr><w:lang w:val=\"nb-NO\"/></w:rPr>")]
    public void An_unchanged_inline_section_break_reports_nothing(string left, string right, string mark)
    {
        // Control: none of these is a section or paragraph change.
        Assert.Empty(Families(InlineSection(left, mark), InlineSection(right)));
    }
}
