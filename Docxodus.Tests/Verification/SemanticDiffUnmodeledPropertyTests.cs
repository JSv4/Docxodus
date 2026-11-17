using System.Linq;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

/// <summary>
/// The edit script pairs blocks under its own format policy, which by default compares modeled properties only.
/// A paragraph it calls unchanged can still differ in a property it does not model, and the semantic change set
/// reports that the same way whether the section break or paragraph is the document's last or inside the body
/// (issue #1785).
/// </summary>
public class SemanticDiffUnmodeledPropertyTests
{
    private const string Page = "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>";

    private static SemanticChangeFamily[] Families(string left, string right) =>
        SemanticDiff.Compare(IrTestDocuments.FromBodyXml(left), IrTestDocuments.FromBodyXml(right)).Changes
            .Select(change => change.Family).Distinct().ToArray();

    private static string InlineSection(string properties) =>
        $"<w:p><w:pPr><w:sectPr>{properties}</w:sectPr></w:pPr><w:r><w:t>First section.</w:t></w:r></w:p>" +
        $"<w:p><w:r><w:t>Second.</w:t></w:r></w:p><w:sectPr>{Page}</w:sectPr>";

    private static string FinalSection(string properties) =>
        $"<w:p><w:r><w:t>Only section.</w:t></w:r></w:p><w:sectPr>{properties}</w:sectPr>";

    [Theory]
    [InlineData(Page + "<w:titlePg/>", Page)]
    [InlineData(Page, Page + "<w:cols w:num=\"2\"/>")]
    public void An_unmodeled_section_change_is_reported_on_an_inline_break_as_on_the_final_section(string left, string right)
    {
        Assert.Contains(SemanticChangeFamily.Section, Families(FinalSection(left), FinalSection(right)));
        Assert.Contains(SemanticChangeFamily.Section, Families(InlineSection(left), InlineSection(right)));
    }

    [Fact]
    public void An_unmodeled_paragraph_property_change_is_reported()
    {
        const string text = "<w:r><w:t>Widow lines.</w:t></w:r></w:p>";
        Assert.Contains(SemanticChangeFamily.ParagraphFormatting,
            Families("<w:p><w:pPr><w:widowControl w:val=\"0\"/></w:pPr>" + text, "<w:p>" + text));
    }

    [Fact]
    public void An_unchanged_document_split_into_different_runs_reports_nothing()
    {
        // Control: run boundaries alone change the paragraph's stored fingerprint but are no semantic change.
        var families = Families(InlineSection(Page),
            $"<w:p><w:pPr><w:sectPr>{Page}</w:sectPr></w:pPr><w:r><w:t xml:space=\"preserve\">First </w:t></w:r>" +
            $"<w:r><w:t>section.</w:t></w:r></w:p><w:p><w:r><w:t>Second.</w:t></w:r></w:p><w:sectPr>{Page}</w:sectPr>");
        Assert.DoesNotContain(SemanticChangeFamily.Section, families);
        Assert.DoesNotContain(SemanticChangeFamily.ParagraphFormatting, families);
        Assert.DoesNotContain(SemanticChangeFamily.RunFormatting, families);
        Assert.DoesNotContain(SemanticChangeFamily.Text, families);
    }
}
