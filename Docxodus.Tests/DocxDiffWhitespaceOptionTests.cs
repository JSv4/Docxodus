#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace OxPt;

/// <summary>
/// Word Compare's <c>White space</c> option on the <see cref="DocxDiff"/> engine.
/// </summary>
public class DocxDiffWhitespaceOptionTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static WmlDocument Doc(params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                paragraphs.Select(text => new Paragraph(new Run(
                    new Text(text) { Space = SpaceProcessingModeValues.Preserve })))));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("test.docx", stream.ToArray());
    }

    private static DocxDiffSettings Off => new() { CompareWhitespace = false };

    private static void AssertValid(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        Assert.Empty(new OpenXmlValidator().Validate(wDoc).Select(e => e.Description));
    }

    /// <summary>
    /// Every character the body carries, deleted text included — <c>w:delText</c> must be in scope or a
    /// deleted double space hides from the assertions and they hold with the option on as well.
    /// </summary>
    private static string TextOf(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        using var partStream = wDoc.MainDocumentPart!.GetStream();
        return string.Concat(XDocument.Load(partStream).Root!
            .Element(W + "body")!
            .Descendants()
            .Where(e => e.Name == W + "t" || e.Name == W + "delText")
            .Select(t => t.Value));
    }

    private static string TextOf(WmlDocument doc) => TextOf(doc.DocumentByteArray);

    [Fact]
    public void CompareWhitespace_defaults_true()
    {
        Assert.True(new DocxDiffSettings().CompareWhitespace);
    }

    [Fact]
    public void Default_reports_a_doubled_space()
    {
        Assert.NotEmpty(DocxDiff.GetRevisions(Doc("Hello world."), Doc("Hello  world.")));
    }

    [Fact]
    public void Off_ignores_a_doubled_space()
    {
        Assert.Empty(DocxDiff.GetRevisions(Doc("Hello world."), Doc("Hello  world."), Off));
    }

    [Fact]
    public void Off_ignores_paragraph_edge_whitespace()
    {
        Assert.Empty(DocxDiff.GetRevisions(Doc("Hello world."), Doc("  Hello world.  "), Off));
    }

    [Fact]
    public void Off_still_reports_a_real_text_change()
    {
        Assert.NotEmpty(DocxDiff.GetRevisions(Doc("Hello  world."), Doc("Hello  there."), Off));
    }

    [Fact]
    public void Off_still_reports_an_inserted_word_between_spaces()
    {
        var revisions = DocxDiff.GetRevisions(Doc("Hello world."), Doc("Hello  big  world."), Off);
        Assert.Contains(revisions, r =>
            r.Type == DocxDiffRevisionType.Inserted && r.Text.Contains("big"));
    }

    [Fact]
    public void Off_produces_no_edit_script_operations_for_a_whitespace_only_change()
    {
        var json = DocxDiff.GetEditScriptJson(Doc("First   paragraph."), Doc("First paragraph. "), Off);

        using var parsed = JsonDocument.Parse(json);
        Assert.All(
            parsed.RootElement.GetProperty("operations").EnumerateArray(),
            op => Assert.Equal("EqualBlock", op.GetProperty("kind").GetString()));
    }

    [Fact]
    public void Off_output_is_valid_and_carries_canonical_whitespace()
    {
        var compared = DocxDiff.Compare(Doc("Hello   world. "), Doc("Hello world."), Off);
        AssertValid(compared);
        Assert.Equal("Hello world.", TextOf(compared));
    }

    [Fact]
    public void Off_round_trips_against_the_canonicalized_inputs()
    {
        // Both sides carry non-canonical spacing so neither half of the round trip is vacuous.
        var left = Doc("Hello   world.", "Second  line.");
        var right = Doc("Hello  world.", "Second   line changed. ");

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);

        var canonicalLeft = TextOf(WhitespaceCanonicalizer.Canonicalize(left));
        var canonicalRight = TextOf(WhitespaceCanonicalizer.Canonicalize(right));

        Assert.Equal(canonicalRight, TextOf(DocxDiffOps.AcceptRevisions(compared.DocumentByteArray)));
        Assert.Equal(canonicalLeft, TextOf(DocxDiffOps.RejectRevisions(compared.DocumentByteArray)));
    }

    [Fact]
    public void Off_applies_to_Consolidate()
    {
        var reviewers = new List<DocxDiffReviewer>
        {
            new() { Author = "reviewer", Document = Doc("Hello world.") },
        };
        var settings = new DocxDiffConsolidateSettings { Diff = Off };

        var consolidated = DocxDiff.Consolidate(Doc("Hello   world."), reviewers, settings);
        AssertValid(consolidated);
        Assert.Empty(DocxDiff.GetConsolidatedRevisions(Doc("Hello   world."), reviewers, settings));
        Assert.Equal("Hello world.", TextOf(consolidated));
    }

    [Fact]
    public void DocxCompare_maps_CompareWhitespace_onto_the_DocxDiff_settings()
    {
        var mapped = DocxCompare.ToDocxDiffSettings(new WmlComparerSettings { CompareWhitespace = false });
        Assert.False(mapped.CompareWhitespace);
        Assert.True(DocxCompare.ToDocxDiffSettings(new WmlComparerSettings()).CompareWhitespace);
    }

    [Fact]
    public void Engine_selector_honors_CompareWhitespace_on_the_DocxDiff_branch()
    {
        var settings = new WmlComparerSettings { CompareWhitespace = false };
        var compared = DocxCompare.Compare(
            Doc("Hello   world."), Doc("Hello world."), ComparisonEngine.DocxDiff, settings);

        AssertValid(compared);
        Assert.Empty(WmlComparer.GetRevisions(compared, settings));
    }
}
