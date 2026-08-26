#nullable enable

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DocxDiff.CreateComparison"/> (issue #594): one alignment pass serving every data
/// product. The contract is (1) every product is identical to what the stateless static
/// produces on the same inputs/settings, and (2) repeated product calls are served from the
/// memoized pass, not recomputed.
/// </summary>
public class DocxDiffComparisonTests
{
    private static WmlDocument Doc(params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                paragraphs.Select(text => new Paragraph(new Run(new Text(text))))));
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

    private static (WmlDocument Left, WmlDocument Right) Pair() => (
        Doc("The quick brown fox.", "Second paragraph here.", "Closing paragraph."),
        Doc("The quick red fox.", "Second paragraph here.", "A fresh closing paragraph."));

    [Fact]
    public void EveryProduct_MatchesItsStatelessStatic()
    {
        var (left, right) = Pair();
        var settings = new DocxDiffSettings { AuthorForRevisions = "Daisy" };
        var comparison = DocxDiff.CreateComparison(left, right, settings);

        Assert.Equal(
            DocxDiff.Compare(left, right, settings).DocumentByteArray,
            comparison.ToRedline().DocumentByteArray);

        var staticRevisions = DocxDiff.GetRevisions(left, right, settings);
        var memoRevisions = comparison.GetRevisions();
        Assert.Equal(staticRevisions.Count, memoRevisions.Count);
        Assert.Equal(
            staticRevisions.Select(r => r.ToString()),
            memoRevisions.Select(r => r.ToString()));

        Assert.Equal(
            DocxDiff.GetEditScriptJson(left, right, settings),
            comparison.GetEditScriptJson());
    }

    [Fact]
    public void SemanticProducts_MatchTheStatics_WithTheComparisonSettingsFlowing()
    {
        var (left, right) = Pair();
        var settings = new DocxDiffSettings { AuthorForRevisions = "Daisy" };
        var comparison = DocxDiff.CreateComparison(left, right, settings);

        // With no explicit options, the comparison's own settings flow into the semantic pass —
        // the same JSON the static produces when handed those settings explicitly.
        var options = new SemanticDiffOptions { DiffSettings = settings };
        Assert.Equal(
            DocxDiff.GetSemanticChangesJson(left, right, options),
            comparison.GetSemanticChangesJson());

        // Explicit options are honored verbatim.
        Assert.Equal(
            DocxDiff.GetSemanticChangesJson(left, right, null),
            comparison.GetSemanticChangesJson(new SemanticDiffOptions()));
    }

    [Fact]
    public void RepeatedCalls_ServeTheMemoizedProduct()
    {
        var (left, right) = Pair();
        var comparison = DocxDiff.CreateComparison(left, right);

        Assert.Same(comparison.GetRevisions(), comparison.GetRevisions());
        Assert.Same(comparison.GetEditScriptJson(), comparison.GetEditScriptJson());
        Assert.Same(comparison.GetSemanticChangesJson(), comparison.GetSemanticChangesJson());
        Assert.Same(
            comparison.ToRedline().DocumentByteArray,
            comparison.ToRedline().DocumentByteArray);
    }

    [Fact]
    public void IdenticalInputs_ShortCircuitLikeTheStatics()
    {
        var doc = Doc("Unchanged text.");
        var same = new WmlDocument("copy.docx", doc.DocumentByteArray);
        var comparison = DocxDiff.CreateComparison(doc, same);

        Assert.Equal(doc.DocumentByteArray, comparison.ToRedline().DocumentByteArray);
        Assert.Empty(comparison.GetRevisions());
    }

    [Fact]
    public void NullInputs_AreRejectedAtCreation()
    {
        var doc = Doc("x");
        Assert.Throws<System.ArgumentNullException>(() => DocxDiff.CreateComparison(null!, doc));
        Assert.Throws<System.ArgumentNullException>(() => DocxDiff.CreateComparison(doc, null!));
    }
}
