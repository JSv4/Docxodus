#nullable enable

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// M-B — the shared comparison-engine selector. <see cref="DocxCompare"/> owns the sole
/// <c>WmlComparer</c>-vs-<c>DocxDiff</c> branch that the CLI / WASM / npm surfaces route through.
/// These tests pin: (1) the <see cref="ComparisonEngine"/> integer contract the byte-level surfaces
/// rely on, (2) that the default engine reproduces <see cref="WmlComparer"/> exactly (no behavior
/// change), (3) that the opt-in <c>DocxDiff</c> branch equals a direct <see cref="DocxDiff.Compare"/>,
/// (4) that the settings map carries the common option set, and (5) the uniform revision-count
/// assumption redline relies on.
/// </summary>
public class DocxCompareTests
{
    private const string FixedDate = "2021-01-01T00:00:00Z";

    // Two paragraphs differing by one word, so either engine yields a real insertion/deletion.
    private static WmlDocument Doc(string text)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }))));
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

    private static string[] RevisionSig(WmlDocument redline, WmlComparerSettings settings) =>
        WmlComparer.GetRevisions(redline, settings)
            .Select(r => $"{r.RevisionType}:{r.Text}")
            .ToArray();

    [Fact]
    public void DefaultEngine_IsWmlComparer()
    {
        Assert.Equal(ComparisonEngine.WmlComparer, default(ComparisonEngine));
        Assert.Equal(ComparisonEngine.WmlComparer, (ComparisonEngine)0);
        Assert.Equal(1, (int)ComparisonEngine.DocxDiff);
    }

    [Fact]
    public void WmlComparerBranch_ProducesSameRevisionsAsDirectWmlComparer()
    {
        var left = Doc("The quick brown fox");
        var right = Doc("The quick red fox");
        var settings = new WmlComparerSettings { DetailThreshold = 0, DateTimeForRevisions = FixedDate };

        var viaFacade = DocxCompare.Compare(left, right, ComparisonEngine.WmlComparer, settings);
        var direct = WmlComparer.Compare(left, right, settings);

        Assert.Equal(RevisionSig(direct, settings), RevisionSig(viaFacade, settings));
        Assert.NotEmpty(WmlComparer.GetRevisions(viaFacade, settings));
    }

    [Fact]
    public void DocxDiffBranch_ProducesSameBytesAsDirectDocxDiff()
    {
        var left = Doc("The quick brown fox");
        var right = Doc("The quick red fox");
        var settings = new WmlComparerSettings { DetailThreshold = 0, DateTimeForRevisions = FixedDate };

        var viaFacade = DocxCompare.Compare(left, right, ComparisonEngine.DocxDiff, settings);
        var direct = DocxDiff.Compare(left, right, DocxCompare.ToDocxDiffSettings(settings));

        Assert.Equal(direct.DocumentByteArray, viaFacade.DocumentByteArray);
    }

    [Fact]
    public void ToDocxDiffSettings_CarriesCommonFields()
    {
        var settings = new WmlComparerSettings
        {
            AuthorForRevisions = "Alice",
            DateTimeForRevisions = "2020-01-02T03:04:05Z",
            CaseInsensitive = true,
            ConflateBreakingAndNonbreakingSpaces = false,
            DetectMoves = false,
            MoveSimilarityThreshold = 0.55,
            MoveMinimumWordCount = 9,
        };

        var mapped = DocxCompare.ToDocxDiffSettings(settings);

        Assert.Equal("Alice", mapped.AuthorForRevisions);
        Assert.Equal("2020-01-02T03:04:05Z", mapped.DateTimeForRevisions);
        Assert.True(mapped.CaseInsensitive);
        Assert.False(mapped.ConflateBreakingAndNonbreakingSpaces);
        Assert.False(mapped.DetectMoves);
        Assert.Equal(0.55, mapped.MoveSimilarityThreshold);
        Assert.Equal(9, mapped.MoveMinimumWordCount);
    }

    [Theory]
    [InlineData("wmlcomparer", ComparisonEngine.WmlComparer)]
    [InlineData("docxdiff", ComparisonEngine.DocxDiff)]
    [InlineData("DocxDiff", ComparisonEngine.DocxDiff)]     // case-insensitive
    [InlineData("  docxdiff  ", ComparisonEngine.DocxDiff)] // trims surrounding whitespace
    public void TryParseEngine_RecognizesKnownNames(string value, ComparisonEngine expected)
    {
        Assert.True(DocxCompare.TryParseEngine(value, out var engine));
        Assert.Equal(expected, engine);
    }

    [Theory]
    [InlineData("bogus")]
    [InlineData("")]
    [InlineData(null)]
    public void TryParseEngine_RejectsUnknown_DefaultsToWmlComparer(string? value)
    {
        Assert.False(DocxCompare.TryParseEngine(value, out var engine));
        Assert.Equal(ComparisonEngine.WmlComparer, engine);
    }

    [Fact]
    public void DocxDiffBranch_OutputIsRevisionCountableViaWmlComparer()
    {
        var left = Doc("The quick brown fox");
        var right = Doc("The quick red fox");
        var settings = new WmlComparerSettings { DetailThreshold = 0, DateTimeForRevisions = FixedDate };

        var output = DocxCompare.Compare(left, right, ComparisonEngine.DocxDiff, settings);

        Assert.True(WmlComparer.GetRevisions(output, settings).Count > 0);
    }

    [Fact]
    public void DocxDiffBranch_PreAcceptsInputRevisions_LikeWmlComparerAndWord()
    {
        // WmlComparer (and Microsoft Word's compare) treat tracked changes in the INPUTS as accepted
        // before comparing — no input revision markup (or its author) survives into the redline body.
        // Engine equivalence through DocxCompare requires the DocxDiff branch to behave identically,
        // so the mapping sets PreAcceptInputRevisions (the raw DocxDiff API default remains opt-in).
        static WmlDocument DocWithTrackedInsertion(string plain, string inserted)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(
                    new Run(new Text(plain) { Space = SpaceProcessingModeValues.Preserve }),
                    new InsertedRun(
                        new Run(new Text(inserted) { Space = SpaceProcessingModeValues.Preserve }))
                    {
                        Author = "PriorReviewer",
                        Id = "99",
                        Date = System.DateTime.Parse("2020-06-01T00:00:00Z",
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.AdjustToUniversal),
                    })));
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                    new ParagraphPropertiesDefault()));
                mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                doc.Save();
            }
            return new WmlDocument("tracked.docx", stream.ToArray());
        }

        var left = DocWithTrackedInsertion("Base text ", "with a prior insertion");
        var right = Doc("Base text with a prior insertion plus fresh words");
        var settings = new WmlComparerSettings { AuthorForRevisions = "Bench", DateTimeForRevisions = FixedDate };

        var output = DocxCompare.Compare(left, right, ComparisonEngine.DocxDiff, settings);

        using var stream = new MemoryStream(output.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var bodyXml = wdoc.MainDocumentPart!.Document.Body!.OuterXml;
        Assert.DoesNotContain("PriorReviewer", bodyXml);
    }
}
