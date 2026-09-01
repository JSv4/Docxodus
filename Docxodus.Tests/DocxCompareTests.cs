#nullable enable

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DocxCompare"/> — the shared front door the CLI / WASM / npm surfaces route their
/// two-document comparison through. Through v10 it also owned the sole
/// <c>WmlComparer</c>-vs-<c>DocxDiff</c> engine branch; with the legacy engine removed in v11.0.0
/// what remains is the POLICY it used to share, which is what these tests pin:
/// <list type="number">
/// <item>byte-identical inputs return a detached exact clone rather than a reserialized package;</item>
/// <item>the front door always pre-accepts AND preserves input revisions, unlike the raw
/// <see cref="DocxDiff"/> API whose flags stay opt-in — the one behavioral difference between the two
/// entry points, and the one most likely to be lost silently;</item>
/// <item>applying that policy never mutates the caller's settings object;</item>
/// <item>the output is revision-countable, which the redline CLI relies on.</item>
/// </list>
/// </summary>
public class DocxCompareTests
{
    private const string FixedDate = "2021-01-01T00:00:00Z";

    // Two paragraphs differing by one word, so a comparison yields a real insertion/deletion.
    private static WmlDocument Doc(string text)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(
                new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }))));
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

    private static string[] SchemaErrors(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDoc = WordprocessingDocument.Open(stream, false);
        return new OpenXmlValidator().Validate(wordDoc).Select(error => error.Description).ToArray();
    }

    [Fact]
    public void FrontDoor_EqualsDocxDiffWithTheFrontDoorPolicyApplied()
    {
        var left = Doc("The quick brown fox");
        var right = Doc("The quick red fox");
        var settings = new DocxDiffSettings { DateTimeForRevisions = FixedDate };

        var viaFacade = DocxCompare.Compare(left, right, settings);
        var direct = DocxDiff.Compare(left, right, DocxCompare.ApplyFrontDoorRevisionPolicy(settings));

        // Part by part, not raw package bytes: two separately produced ZIPs also differ in the entry
        // timestamps the container writes, which says nothing about the facade. See PackageEquivalence.
        PackageEquivalence.AssertSamePackage(direct, viaFacade);
    }

    [Fact]
    public void ByteIdenticalInputs_ReturnDetachedExactClone()
    {
        var source = Doc("Unchanged paragraph.");
        var samePackage = new WmlDocument(source);
        samePackage.FileName = "right.docx";

        var result = DocxCompare.Compare(source, samePackage);

        Assert.NotSame(source, result);
        Assert.NotSame(source.DocumentByteArray, result.DocumentByteArray);
        Assert.Equal(source.FileName, result.FileName);
        Assert.Equal(source.DocumentByteArray, result.DocumentByteArray);
    }

    [Fact]
    public void ByteIdenticalMalformedMathRevision_PassesThroughUnrepaired()
    {
        // WC012 carries tracked-revision wrappers inside an Office Math run — schema-invalid markup that
        // WmlComparer's preprocessing repaired as a side effect, so through v10 CanReturnExactNoOp
        // deliberately REFUSED the exact-clone shortcut for it. DocxDiff performs no such repair (it
        // returns the source bytes unchanged, same validation error), so the guard bought nothing but a
        // full comparison and went with the engine that motivated it. This pins the honest replacement
        // behavior: the shortcut is taken, and invalid input survives rather than being silently
        // rewritten. Repairing it is tracked as issue #642.
        var source = new WmlDocument(Path.GetFullPath(Path.Combine(
            "../../../../TestFiles", "WC", "WC012-Math-After.docx")));
        var samePackage = new WmlDocument(source);

        Assert.NotEmpty(SchemaErrors(source));
        Assert.True(DocxCompare.CanReturnExactNoOp(source, samePackage));

        var result = DocxCompare.Compare(source, samePackage);

        Assert.Equal(source.DocumentByteArray, result.DocumentByteArray);
        Assert.NotEmpty(SchemaErrors(result));
    }

    [Fact]
    public void FrontDoorPolicy_DoesNotMutateTheCallersSettings()
    {
        var settings = new DocxDiffSettings { AuthorForRevisions = "Bench" };

        var applied = DocxCompare.ApplyFrontDoorRevisionPolicy(settings);

        Assert.True(applied.PreAcceptInputRevisions);
        Assert.True(applied.PreserveInputRevisions);
        Assert.Equal("Bench", applied.AuthorForRevisions);

        // The caller's object keeps the raw-engine opt-in defaults.
        Assert.False(settings.PreAcceptInputRevisions);
        Assert.False(settings.PreserveInputRevisions);
    }

    [Fact]
    public void FrontDoorPolicy_AcceptsNullSettings()
    {
        var applied = DocxCompare.ApplyFrontDoorRevisionPolicy(null);

        Assert.True(applied.PreAcceptInputRevisions);
        Assert.True(applied.PreserveInputRevisions);
    }

    [Fact]
    public void Output_IsRevisionCountable()
    {
        var left = Doc("The quick brown fox");
        var right = Doc("The quick red fox");

        var output = DocxCompare.Compare(
            left, right, new DocxDiffSettings { DateTimeForRevisions = FixedDate });

        using var session = new DocxSession(output.DocumentByteArray);
        Assert.NotEmpty(session.ListRevisions());
    }

    [Fact]
    public void FrontDoor_PreAcceptsInputRevisions_LikeWord()
    {
        // Word's compare treats tracked changes in the INPUTS as accepted before comparing — no input
        // revision markup (or its author) survives into the redline body. The front door sets
        // PreAcceptInputRevisions to match; the raw DocxDiff API leaves it opt-in, which is exactly why
        // this assertion lives here and not on DocxDiff.
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

        var output = DocxCompare.Compare(left, right, new DocxDiffSettings
        {
            AuthorForRevisions = "Bench",
            DateTimeForRevisions = FixedDate,
        });

        using var stream = new MemoryStream(output.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var bodyXml = wdoc.MainDocumentPart!.Document.Body!.OuterXml;
        Assert.DoesNotContain("PriorReviewer", bodyXml);
    }
}
