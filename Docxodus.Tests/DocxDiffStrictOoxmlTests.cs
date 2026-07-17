#nullable enable

using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// DocxDiff over ISO/IEC 29500 STRICT-conformance inputs (namespace family
/// http://purl.oclc.org/ooxml/*). Word reads strict packages transparently; DocxDiff must
/// too — the engine normalizes strict inputs to transitional before the IR read, so the
/// public contract (accept ≡ right, reject ≡ left) holds regardless of which side (or both)
/// is strict. Regression coverage for the corpus family Strict01/strict01_sdt_controls/
/// ole_object/word_clean_strict01 that previously failed with "Document has no w:body element".
/// </summary>
public class DocxDiffStrictOoxmlTests
{
    private const string TransitionalMain = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StrictMain = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    private const string TransitionalRels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string StrictRels = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    // A minimal programmatic doc, mirroring DocxDiffTests.Doc (all required parts present).
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

    /// <summary>
    /// Rewrite a transitional package into its strict-conformance equivalent by mapping the
    /// wordprocessingml and relationship namespace URIs in every XML part and .rels stream —
    /// the same shape Word's "Strict Open XML Document (*.docx)" save format produces.
    /// </summary>
    private static WmlDocument ToStrict(WmlDocument doc)
    {
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries.ToList())
            {
                if (!entry.FullName.EndsWith(".xml") && !entry.FullName.EndsWith(".rels"))
                    continue;
                string text;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    text = reader.ReadToEnd();
                var rewritten = text
                    .Replace(TransitionalMain, StrictMain)
                    .Replace(TransitionalRels, StrictRels);
                if (ReferenceEquals(rewritten, text) || rewritten == text)
                    continue;
                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.BaseStream.SetLength(0);
                writer.Write(rewritten);
            }
        }
        return new WmlDocument("strict.docx", ms.ToArray());
    }

    private static List<string> BodyTexts(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart?.Document.Body;
        return body is null
            ? new List<string>()
            : body.Descendants<Paragraph>().Select(p => p.InnerText).ToList();
    }

    [Fact]
    public void ToStrict_ProducesAStrictPackage()
    {
        // Guard the test helper itself: the rewritten package must really be strict
        // (root document element in the purl.oclc.org namespace) or the other tests
        // would silently exercise the transitional path.
        var strict = ToStrict(Doc("Hello world."));
        using var stream = new MemoryStream(strict.DocumentByteArray);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry("word/document.xml")!.Open());
        var xml = reader.ReadToEnd();
        Assert.Contains(StrictMain, xml);
        Assert.DoesNotContain(TransitionalMain, xml);
    }

    [Fact]
    public void Compare_StrictLeft_TransitionalRight_RoundTrips()
    {
        var left = ToStrict(Doc("The quick brown fox.", "Second paragraph."));
        var right = Doc("The quick red fox.", "Second paragraph.");

        var result = DocxDiff.Compare(left, right);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(new List<string> { "The quick brown fox.", "Second paragraph." },
            BodyTexts(rejected));
    }

    [Fact]
    public void Compare_TransitionalLeft_StrictRight_RoundTrips()
    {
        var left = Doc("The quick brown fox.");
        var right = ToStrict(Doc("The quick red fox."));

        var result = DocxDiff.Compare(left, right);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(new List<string> { "The quick red fox." }, BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void Compare_BothStrict_ProducesTrackedChangesAndRoundTrips()
    {
        var left = ToStrict(Doc("Alpha bravo charlie.", "Tail."));
        var right = ToStrict(Doc("Alpha delta charlie.", "Tail."));

        var result = DocxDiff.Compare(left, right);

        using (var stream = new MemoryStream(result.DocumentByteArray))
        using (var wdoc = WordprocessingDocument.Open(stream, false))
        {
            var body = wdoc.MainDocumentPart!.Document.Body!;
            Assert.True(
                body.Descendants<InsertedRun>().Any() || body.Descendants<DeletedRun>().Any(),
                "expected w:ins/w:del markup in the redline output");
        }

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(new List<string> { "Alpha delta charlie.", "Tail." }, BodyTexts(accepted));
        Assert.Equal(new List<string> { "Alpha bravo charlie.", "Tail." }, BodyTexts(rejected));
    }

    [Fact]
    public void GetRevisions_StrictSelfCompare_IsEmpty()
    {
        var strict = ToStrict(Doc("Identity paragraph.", "Another one."));

        var revisions = DocxDiff.GetRevisions(strict, strict);

        Assert.Empty(revisions);
    }

    [Fact]
    public void GetEditScriptJson_StrictInput_Parses()
    {
        var left = ToStrict(Doc("One two three."));
        var right = Doc("One two four.");

        var json = DocxDiff.GetEditScriptJson(left, right);

        using var parsed = System.Text.Json.JsonDocument.Parse(json);
        Assert.True(parsed.RootElement.TryGetProperty("blockOps", out _)
            || parsed.RootElement.ValueKind == System.Text.Json.JsonValueKind.Object);
    }
}
