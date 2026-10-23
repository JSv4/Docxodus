using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1268: a STRICT-conformance package (purl.oclc.org namespaces and relationship types)
/// must resolve and render like its transitional twin through the WmlDocument entry points and
/// <see cref="DocxSession"/>. Each test runs on both conformance classes from the same content: a
/// paragraph that deletes "Amber prism" and inserts "Indigo prism". The transitional case is the
/// control.
/// </summary>
public class StrictOoxmlRevisionTests
{
    private const string RevisionParagraph =
        "<w:p><w:del w:id=\"31\" w:author=\"Review Example\" w:date=\"2026-09-12T10:00:00Z\"><w:r><w:delText>Amber prism</w:delText></w:r></w:del>" +
        "<w:ins w:id=\"32\" w:author=\"Review Example\" w:date=\"2026-09-12T10:01:00Z\"><w:r><w:t xml:space=\"preserve\">Indigo prism</w:t></w:r></w:ins></w:p>";

    private static byte[] RevisionPackage(bool strict) => Package(strict, RevisionParagraph);

    /// <summary>A minimal package holding <paramref name="paragraphs"/> and, after Normal, the
    /// <paramref name="styles"/>, written in the Strict or the Transitional conformance class.</summary>
    private static byte[] Package(bool strict, string paragraphs, string styles = "", string? extraEntry = null)
    {
        var w = strict ? "http://purl.oclc.org/ooxml/wordprocessingml/main" : "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var rel = strict ? "http://purl.oclc.org/ooxml/officeDocument/relationships" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string ct = "application/vnd.openxmlformats-officedocument.wordprocessingml";
        var parts = new (string Name, string Xml)[]
        {
            ("[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                $"<Override PartName=\"/word/document.xml\" ContentType=\"{ct}.document.main+xml\"/>" +
                $"<Override PartName=\"/word/styles.xml\" ContentType=\"{ct}.styles+xml\"/>" +
                $"<Override PartName=\"/word/settings.xml\" ContentType=\"{ct}.settings+xml\"/></Types>"),
            ("_rels/.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                $"<Relationship Id=\"document\" Type=\"{rel}/officeDocument\" Target=\"word/document.xml\"/></Relationships>"),
            ("word/_rels/document.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                $"<Relationship Id=\"styles\" Type=\"{rel}/styles\" Target=\"styles.xml\"/>" +
                $"<Relationship Id=\"settings\" Type=\"{rel}/settings\" Target=\"settings.xml\"/></Relationships>"),
            ("word/document.xml",
                $"<w:document xmlns:w=\"{w}\" xmlns:r=\"{rel}\"><w:body>{paragraphs}" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>"),
            ("word/styles.xml",
                $"<w:styles xmlns:w=\"{w}\"><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>{styles}</w:styles>"),
            ("word/settings.xml", $"<w:settings xmlns:w=\"{w}\"/>"),
        };
        if (extraEntry is not null)
            parts = parts.Append((extraEntry, $"<root xmlns:w=\"{w}\">{new string('x', 400)}</root>")).ToArray();
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, xml) in parts)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                writer.Write(xml);
            }
        }
        return ms.ToArray();
    }

    private static WmlDocument RevisionDocument(bool strict) => new("revisions.docx", RevisionPackage(strict));

    private static XDocument Part(WmlDocument document, string name)
    {
        using var ms = new MemoryStream(document.DocumentByteArray);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        using var stream = zip.GetEntry(name)!.Open();
        return XDocument.Load(stream);
    }

    private static int GlossaryDocPartCount(WmlDocument document)
    {
        using var ms = new MemoryStream(document.DocumentByteArray);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var entry = zip.GetEntry("word/glossary/document.xml");
        if (entry is null)
            return 0;
        using var stream = entry.Open();
        return XDocument.Load(stream).Descendants(W.docPart).Count();
    }

    /// <summary>The body's text, after asserting no insertion, deletion or deleted text remains.</summary>
    private static string ResolvedBodyText(byte[] package)
    {
        using var ms = new MemoryStream(package);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        using var stream = zip.GetEntry("word/document.xml")!.Open();
        var body = XDocument.Load(stream).Root!.Element(W.body)!;
        Assert.DoesNotContain(body.Descendants(), e => e.Name == W.ins || e.Name == W.del || e.Name == W.delText);
        return string.Concat(body.Descendants(W.t).Select(t => t.Value));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void AcceptRevisions_KeepsOnlyTheInsertedText(bool strict)
    {
        var accepted = RevisionProcessor.AcceptRevisions(RevisionDocument(strict));
        Assert.Equal("Indigo prism", ResolvedBodyText(accepted.DocumentByteArray));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RejectRevisions_KeepsOnlyTheDeletedText(bool strict)
    {
        var rejected = RevisionProcessor.RejectRevisions(RevisionDocument(strict));
        Assert.Equal("Amber prism", ResolvedBodyText(rejected.DocumentByteArray));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void HasTrackedRevisions_SeesTheRevisions(bool strict)
    {
        Assert.True(RevisionProcessor.HasTrackedRevisions(RevisionDocument(strict)));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ConvertToHtml_RendersBothSidesOfTheTrackedChange(bool strict)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(
            RevisionDocument(strict),
            new WmlToHtmlConverterSettings { RenderTrackedChanges = true }).ToString();
        Assert.Contains("Amber prism", html);
        Assert.Contains("Indigo prism", html);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MarkdownProjection_ShowsTheCurrentText(bool strict)
    {
        var markdown = WmlToMarkdownConverter.Convert(RevisionDocument(strict), new WmlToMarkdownConverterSettings()).Markdown;
        Assert.Contains("Indigo prism", markdown);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void BuildDocument_AcceptsTheSourceAndKeepsItsRevisions(bool strict)
    {
        var built = DocumentBuilder.BuildDocument(new List<Source> { new(RevisionDocument(strict), true) });
        Assert.True(RevisionProcessor.HasTrackedRevisions(built));
        Assert.Equal("Indigo prism", ResolvedBodyText(RevisionProcessor.AcceptRevisions(built).DocumentByteArray));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void BuildDocument_KeepsTheSourcesGlossary(bool strict)
    {
        var template = new WmlDocument("../../../../TestFiles/DA225-TemplateDocument.docx");
        var source = strict ? DocxDiffStrictConformanceTests.AsStrict(template) : template;
        var built = DocumentBuilder.BuildDocument(new List<Source> { new(source, true) });
        Assert.Equal(GlossaryDocPartCount(template), GlossaryDocPartCount(built));
        Assert.True(GlossaryDocPartCount(template) > 0);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void BuildDocument_NormalizeStyleIds_KeepsEachStyleNameOnItsOwnId(bool strict)
    {
        static string Styled(string text) => $"<w:p><w:pPr><w:pStyle w:val=\"S1\"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";
        static string Style(string name) => $"<w:style w:type=\"paragraph\" w:styleId=\"S1\"><w:name w:val=\"{name}\"/></w:style>";
        var a = new WmlDocument("a.docx", Package(false, Styled("from A"), Style("Alpha")));
        var b = new WmlDocument("b.docx", Package(strict, Styled("from B"), Style("Beta")));

        var built = DocumentBuilder.BuildDocument(
            new List<Source> { new(a, true), new(b, true) },
            new DocumentBuilderSettings { NormalizeStyleIds = true });

        var styleNames = Part(built, "word/styles.xml").Descendants(W.style)
            .ToDictionary(e => (string)e.Attribute(W.styleId)!, e => (string)e.Element(W.name)!.Attribute(W.val)!);
        var paragraphStyles = Part(built, "word/document.xml").Descendants(W.p).Where(p => p.Descendants(W.t).Any())
            .ToDictionary(p => string.Concat(p.Descendants(W.t).Select(t => t.Value)),
                          p => styleNames[(string)p.Element(W.pPr)!.Element(W.pStyle)!.Attribute(W.val)!]);
        Assert.Equal("Alpha", paragraphStyles["from A"]);
        Assert.Equal("Beta", paragraphStyles["from B"]);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OtherWmlDocumentReaders_SeeTheContent(bool strict)
    {
        var document = new WmlDocument("readers.docx", Package(strict, "<w:p><w:r><w:t>Plain prism</w:t></w:r></w:p>" + RevisionParagraph));
        Assert.Equal(W.w, document.MainDocumentPart.Name.Namespace);
        Assert.NotEmpty(DocumentStructureAnalyzer.Analyze(document).Root.Children);
        Assert.Contains("Plain prism", OpenContractExporter.Export(document).Content);
        Assert.NotNull(MetricsGetter.GetDocxMetrics(document, new MetricsGetterSettings()).Element(H.RevisionTracking));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SessionAcceptAndRejectAll_ResolveTheRevisions(bool strict)
    {
        using (var session = new DocxSession(RevisionPackage(strict)))
        {
            session.AcceptAllRevisions();
            Assert.Equal("Indigo prism", ResolvedBodyText(session.Save()));
        }
        using (var session = new DocxSession(RevisionPackage(strict)))
        {
            session.RejectAllRevisions();
            Assert.Equal("Amber prism", ResolvedBodyText(session.Save()));
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SessionSemanticChanges_ReportNothingBeforeAnyEdit(bool strict)
    {
        using var session = new DocxSession(RevisionPackage(strict));
        Assert.Equal(0, session.GetSemanticChanges().ChangeCount);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SessionDeliveryEvidence_KeepsTheCallersExactOpeningBytes(bool strict)
    {
        var opening = RevisionPackage(strict);
        using var session = new DocxSession(opening);
        Assert.Equal(opening, session.ExportDeliveryEvidence().Source.Bytes);
    }

    [Fact]
    public void SemanticDiff_FindsNoChangeBetweenAStrictPackageAndItsTransitionalTwin()
    {
        Assert.Equal(0, Docxodus.Verification.SemanticDiff.Compare(RevisionPackage(true), RevisionPackage(false)).ChangeCount);
    }

    [Fact]
    public void SessionDeliveryReceipt_ReportsTheSameDeltaForAStrictSessionAsForItsTwin()
    {
        static int DeliveredDeltaCount(bool strict)
        {
            using var session = new DocxSession(RevisionPackage(strict), new DocxSessionSettings { CaptureDeliveryEvidence = true });
            var anchor = session.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:body", StringComparison.Ordinal));
            session.InsertParagraph(anchor, Position.After, "Added clause.");
            var delta = session.BuildDeliveryReceipt().GetArtifactBytes("semantic-source-to-delivered");
            return System.Text.Json.JsonDocument.Parse(delta).RootElement.GetProperty("changeCount").GetInt32();
        }

        Assert.Equal(DeliveredDeltaCount(strict: false), DeliveredDeltaCount(strict: true));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ADamagedPartNothingReads_DoesNotStopTheConversion(bool strict)
    {
        // An unreferenced part whose deflated bytes are damaged: no reader opens it, so the rest
        // of the package must still be converted and resolved.
        var package = Package(strict, RevisionParagraph, extraEntry: "customXml/item9.xml");
        Damage(package, "customXml/item9.xml");
        var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("damaged.docx", package));
        Assert.Equal("Indigo prism", ResolvedBodyText(accepted.DocumentByteArray));
        using var session = new DocxSession(package);
        session.RejectAllRevisions();
        Assert.Equal("Amber prism", ResolvedBodyText(session.Save()));
    }

    [Fact]
    public void Session_ReportsADamagedStrictPackageAsItsTransitionalTwin()
    {
        // Damage [Content_Types].xml, which only the Strict conversion reads before the package is
        // opened: an unconvertible package must fail the way it did before conversion existed.
        static byte[] Damaged(bool strict)
        {
            var package = RevisionPackage(strict);
            Damage(package, "[Content_Types].xml");
            return package;
        }

        var control = Record.Exception(() => new DocxSession(Damaged(strict: false)).Dispose());
        var strict = Record.Exception(() => new DocxSession(Damaged(strict: true)).Dispose());
        Assert.NotNull(control);
        Assert.IsType(control.GetType(), strict);
    }

    /// <summary>Corrupt the first deflated bytes of the named zip entry.</summary>
    private static void Damage(byte[] package, string name)
    {
        var header = FindLocalHeader(package, name);
        var dataStart = header + 30 + name.Length + (package[header + 28] | (package[header + 29] << 8));
        for (var i = dataStart; i < dataStart + 24; i++)
            package[i] ^= 0x5A;
    }

    private static int FindLocalHeader(byte[] package, string name)
    {
        var signature = new byte[] { 0x50, 0x4B, 0x03, 0x04 };
        var nameBytes = Encoding.ASCII.GetBytes(name);
        for (var i = 0; i + 30 + nameBytes.Length <= package.Length; i++)
            if (package.AsSpan(i, 4).SequenceEqual(signature) && package.AsSpan(i + 30, nameBytes.Length).SequenceEqual(nameBytes))
                return i;
        throw new InvalidDataException(name + " not found");
    }
}
