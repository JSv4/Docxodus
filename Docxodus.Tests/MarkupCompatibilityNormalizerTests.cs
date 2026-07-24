#nullable enable

using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Word resolves <c>mc:AlternateContent</c> on open and its compare output re-serializes the
/// RESOLVED content, not the wrapper. Two oracle-proven shapes: (A) a <c>mc:Choice
/// Requires="v"</c> VML payload (strict-save watermarks) is unwrapped to the bare <c>w:pict</c> —
/// LibreOffice does not render the wrapped form; (B) a Choice requiring an obsolete draft
/// namespace (Office 2008/6/28 beta wordprocessingShape) is not understood by any modern reader,
/// so the <c>mc:Fallback</c> VML is inlined instead — LibreOffice renders nothing for the
/// original. Modern DrawingML choices (canonical 2010 wps) are left untouched.
/// </summary>
public class MarkupCompatibilityNormalizerTests
{
    private const string McNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string VNs = "urn:schemas-microsoft-com:vml";
    private const string Wps2010 = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    private const string Wps2008Draft = "http://schemas.microsoft.com/office/word/2008/6/28/wordprocessingShape";

    private static byte[] DocWithBodyXml(string runInnerXml, string anchorText = "anchor text")
    {
        return DocWithParagraphXml(
            "<w:r>" + runInnerXml + "</w:r>" +
            "<w:r><w:t>" + anchorText + "</w:t></w:r>");
    }

    private static byte[] DocWithParagraphXml(string paragraphInnerXml)
    {
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<w:document xmlns:w=\"" + WNs + "\"" +
            " xmlns:mc=\"" + McNs + "\"" +
            " xmlns:v=\"" + VNs + "\">" +
            "<w:body><w:p>" + paragraphInnerXml + "</w:p>" +
            "<w:sectPr/></w:body></w:document>";
        var relsXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
            "</Relationships>";
        var contentTypes =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
            "</Types>";
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string content)
            {
                var entry = zip.CreateEntry(name);
                using var w = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                w.Write(content);
            }
            Add("[Content_Types].xml", contentTypes);
            Add("_rels/.rels", relsXml);
            Add("word/document.xml", documentXml);
        }
        return ms.ToArray();
    }

    private static XDocument MainPart(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry("word/document.xml")!.Open());
        return XDocument.Parse(reader.ReadToEnd());
    }

    [Fact]
    public void VmlChoice_IsUnwrappedToBarePict()
    {
        var doc = new WmlDocument("d.docx", DocWithBodyXml(
            "<mc:AlternateContent><mc:Choice Requires=\"v\">" +
            "<w:pict><v:shape id=\"s1\" style=\"width:10pt;height:10pt\"/></w:pict>" +
            "</mc:Choice><mc:Fallback/></mc:AlternateContent>"));

        var normalized = MarkupCompatibilityNormalizer.Normalize(doc);

        var main = MainPart(normalized);
        XNamespace mc = McNs, w = WNs, v = VNs;
        Assert.Empty(main.Descendants(mc + "AlternateContent"));
        var pict = main.Descendants(w + "pict").SingleOrDefault();
        Assert.NotNull(pict);
        Assert.NotNull(pict!.Element(v + "shape"));
    }

    [Fact]
    public void ObsoleteDraftChoice_FallsBackToVml()
    {
        var doc = new WmlDocument("d.docx", DocWithBodyXml(
            "<mc:AlternateContent xmlns:wps=\"" + Wps2008Draft + "\">" +
            "<mc:Choice Requires=\"wps\"><w:drawing/></mc:Choice>" +
            "<mc:Fallback><w:pict><v:shape id=\"fb1\" style=\"width:10pt;height:10pt\"/></w:pict></mc:Fallback>" +
            "</mc:AlternateContent>"));

        var normalized = MarkupCompatibilityNormalizer.Normalize(doc);

        var main = MainPart(normalized);
        XNamespace mc = McNs, w = WNs, v = VNs;
        Assert.Empty(main.Descendants(mc + "AlternateContent"));
        Assert.Empty(main.Descendants(w + "drawing"));
        var shape = main.Descendants(v + "shape").SingleOrDefault();
        Assert.Equal("fb1", (string?)shape?.Attribute("id"));
    }

    [Fact]
    public void Compare_ResolvesAlternateContentInInputs()
    {
        var left = new WmlDocument("l.docx", DocWithBodyXml(
            "<mc:AlternateContent><mc:Choice Requires=\"v\">" +
            "<w:pict><v:shape id=\"s1\" style=\"width:10pt;height:10pt\"/></w:pict>" +
            "</mc:Choice><mc:Fallback/></mc:AlternateContent>"));
        var right = new WmlDocument("r.docx", DocWithBodyXml(
            "<mc:AlternateContent><mc:Choice Requires=\"v\">" +
            "<w:pict><v:shape id=\"s1\" style=\"width:10pt;height:10pt\"/></w:pict>" +
            "</mc:Choice><mc:Fallback/></mc:AlternateContent>",
            "changed anchor text"));

        var result = DocxDiff.Compare(left, right);

        var main = MainPart(result);
        XNamespace mc = McNs, w = WNs;
        Assert.Empty(main.Descendants(mc + "AlternateContent"));
        Assert.NotEmpty(main.Descendants(w + "pict"));
    }

    [Fact]
    public void ModernWpsChoice_IsLeftUntouched()
    {
        var doc = new WmlDocument("d.docx", DocWithBodyXml(
            "<mc:AlternateContent xmlns:wps=\"" + Wps2010 + "\">" +
            "<mc:Choice Requires=\"wps\"><w:drawing/></mc:Choice>" +
            "<mc:Fallback><w:pict><v:shape id=\"fb1\"/></w:pict></mc:Fallback>" +
            "</mc:AlternateContent>"));

        var normalized = MarkupCompatibilityNormalizer.Normalize(doc);

        // Same instance back — nothing to rewrite.
        Assert.Same(doc, normalized);
        var main = MainPart(normalized);
        XNamespace mc = McNs;
        Assert.Single(main.Descendants(mc + "AlternateContent"));
    }

    [Fact]
    public void DisjointDuplicateParagraphProperties_AreCoalescedAndOrdered()
    {
        var doc = new WmlDocument("d.docx", DocWithParagraphXml(
            "<w:pPr w:rsidR=\"001\"><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr>" +
            "<w:r><w:t>anchor text</w:t></w:r>" +
            "<w:pPr w:rsidRDefault=\"002\"><w:numPr><w:numId w:val=\"42\"/></w:numPr></w:pPr>"));

        var normalized = MarkupCompatibilityNormalizer.Normalize(doc);

        Assert.NotSame(doc, normalized);
        XNamespace w = WNs;
        var paragraph = MainPart(normalized).Descendants(w + "p").Single();
        var properties = paragraph.Elements(w + "pPr").ToList();
        var pPr = Assert.Single(properties);
        Assert.Same(pPr, paragraph.Elements().First());
        Assert.Equal(new[] { w + "numPr", w + "spacing" }, pPr.Elements().Select(e => e.Name));
        Assert.Equal("001", (string?)pPr.Attribute(w + "rsidR"));
        Assert.Equal("002", (string?)pPr.Attribute(w + "rsidRDefault"));
    }

    [Fact]
    public void ConflictingDuplicateParagraphProperties_AreLeftUntouched()
    {
        var doc = new WmlDocument("d.docx", DocWithParagraphXml(
            "<w:pPr><w:spacing w:after=\"0\"/></w:pPr>" +
            "<w:r><w:t>anchor text</w:t></w:r>" +
            "<w:pPr><w:spacing w:after=\"240\"/></w:pPr>"));

        var normalized = MarkupCompatibilityNormalizer.Normalize(doc);

        Assert.Same(doc, normalized);
        XNamespace w = WNs;
        Assert.Equal(2, MainPart(normalized).Descendants(w + "p").Single().Elements(w + "pPr").Count());
    }

    [Fact]
    public void RevisionBearingDuplicateParagraphProperties_AreLeftUntouched()
    {
        var doc = new WmlDocument("d.docx", DocWithParagraphXml(
            "<w:pPr><w:numPr><w:numId w:val=\"42\"/></w:numPr></w:pPr>" +
            "<w:r><w:t>anchor text</w:t></w:r>" +
            "<w:pPr><w:pPrChange w:id=\"1\" w:author=\"test\" w:date=\"2026-07-18T00:00:00Z\"><w:pPr><w:spacing w:after=\"0\"/></w:pPr></w:pPrChange></w:pPr>"));

        var normalized = MarkupCompatibilityNormalizer.Normalize(doc);

        Assert.Same(doc, normalized);
    }

    [Fact]
    public void Compare_NormalizesDisjointDuplicateParagraphPropertiesBeforeDiff()
    {
        var left = new WmlDocument("l.docx", DocWithParagraphXml(
            "<w:pPr><w:spacing w:after=\"0\"/></w:pPr>" +
            "<w:r><w:t>anchor text</w:t></w:r>" +
            "<w:pPr><w:numPr><w:numId w:val=\"42\"/></w:numPr></w:pPr>"));
        var right = new WmlDocument("r.docx", DocWithParagraphXml(
            "<w:pPr><w:spacing w:after=\"0\"/></w:pPr>" +
            "<w:r><w:t>changed anchor text</w:t></w:r>" +
            "<w:pPr><w:numPr><w:numId w:val=\"42\"/></w:numPr></w:pPr>"));

        var result = DocxDiff.Compare(left, right);

        XNamespace w = WNs;
        var paragraph = MainPart(result).Descendants(w + "p").Single();
        var pPr = Assert.Single(paragraph.Elements(w + "pPr"));
        Assert.Same(pPr, paragraph.Elements().First());
    }
}
