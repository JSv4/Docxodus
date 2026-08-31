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

    private static byte[] DocWithParagraphXml(string paragraphInnerXml) =>
        DocWithBodyChildrenXml("<w:p>" + paragraphInnerXml + "</w:p>");

    private static byte[] DocWithBodyChildrenXml(
        string bodyChildrenXml, params (string Name, string Xml)[] extraParts)
    {
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<w:document xmlns:w=\"" + WNs + "\"" +
            " xmlns:mc=\"" + McNs + "\"" +
            " xmlns:v=\"" + VNs + "\">" +
            "<w:body>" + bodyChildrenXml +
            "<w:sectPr/></w:body></w:document>";
        return Package(documentXml, extraParts);
    }

    private static byte[] Package(string documentXml, params (string Name, string Xml)[] extraParts)
    {
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
            foreach (var (name, xml) in extraParts)
                Add(name, xml);
        }
        return ms.ToArray();
    }

    private static XDocument PartOf(WmlDocument doc, string entryName)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry(entryName)!.Open());
        return XDocument.Parse(reader.ReadToEnd());
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

    // ─── The detection gate (issue #619) ────────────────────────────────
    //
    // Deciding whether a part needs either repair used to mean building its whole XDocument, and
    // the literal that gated it ("pPr") is in every real word/document.xml, so the gate never
    // closed. A streaming pass answers the same two questions without a DOM. The tests below are
    // about that gate specifically: it must stay a SUPERSET of what the repairs act on, because a
    // part that needs repair and is skipped is a correctness bug where an over-broad gate is only
    // a performance one.

    /// <summary>
    /// The two duplicates are separated by a whole nested subtree — the shape a comparer produces
    /// when it inserts revision runs between the original malformed property elements. A detector
    /// tracking the open-element stack has to still pair them across that depth excursion.
    /// </summary>
    [Fact]
    public void Gate_FindsDuplicateParagraphPropertiesSeparatedByNestedContent()
    {
        var doc = new WmlDocument("d.docx", DocWithParagraphXml(
            "<w:pPr><w:spacing w:after=\"0\"/></w:pPr>" +
            "<w:ins w:id=\"9\" w:author=\"a\" w:date=\"2026-01-01T00:00:00Z\">" +
            "<w:r><w:rPr><w:b/></w:rPr><w:t>inserted</w:t></w:r></w:ins>" +
            "<w:pPr><w:numPr><w:numId w:val=\"42\"/></w:numPr></w:pPr>"));

        var normalized = MarkupCompatibilityNormalizer.Normalize(doc);

        XNamespace w = WNs;
        Assert.NotSame(doc, normalized);
        Assert.Single(MainPart(normalized).Descendants(w + "p").First().Elements(w + "pPr"));
    }

    /// <summary>
    /// The duplicate is inside a text-box paragraph nested in another paragraph, while the OUTER
    /// paragraph is well formed. Counting per paragraph rather than per part is what makes this
    /// visible; the repair walks every descendant paragraph, so the gate has to as well.
    /// </summary>
    [Fact]
    public void Gate_FindsDuplicateParagraphPropertiesInsideANestedParagraph()
    {
        var doc = new WmlDocument("d.docx", DocWithParagraphXml(
            "<w:pPr><w:spacing w:after=\"0\"/></w:pPr>" +
            "<w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p>" +
            "<w:pPr><w:spacing w:before=\"120\"/></w:pPr>" +
            "<w:r><w:t>boxed</w:t></w:r>" +
            "<w:pPr><w:numPr><w:numId w:val=\"7\"/></w:numPr></w:pPr>" +
            "</w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>"));

        var normalized = MarkupCompatibilityNormalizer.Normalize(doc);

        XNamespace w = WNs;
        Assert.NotSame(doc, normalized);
        var nested = MainPart(normalized).Descendants(w + "p")
            .Single(p => !p.Descendants(w + "p").Any());
        Assert.Single(nested.Elements(w + "pPr"));
        Assert.Contains("boxed", nested.Descendants(w + "t").Select(t => (string)t));
    }

    /// <summary>The counterpart: three well-formed sibling paragraphs carry three
    /// <c>w:pPr</c> between them, and none of them is a duplicate. Counting per part rather than
    /// per paragraph would read that as a repair candidate.</summary>
    [Fact]
    public void Gate_DoesNotMistakeOnePPrPerSiblingParagraphForADuplicate()
    {
        var doc = new WmlDocument("d.docx", DocWithBodyChildrenXml(
            "<w:p><w:pPr><w:spacing w:after=\"0\"/></w:pPr><w:r><w:t>one</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:spacing w:after=\"1\"/></w:pPr><w:r><w:t>two</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:spacing w:after=\"2\"/></w:pPr><w:r><w:t>three</w:t></w:r></w:p>"));

        Assert.Same(doc, MarkupCompatibilityNormalizer.Normalize(doc));
    }

    /// <summary>Every <c>.xml</c> part is a candidate, not just <c>word/document.xml</c>.</summary>
    [Fact]
    public void Gate_ResolvesAlternateContentInAPartOtherThanTheMainDocument()
    {
        var headerXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<w:hdr xmlns:w=\"" + WNs + "\" xmlns:mc=\"" + McNs + "\" xmlns:v=\"" + VNs + "\">" +
            "<w:p><w:r>" +
            "<mc:AlternateContent><mc:Choice Requires=\"v\">" +
            "<w:pict><v:shape id=\"wm\"/></w:pict>" +
            "</mc:Choice></mc:AlternateContent>" +
            "</w:r></w:p></w:hdr>";
        var doc = new WmlDocument("d.docx", DocWithBodyChildrenXml(
            "<w:p><w:r><w:t>body</w:t></w:r></w:p>", ("word/header1.xml", headerXml)));

        var normalized = MarkupCompatibilityNormalizer.Normalize(doc);

        XNamespace mc = McNs;
        XNamespace w = WNs;
        Assert.NotSame(doc, normalized);
        var header = PartOf(normalized, "word/header1.xml");
        Assert.Empty(header.Descendants(mc + "AlternateContent"));
        Assert.Single(header.Descendants(w + "pict"));
    }

    /// <summary>
    /// The repairs are namespace-exact and prefix-agnostic, so the gate matches on local names.
    /// A part that binds the WordprocessingML namespace to some prefix other than <c>w</c> is
    /// still repaired — which the literal substring gate this replaced also managed, and which is
    /// the property that must not regress.
    /// </summary>
    [Fact]
    public void Gate_FindsDuplicatesUnderANonstandardNamespacePrefix()
    {
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<x:document xmlns:x=\"" + WNs + "\"><x:body><x:p>" +
            "<x:pPr><x:spacing x:after=\"0\"/></x:pPr>" +
            "<x:r><x:t>anchor text</x:t></x:r>" +
            "<x:pPr><x:numPr><x:numId x:val=\"42\"/></x:numPr></x:pPr>" +
            "</x:p><x:sectPr/></x:body></x:document>";
        var doc = new WmlDocument("d.docx", Package(documentXml));

        var normalized = MarkupCompatibilityNormalizer.Normalize(doc);

        XNamespace w = WNs;
        Assert.NotSame(doc, normalized);
        Assert.Single(MainPart(normalized).Descendants(w + "p").Single().Elements(w + "pPr"));
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
