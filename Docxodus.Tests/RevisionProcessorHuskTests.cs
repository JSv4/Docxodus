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
/// Accepting a paragraph-mark insertion (<c>w:pPr/w:rPr/w:ins</c>) must not leave an empty
/// <c>&lt;w:rPr/&gt;</c> (or fully empty <c>&lt;w:pPr/&gt;</c>) husk behind — Word's accept
/// removes the emptied shells. The husk is not cosmetic: text-identical paragraphs from two
/// accepted inputs then differ structurally, which desynchronizes downstream block matching
/// (the comments-redline corpus family diffs ~250 phantom revisions over equal text).
/// </summary>
public class RevisionProcessorHuskTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static WmlDocument DocFromBody(string bodyXml)
    {
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<w:document xmlns:w=\"" + WNs + "\"><w:body>" + bodyXml + "<w:sectPr/></w:body></w:document>";
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
        return new WmlDocument("d.docx", ms.ToArray());
    }

    private static XElement AcceptedBody(WmlDocument doc)
    {
        var accepted = RevisionProcessor.AcceptRevisions(doc);
        using var ms = new MemoryStream(accepted.DocumentByteArray);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry("word/document.xml")!.Open());
        XNamespace w = WNs;
        return XDocument.Parse(reader.ReadToEnd()).Root!.Element(w + "body")!;
    }

    [Fact]
    public void Accept_MarkInsOnly_RemovesEmptyShells()
    {
        var doc = DocFromBody(
            "<w:p><w:pPr><w:rPr><w:ins w:id=\"1\" w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"/></w:rPr></w:pPr>" +
            "<w:r><w:t>inserted paragraph text</w:t></w:r></w:p>");

        var body = AcceptedBody(doc);

        XNamespace w = WNs;
        var para = body.Elements(w + "p").Single();
        var pPr = para.Element(w + "pPr");
        // Word's accept: mark-ins removed AND the emptied rPr (and thus pPr) shells removed.
        Assert.True(pPr is null || (!pPr.HasElements && !pPr.HasAttributes),
            $"expected no residual pPr shell, got: {pPr}");
    }

    [Fact]
    public void Accept_MarkInsBesideRealProps_KeepsPropsDropsEmptyRPr()
    {
        var doc = DocFromBody(
            "<w:p><w:pPr><w:pStyle w:val=\"ListBullet\"/>" +
            "<w:rPr><w:ins w:id=\"1\" w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"/></w:rPr></w:pPr>" +
            "<w:r><w:t>inserted bullet text</w:t></w:r></w:p>");

        var body = AcceptedBody(doc);

        XNamespace w = WNs;
        var pPr = body.Elements(w + "p").Single().Element(w + "pPr");
        Assert.NotNull(pPr);
        Assert.Equal("ListBullet", (string?)pPr!.Element(w + "pStyle")?.Attribute(w + "val"));
        Assert.Null(pPr.Element(w + "rPr"));
    }

    [Fact]
    public void Accept_MarkRPrWithRealProps_KeepsThem()
    {
        // The mark rPr carries a real property besides the revision element — it must survive.
        var doc = DocFromBody(
            "<w:p><w:pPr><w:rPr><w:b/>" +
            "<w:ins w:id=\"1\" w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"/></w:rPr></w:pPr>" +
            "<w:r><w:t>text</w:t></w:r></w:p>");

        var body = AcceptedBody(doc);

        XNamespace w = WNs;
        var rPr = body.Elements(w + "p").Single().Element(w + "pPr")?.Element(w + "rPr");
        Assert.NotNull(rPr);
        Assert.NotNull(rPr!.Element(w + "b"));
        Assert.Null(rPr.Element(w + "ins"));
    }
}
