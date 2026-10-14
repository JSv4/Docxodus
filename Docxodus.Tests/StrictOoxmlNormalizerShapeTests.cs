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
/// Word's "Strict Open XML" save FOLDS the MS-2010 wordprocessingShape namespace into the strict
/// wordprocessingDrawing namespace: a shape's <c>wsp/spPr/bodyPr/…</c> elements are written in the
/// wpDrawing namespace with only <c>a:graphicData/@uri</c> marking the real payload namespace. Word
/// un-folds on open; a flat URI substitution leaves them in the TRANSITIONAL wpDrawing namespace —
/// names that do not exist there — and LibreOffice silently drops the whole shape. The normalizer
/// must re-home such descendants to the <c>@uri</c> namespace.
/// </summary>
public class StrictOoxmlNormalizerShapeTests
{
    private const string StrictWpDrawing = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
    private const string TransitionalWpDrawing = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private const string Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

    private static byte[] StrictDocWithFoldedShape()
    {
        // Minimal strict package: document.xml carries a drawing whose graphicData payload declares
        // uri=wordprocessingShape but whose child elements are FOLDED into the strict wpDrawing
        // namespace (no wps declaration anywhere) — exactly Word's strict save shape.
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<w:document xmlns:w=\"http://purl.oclc.org/ooxml/wordprocessingml/main\"" +
            " xmlns:r=\"http://purl.oclc.org/ooxml/officeDocument/relationships\"" +
            " xmlns:wp=\"" + StrictWpDrawing + "\"" +
            " xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\">" +
            "<w:body><w:p><w:r><w:drawing>" +
            "<wp:inline><wp:extent cx=\"100\" cy=\"100\"/>" +
            "<a:graphic><a:graphicData uri=\"" + Wps + "\">" +
            "<wp:wsp><wp:cNvSpPr/><wp:spPr/><wp:bodyPr/></wp:wsp>" +
            "</a:graphicData></a:graphic>" +
            "</wp:inline>" +
            "</w:drawing></w:r><w:r><w:t>shape para</w:t></w:r></w:p>" +
            "<w:sectPr/></w:body></w:document>";
        var relsXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
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

    [Fact]
    public void NormalizeToTransitional_UnfoldsShapeNamespaces()
    {
        var strict = new WmlDocument("strict.docx", StrictDocWithFoldedShape());

        var normalized = StrictOoxmlNormalizer.NormalizeToTransitional(strict);

        using var ms = new MemoryStream(normalized.DocumentByteArray);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry("word/document.xml")!.Open());
        var doc = XDocument.Parse(reader.ReadToEnd());

        XNamespace wps = Wps;
        XNamespace wpTrans = TransitionalWpDrawing;
        // The payload elements are re-homed to the wps namespace declared by graphicData/@uri…
        Assert.NotNull(doc.Descendants(wps + "wsp").SingleOrDefault());
        Assert.NotNull(doc.Descendants(wps + "spPr").SingleOrDefault());
        // …while the drawing CONTAINER (wp:inline/wp:extent) stays in transitional wpDrawing.
        Assert.NotNull(doc.Descendants(wpTrans + "inline").SingleOrDefault());
        Assert.NotNull(doc.Descendants(wpTrans + "extent").SingleOrDefault());
        // Nothing remains in wpDrawing that belongs to the shape payload.
        Assert.Null(doc.Descendants(wpTrans + "wsp").SingleOrDefault());
    }
}
