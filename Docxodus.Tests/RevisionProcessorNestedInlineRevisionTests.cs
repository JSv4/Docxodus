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
/// A deleted paragraph mark can join a paragraph whose surviving content is nested beneath
/// inline wrapper elements. Acceptance must inspect through every wrapper before deciding
/// whether the joined paragraph is empty.
/// </summary>
public class RevisionProcessorNestedInlineRevisionTests
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

        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string content)
            {
                var entry = zip.CreateEntry(name);
                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.Write(content);
            }

            Add("[Content_Types].xml", contentTypes);
            Add("_rels/.rels", relsXml);
            Add("word/document.xml", documentXml);
        }

        return new WmlDocument("nested-inline-revision.docx", stream.ToArray());
    }

    private static XElement AcceptedBody(WmlDocument document)
    {
        var accepted = RevisionProcessor.AcceptRevisions(document);
        using var stream = new MemoryStream(accepted.DocumentByteArray);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry("word/document.xml")!.Open());
        XNamespace w = WNs;
        return XDocument.Parse(reader.ReadToEnd()).Root!.Element(w + "body")!;
    }

    [Fact]
    public void Accept_DeletedParagraphMarkWithNestedHyperlinkInsertion_PreservesInsertedContent()
    {
        var document = DocFromBody(
            "<w:p><w:pPr><w:rPr><w:del w:id=\"1\" w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"/>" +
            "</w:rPr></w:pPr><w:hyperlink w:anchor=\"bookmark\"><w:ins w:id=\"2\" w:author=\"A\" " +
            "w:date=\"2026-01-01T00:00:00Z\"><w:r><w:t>linked insertion</w:t></w:r></w:ins></w:hyperlink></w:p>" +
            "<w:p><w:r><w:t>following paragraph</w:t></w:r></w:p>");

        var body = AcceptedBody(document); // regression: previously threw while checking the joined paragraph

        XNamespace w = WNs;
        var paragraphs = body.Elements(w + "p").ToList();
        var paragraph = Assert.Single(paragraphs);
        Assert.Equal("linked insertionfollowing paragraph", paragraph.Value);
        Assert.Single(paragraph.Elements(w + "hyperlink"));
        Assert.Empty(paragraph.Descendants(w + "ins"));
    }
}
