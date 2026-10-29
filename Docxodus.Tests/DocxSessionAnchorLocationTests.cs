using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml.XPath;
using Docxodus.Tests.Ir;
using DocumentFormat.OpenXml.Packaging;
using Word = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionAnchorLocationTests
{
    private static byte[] Fixture()
    {
        var source = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Alpha</w:t></w:r></w:p><!--not an element--><w:p/>"
            + "<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>Wrapped</w:t></w:r></w:p></w:sdtContent></w:sdt>"
            + "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid><w:tr><w:tc><w:tcPr/>"
            + "<w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/>");
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        stream.Position = 0;
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var main = doc.MainDocumentPart!;
            for (int i = 0; i < 2; i++)
                main.AddNewPart<HeaderPart>().Header = new Word.Header(Paragraph("Same header"));
            main.AddNewPart<FooterPart>().Footer = new Word.Footer(Paragraph("Footer"));
            main.AddNewPart<FootnotesPart>().Footnotes = new Word.Footnotes(
                new Word.Footnote(Paragraph("Separator")) { Id = -1, Type = Word.FootnoteEndnoteValues.Separator },
                new Word.Footnote(Paragraph("Note")) { Id = 1 });
            main.AddNewPart<EndnotesPart>().Endnotes = new Word.Endnotes(new Word.Endnote(Paragraph("Endnote")) { Id = 1 });
            main.AddNewPart<WordprocessingCommentsPart>().Comments = new Word.Comments(
                new Word.Comment(Paragraph("Comment")) { Id = "0", Author = "Cedar" });
        }
        return stream.ToArray();

        static Word.Paragraph Paragraph(string text) => new(new Word.Run(new Word.Text(text)));
    }

    private static DocxSessionSettings Settings(AnchorIdRendering rendering = AnchorIdRendering.FullUnid) => new()
    {
        CaptureInitialProjection = false,
        ProjectionSettings = new WmlToMarkdownConverterSettings { Scopes = ProjectionScopes.All, AnchorIdRendering = rendering },
    };

    private static string Snapshot(DocxSession session) => string.Join("\n",
        WmlToMarkdownConverter.ProjectedScopes(session.LiveDocument.MainDocumentPart!, session.ProjectionSettings)
            .Select(scope => scope.Part.Uri + "=" + scope.Part.GetXDocument().ToString(SaveOptions.DisableFormatting)));

    [Theory]
    [InlineData(AnchorIdRendering.FullUnid)]
    [InlineData(AnchorIdRendering.Abbreviated)]
    [InlineData(AnchorIdRendering.Sequential)]
    public void Fresh_query_never_stamps_the_live_parts_and_matches_every_projected_key(AnchorIdRendering rendering)
    {
        var bytes = Fixture();
        using var session = new DocxSession(bytes, Settings(rendering));
        var before = Snapshot(session);
        var result = session.GetAnchorLocations();
        Assert.Equal(before, Snapshot(session));
        Assert.Equal(0, result.Version);
        Assert.False(session.Undo());
        Assert.False(session.Redo());
        Assert.DoesNotContain("Unid=", Snapshot(session));

        var projected = session.Project().AnchorIndex;
        Assert.Equal(projected.Keys.Order(), result.Locations.Keys.Order());
        using var original = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        foreach (var (id, target) in projected)
        {
            var location = Assert.IsType<AnchorLocation>(result.Locations[id]);
            Assert.Equal(target.Anchor.Id, location.AnchorId);
            Assert.Equal(target.PartUri, location.PartUri);
            using var stream = original.GetEntry(location.PartUri.TrimStart('/'))!.Open();
            var element = Assert.IsType<XElement>(XDocument.Load(stream).XPathSelectElement(location.ElementPath));
            Assert.Equal(target.Resolve(session.LiveDocument)!.Name.LocalName, element.Name.LocalName);
            Assert.Equal(Text(target.Resolve(session.LiveDocument)!), Text(element));
        }
        Assert.Equal(2, result.Locations.Values.Where(l => l!.PartUri.Contains("header")).Select(l => l!.PartUri).Distinct().Count());
        Assert.DoesNotContain(result.Locations.Values, l => l!.ElementPath == "/*[1]/*[1]" && l.PartUri.Contains("footnotes"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Inherited_list_styles_keep_the_same_projected_key(bool capture)
    {
        var body = IrTestDocuments.FromBodyXml("<w:p><w:pPr><w:pStyle w:val=\"Derived\"/></w:pPr><w:r><w:t>Styled item</w:t></w:r></w:p>");
        using var bytes = new MemoryStream();
        bytes.Write(body.DocumentByteArray);
        bytes.Position = 0;
        using (var doc = WordprocessingDocument.Open(bytes, true))
        {
            Write(doc.MainDocumentPart!.StyleDefinitionsPart!,
                $"<w:styles xmlns:w=\"{IrTestDocuments.W}\"><w:style w:type=\"paragraph\" w:styleId=\"BaseList\">"
                + "<w:name w:val=\"BaseList\"/><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr></w:style>"
                + "<w:style w:type=\"paragraph\" w:styleId=\"Derived\"><w:name w:val=\"Derived\"/><w:basedOn w:val=\"BaseList\"/></w:style></w:styles>");
            Write(doc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>(),
                $"<w:numbering xmlns:w=\"{IrTestDocuments.W}\"><w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\">"
                + "<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>"
                + "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>");
        }
        using var session = new DocxSession(bytes.ToArray(), new DocxSessionSettings { CaptureInitialProjection = capture });
        var before = Snapshot(session);
        var locations = session.GetAnchorLocations();
        Assert.Equal(before, Snapshot(session));
        var projected = Assert.Single(session.Project().AnchorIndex).Key;
        Assert.StartsWith("li:body:", projected);
        Assert.Contains(projected, locations.Locations.Keys);
        Assert.NotNull(session.GetAnchorLocations(new[] { projected }).Locations[projected]);

        static void Write(OpenXmlPart part, string xml)
        {
            using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream);
            writer.Write(xml);
        }
    }

    [Fact]
    public void A_requested_subset_retains_unknown_ids_and_deduplicates_without_projection()
    {
        using var session = new DocxSession(Fixture(), Settings());
        var all = session.GetAnchorLocations();
        var key = all.Locations.Keys.First();
        var before = Snapshot(session);
        var subset = session.GetAnchorLocations(new[] { key, "p:body:unknown", key });
        Assert.Equal(2, subset.Locations.Count);
        Assert.Equal(all.Locations[key], subset.Locations[key]);
        Assert.Null(subset.Locations["p:body:unknown"]);
        Assert.Empty(session.GetAnchorLocations(Array.Empty<string>()).Locations);
        Assert.Equal(before, Snapshot(session));
    }

    [Fact]
    public void Positional_paths_follow_live_structure_and_report_the_version_through_undo_and_redo()
    {
        using var session = new DocxSession(Fixture(), Settings());
        var first = session.GetAnchorLocations().Locations.Single(pair => pair.Value!.PartUri == "/word/document.xml"
            && pair.Value.ElementPath == "/*[1]/*[1]/*[1]");
        Assert.True(session.InsertParagraph(first.Key, Position.Before, "Inserted").Success);
        Assert.Equal(1, session.GetAnchorLocations().Version);
        Assert.Equal("/*[1]/*[1]/*[2]", session.GetAnchorLocations(new[] { first.Key }).Locations[first.Key]!.ElementPath);
        Assert.True(session.Undo());
        Assert.Equal(2, session.GetAnchorLocations().Version);
        Assert.Equal(first.Value, session.GetAnchorLocations(new[] { first.Key }).Locations[first.Key]);
        Assert.True(session.Redo());
        Assert.Equal(3, session.GetAnchorLocations().Version);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Locations_address_the_callers_XML_even_when_open_repairs_types_or_Strict_namespaces(bool strict)
    {
        var bytes = Rewrite(Fixture(), strict);
        using var session = new DocxSession(bytes, Settings());
        var before = Snapshot(session);
        var result = session.GetAnchorLocations();
        Assert.Equal(before, Snapshot(session));
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        foreach (var location in result.Locations.Values)
        {
            using var part = zip.GetEntry(location!.PartUri.TrimStart('/'))!.Open();
            Assert.NotNull(XDocument.Load(part).XPathSelectElement(location.ElementPath));
        }
    }

    [Fact]
    public void Disposed_queries_throw_instead_of_reading_retained_state()
    {
        var session = new DocxSession(Fixture(), Settings());
        session.Dispose();
        Assert.Throws<ObjectDisposedException>(() => session.GetAnchorLocations());
    }

    private static string Text(XElement element) => string.Concat(element.DescendantsAndSelf()
        .Where(e => e.Name.LocalName == "t").Select(e => e.Value));

    private static byte[] Rewrite(byte[] bytes, bool strict)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, true))
        foreach (var entry in zip.Entries.Where(e => e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".rels")).ToArray())
        {
            string xml;
            using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
            if (strict) xml = xml.Replace(IrTestDocuments.W, "http://purl.oclc.org/ooxml/wordprocessingml/main")
                .Replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships");
            else if (entry.FullName == "[Content_Types].xml")
            {
                var types = XDocument.Parse(xml);
                types.Root!.Elements().Where(e => (string?)e.Attribute("PartName") == "/word/styles.xml").Remove();
                types.Root.Elements().Single(e => (string?)e.Attribute("Extension") == "xml")
                    .SetAttributeValue("ContentType", "application/xml");
                if (!types.Root.Elements().Any(e => (string?)e.Attribute("PartName") == "/word/document.xml"))
                    types.Root.Add(new XElement(types.Root.Name.Namespace + "Override",
                        new XAttribute("PartName", "/word/document.xml"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")));
                xml = types.ToString();
            }
            using var output = entry.Open();
            output.SetLength(0);
            using var writer = new StreamWriter(output, new UTF8Encoding(false));
            writer.Write(xml);
        }
        return stream.ToArray();
    }
}
