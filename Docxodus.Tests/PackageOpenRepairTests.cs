using System;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Package defects Word opens but the Open XML SDK rejects, repaired once on the way in so every
/// front door (session, comparison, conversion) opens the package and every save writes it back
/// (issue #1384): a relationship to a part with no content type or no zip entry, a part a known
/// relationship reaches but typed only as generic XML, and two relationships sharing an Id.
/// </summary>
public class PackageOpenRepairTests
{
    private const string FontEntry = "word/fonts/ExampleSans-regular.ttf";
    private const string ImageEntry = "word/media/image1.jpeg";
    private const string ThumbnailEntry = "docProps/thumbnail.jpeg";
    private const string SiblingTarget = "other.docx";
    private const string LinkTarget = "https://link.example.com/";
    private const string TtfContentType = "application/x-font-ttf";
    private const string GlossaryContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";
    private const string CustomXmlPropsContentType = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
    private static readonly byte[] FontBytes = Enumerable.Range(0, 64).Select(i => (byte)(i * 7)).ToArray();
    private static readonly byte[] ImageBytes = Enumerable.Range(0, 32).Select(i => (byte)(255 - i)).ToArray();
    private static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace PackageRels = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>A well-formed package carrying every part kind the defects below damage: an embedded
    /// font, a glossary document, a custom XML data store with its properties, a theme, a header and a
    /// footer, a JPEG image and thumbnail, an embedded OLE object, an external hyperlink, and an
    /// internal hyperlink to a sibling file (a relationship the SDK reads as a reference, not a part).</summary>
    private static byte[] Pristine(bool siblingLink = false)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddHyperlinkRelationship(new Uri(LinkTarget), true, "rIdLink");
            if (siblingLink)
                main.AddHyperlinkRelationship(new Uri(SiblingTarget, UriKind.Relative), false, "rIdSibling");
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Alpha paragraph."))),
                new Paragraph(new Hyperlink(new Run(new Text("Linked text"))) { Id = "rIdLink" }),
                new Paragraph(new Hyperlink(new Run(new Text("Sibling file"))) { Id = siblingLink ? "rIdSibling" : "rIdLink" }),
                new SectionProperties(
                    new HeaderReference { Type = HeaderFooterValues.Default, Id = "rIdHdr" },
                    new FooterReference { Type = HeaderFooterValues.Default, Id = "rIdFtr" })));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.AddNewPart<ThemePart>("rIdTheme").Theme = new DocumentFormat.OpenXml.Drawing.Theme { Name = "Example" };
            main.AddNewPart<HeaderPart>("rIdHdr").Header = new Header(new Paragraph(new Run(new Text("Header text"))));
            main.AddNewPart<FooterPart>("rIdFtr").Footer = new Footer(new Paragraph(new Run(new Text("Footer text"))));
            var fonts = main.AddNewPart<FontTablePart>();
            Feed(fonts.AddFontPart(FontPartType.FontTtf, "rIdEmb1"), FontBytes);
            fonts.Fonts = new Fonts(new Font(new EmbedRegularFont { Id = "rIdEmb1" }) { Name = "Example Sans" });
            main.AddNewPart<GlossaryDocumentPart>().GlossaryDocument = new GlossaryDocument(new DocParts());
            var data = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            Feed(data, Encoding.UTF8.GetBytes("<root xmlns=\"urn:example:data\"/>"));
            data.AddNewPart<CustomXmlPropertiesPart>().DataStoreItem =
                new DocumentFormat.OpenXml.CustomXmlDataProperties.DataStoreItem { ItemId = "{00000000-0000-0000-0000-000000000001}" };
            Feed(main.AddImagePart(ImagePartType.Jpeg, "rIdImg"), ImageBytes);
            Feed(main.AddEmbeddedObjectPart("application/vnd.openxmlformats-officedocument.oleObject"), ImageBytes);
            Feed(doc.AddThumbnailPart(ThumbnailPartType.Jpeg), ImageBytes);
        }
        return Rewrite(ms.ToArray(), zip =>
        {
            Rename(zip, EntryNamed(zip, ".ttf"), FontEntry);
            Rename(zip, zip.Entries.Single(e => e.FullName.StartsWith("media/", StringComparison.Ordinal)).FullName, ImageEntry);
            Rename(zip, zip.Entries.Single(e => e.FullName.Contains("thumbnail", StringComparison.OrdinalIgnoreCase)).FullName, ThumbnailEntry);
        });
    }

    private static void Feed(OpenXmlPart part, byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        part.FeedData(stream);
    }

    private static byte[] Rewrite(byte[] package, Action<ZipArchive> edit)
    {
        using var ms = new MemoryStream();
        ms.Write(package);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
            edit(zip);
        return ms.ToArray();
    }

    private static string EntryNamed(ZipArchive zip, string suffix) =>
        zip.Entries.Single(e => e.FullName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)).FullName;

    /// <summary>Renames a part, keeping its relationships and its content type (as an Override).</summary>
    private static void Rename(ZipArchive zip, string oldName, string newName)
    {
        var entry = zip.GetEntry(oldName)!;
        using var copy = new MemoryStream();
        using (var input = entry.Open())
            input.CopyTo(copy);
        entry.Delete();
        using (var output = zip.CreateEntry(newName).Open())
            output.Write(copy.ToArray());
        foreach (var rels in zip.Entries.Where(e => e.FullName.EndsWith(".rels", StringComparison.Ordinal)).Select(e => e.FullName).ToList())
            EditXml(zip, rels, xml =>
            {
                foreach (var target in xml.Descendants().Attributes("Target").Where(t => t.Value.EndsWith(Path.GetFileName(oldName), StringComparison.Ordinal)))
                    target.Value = "/" + newName;
            });
        EditXml(zip, "[Content_Types].xml", types =>
        {
            var contentType = types.Root!.Elements().First(e => (string?)e.Attribute("PartName") == "/" + oldName
                || (string?)e.Attribute("Extension") == Path.GetExtension(oldName).TrimStart('.')).Attribute("ContentType")!.Value;
            types.Root.Elements().Where(e => (string?)e.Attribute("PartName") == "/" + oldName).Remove();
            types.Root.Add(new XElement(ContentTypes + "Override",
                new XAttribute("PartName", "/" + newName), new XAttribute("ContentType", contentType)));
        });
    }

    private static void EditXml(ZipArchive zip, string name, Action<XDocument> edit)
    {
        var entry = zip.GetEntry(name)!;
        XDocument xml;
        using (var stream = entry.Open())
            xml = XDocument.Load(stream);
        edit(xml);
        entry.Delete();
        using var output = zip.CreateEntry(name).Open();
        xml.Save(output);
    }

    /// <summary>Removes every content type that would type <paramref name="partName"/>.</summary>
    private static byte[] Untyped(byte[] package, string partName, string? extension = null) =>
        Rewrite(package, zip => EditXml(zip, "[Content_Types].xml", types =>
            types.Root!.Elements().Where(e =>
                    (string?)e.Attribute("PartName") == partName
                    || (extension is not null && (string?)e.Attribute("Extension") == extension))
                .Remove()));

    private static byte[] UntypedFont() => Untyped(Pristine(), "/" + FontEntry, "ttf");

    private static byte[] AbsentFont() => Rewrite(UntypedFont(), zip => zip.GetEntry(FontEntry)!.Delete());

    /// <summary>The part typed as generic XML, as a writer that types every <c>.xml</c> entry
    /// <c>application/xml</c> leaves it.</summary>
    private static byte[] GenericXml(string partSuffix) => Rewrite(Pristine(), zip =>
    {
        var partName = "/" + EntryNamed(zip, partSuffix);
        EditXml(zip, "[Content_Types].xml", types =>
            types.Root!.Elements(ContentTypes + "Override").Single(o => (string?)o.Attribute("PartName") == partName)
                .SetAttributeValue("ContentType", "application/xml"));
    });

    /// <summary>The theme relationship given the hyperlink's Id and placed before it (as a writer
    /// that numbers new relationships without reading the existing ones leaves it) or after it.</summary>
    private static byte[] DuplicateIds(bool themeFirst = true) => Rewrite(Pristine(), zip => EditXml(zip, "word/_rels/document.xml.rels", rels =>
    {
        var theme = rels.Root!.Elements().Single(r => (string?)r.Attribute("Id") == "rIdTheme");
        theme.Remove();
        theme.SetAttributeValue("Id", "rIdLink");
        if (themeFirst)
            rels.Root.AddFirst(theme);
        else
            rels.Root.Add(theme);
    }));

    /// <summary>The header and footer relationships sharing one Id, which both section references name.</summary>
    private static byte[] HeaderAndFooterSharingAnId() => Rewrite(Pristine(), zip =>
    {
        EditXml(zip, "word/_rels/document.xml.rels", rels =>
            rels.Root!.Elements().Single(r => (string?)r.Attribute("Id") == "rIdFtr").SetAttributeValue("Id", "rIdHdr"));
        var entry = zip.GetEntry("word/document.xml")!;
        string xml;
        using (var reader = new StreamReader(entry.Open()))
            xml = reader.ReadToEnd();
        entry.Delete();
        using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open());
        writer.Write(xml.Replace("\"rIdFtr\"", "\"rIdHdr\""));
    });

    /// <summary>The pristine package with the JPEG shown by a DrawingML picture and a VML picture.</summary>
    private static byte[] WithPicturesOfTheImage() => Relaid(Pristine(), "word/document.xml", "<w:body>",
        "<w:body><w:p><w:r><w:drawing>" +
        "<wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"><wp:extent cx=\"914400\" cy=\"914400\"/>" +
        "<wp:docPr id=\"1\" name=\"Picture 1\"/><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
        "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:nvPicPr><pic:cNvPr id=\"1\" name=\"Picture 1\"/><pic:cNvPicPr/></pic:nvPicPr>" +
        "<pic:blipFill><a:blip r:embed=\"rIdImg\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
        "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\"rect\"/></pic:spPr>" +
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>" +
        "<w:p><w:r><w:pict><v:shape id=\"s1\" style=\"width:10pt;height:10pt\" xmlns:v=\"urn:schemas-microsoft-com:vml\">" +
        "<v:imagedata r:id=\"rIdImg\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/></v:shape></w:pict></w:r></w:p>");

    private static byte[] Without(byte[] package, string entry) => Rewrite(package, zip => zip.GetEntry(entry)!.Delete());

    private static string HeaderEntry => EntryNamed(new ZipArchive(new MemoryStream(Pristine())), "header1.xml");

    private static string FooterEntry => EntryNamed(new ZipArchive(new MemoryStream(Pristine())), "footer1.xml");

    private static byte[] WithAddedParagraph(byte[] package) => Rewrite(package, zip =>
    {
        var entry = zip.GetEntry("word/document.xml")!;
        string xml;
        using (var reader = new StreamReader(entry.Open()))
            xml = reader.ReadToEnd();
        entry.Delete();
        var at = xml.IndexOf("<w:body>", StringComparison.Ordinal) + "<w:body>".Length;
        using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open());
        writer.Write(xml.Insert(at, "<w:p><w:r><w:t>An added paragraph.</w:t></w:r></w:p>"));
    });

    private static (string ContentType, byte[] Bytes)? Part(byte[] package, string partName)
    {
        using var package2 = Package.Open(new MemoryStream(package), FileMode.Open, FileAccess.Read);
        var uri = new Uri(partName, UriKind.Relative);
        if (!package2.PartExists(uri))
            return null;
        var part = package2.GetPart(uri);
        using var stream = part.GetStream();
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return (part.ContentType, copy.ToArray());
    }

    private static string? ContentTypeOverride(byte[] package, string partName)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var stream = zip.GetEntry("[Content_Types].xml")!.Open();
        return (string?)XDocument.Load(stream).Root!.Elements(ContentTypes + "Override")
            .FirstOrDefault(o => (string?)o.Attribute("PartName") == partName)?.Attribute("ContentType");
    }

    private static string[] RelationshipTargets(byte[] package, string relsEntry)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var stream = zip.GetEntry(relsEntry)!.Open();
        return XDocument.Load(stream).Root!.Elements(PackageRels + "Relationship")
            .Select(r => (string)r.Attribute("Target")!).ToArray();
    }

    private static string[] RelationshipIds(byte[] package, string relsEntry)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var stream = zip.GetEntry(relsEntry)!.Open();
        return XDocument.Load(stream).Root!.Elements(PackageRels + "Relationship")
            .Select(r => (string)r.Attribute("Id")!).ToArray();
    }

    public static TheoryData<string> Defects => new()
    {
        "untyped font", "absent font", "untyped image", "untyped thumbnail", "absent referenced image",
        "generic glossary", "generic custom XML properties", "duplicate ids",
    };

    private static byte[] Defective(string defect) => defect switch
    {
        "untyped font" => UntypedFont(),
        "absent font" => AbsentFont(),
        "untyped image" => Untyped(Pristine(), "/" + ImageEntry),
        "untyped thumbnail" => Untyped(Pristine(), "/" + ThumbnailEntry),
        "absent referenced image" => Without(WithPicturesOfTheImage(), ImageEntry),
        "generic glossary" => GenericXml("glossary/document.xml"),
        "generic custom XML properties" => GenericXml("itemProps1.xml"),
        "duplicate ids" => DuplicateIds(),
        _ => throw new ArgumentOutOfRangeException(nameof(defect)),
    };

    [Theory]
    [MemberData(nameof(Defects))]
    public void Session_opens_and_saves(string defect)
    {
        using var session = new DocxSession(Defective(defect));
        Assert.Contains("Alpha paragraph.", session.Project().Markdown);

        using var reopened = new DocxSession(session.Save(persistAnchorIds: true));
        Assert.Contains("Alpha paragraph.", reopened.Project().Markdown);
        Assert.NotEmpty(session.Save(persistAnchorIds: false));
    }

    // Not "absent font": dropping the relationship leaves fontTable.xml's w:embedRegular r:id with
    // nothing to resolve, which semantic diff's package preflight rejects on both sides.
    [Theory]
    [InlineData("untyped font")]
    [InlineData("generic glossary")]
    [InlineData("generic custom XML properties")]
    [InlineData("duplicate ids")]
    public void An_unedited_session_reports_no_semantic_changes(string defect)
    {
        using var session = new DocxSession(Defective(defect), new DocxSessionSettings { CaptureInitialProjection = true });

        Assert.Empty(session.GetSemanticChanges().Changes);
    }

    // "untyped font" and "absent font" are controls here: compare opened them before #1384 (it dropped
    // the relationship and so the untyped font, which Compare_keeps_an_untyped_font guards).
    [Theory]
    [InlineData("pristine")]
    [InlineData("untyped font")]
    [InlineData("generic glossary")]
    [InlineData("duplicate ids")]
    public void Verifying_an_edit_blames_the_edit_not_the_input(string defect)
    {
        using var session = new DocxSession(defect == "pristine" ? Pristine() : Defective(defect),
            new DocxSessionSettings { CaptureInitialProjection = true });
        var alpha = System.Text.RegularExpressions.Regex.Match(session.Project().Markdown, @"\{#([^}]+)\} Alpha").Groups[1].Value;
        Assert.True(session.ReplaceText(alpha, "Alpha paragraph, edited.").Success);

        var report = session.VerifyDeliverable();

        Assert.Empty(report.Findings.Where(f => f.BlocksDelivery && f.Disposition == Docxodus.Verification.DeliverableFindingDisposition.New)
            .Select(f => f.Code + ": " + f.Message));
    }

    // A dropped absent font or image leaves its r:embed / w:embedRegular unresolved on both sides, and the
    // semantic comparison refuses packages with an unresolved reference: verification can say that much
    // and no more, and nothing else is charged to the edit.
    [Theory]
    [InlineData("absent font")]
    [InlineData("absent referenced image")]
    public void Verifying_an_edit_after_an_absent_part_was_dropped_reports_only_the_unavailable_comparison(string defect)
    {
        using var session = new DocxSession(Defective(defect), new DocxSessionSettings { CaptureInitialProjection = true });
        var alpha = System.Text.RegularExpressions.Regex.Match(session.Project().Markdown, @"\{#([^}]+)\} Alpha").Groups[1].Value;
        Assert.True(session.ReplaceText(alpha, "Alpha paragraph, edited.").Success);

        var report = session.VerifyDeliverable();

        Assert.Equal(new[] { "delta.semantic_comparison_unavailable" }, report.Findings
            .Where(f => f.BlocksDelivery && f.Disposition == Docxodus.Verification.DeliverableFindingDisposition.New)
            .Select(f => f.Code).Distinct());
    }

    // The receipt identifies the caller's own bytes, so it refuses a source whose manifest has error
    // findings (an untyped part, a duplicate Id) as it always has; a generically typed part is not one.
    [Fact]
    public void An_unedited_session_receipt_reports_no_package_change()
    {
        using var session = new DocxSession(Defective("generic glossary"),
            new DocxSessionSettings { CaptureInitialProjection = true, CaptureDeliveryEvidence = true });

        var bundle = session.BuildDeliveryReceipt();

        Assert.True(bundle.Verification.IsValid, string.Join("; ", bundle.Verification.Findings));
        var receipt = Encoding.UTF8.GetString(bundle.ArtifactBytes.Single(a => a.Key.Contains("receipt", StringComparison.Ordinal)).Value);
        Assert.Contains("\"packageChanges\":[]", receipt);
        Assert.Contains("\"changeCount\":0", receipt);
    }

    [Theory]
    [MemberData(nameof(Defects))]
    public void Compare_opens_both_sides(string defect)
    {
        var original = Defective(defect);

        var redline = DocxCompare.Compare(
            new WmlDocument("original.docx", original),
            new WmlDocument("revised.docx", WithAddedParagraph(original)));

        Assert.Contains("An added paragraph.", new DocxSession(redline.DocumentByteArray).Project().Markdown);
    }

    [Theory]
    [MemberData(nameof(Defects))]
    public void Html_conversion_opens(string defect)
    {
        var package = Defective(defect);
        using var session = new DocxSession(package);
        var anchor = System.Text.RegularExpressions.Regex.Match(session.Project().Markdown, @"\{#([^}]+)\} Alpha").Groups[1].Value;

        Assert.Contains("Alpha paragraph.", WmlToHtmlConverter.ConvertToHtml(
            new WmlDocument("defect.docx", package), new WmlToHtmlConverterSettings()).Value);
        Assert.Contains("Alpha paragraph.", Docxodus.Internal.HtmlConversionOps.ConvertToHtml(
            package, new Docxodus.Internal.HtmlConversionOptions()));
        Assert.Contains("Alpha paragraph.", Docxodus.Internal.HtmlConversionOps.ConvertToHtml(
            package, new Docxodus.Internal.HtmlConversionOptions { RenderHeadersAndFooters = true }));
        Assert.Contains("Alpha paragraph.", Docxodus.Internal.HtmlConversionOps.ConvertToHtml(
            package, new Docxodus.Internal.HtmlConversionOptions { PaginationMode = 1, RenderHeadersAndFooters = true }));
        Assert.Contains("Alpha paragraph.", Docxodus.Internal.HtmlConversionOps.RenderBlockHtml(
            package, anchor, new Docxodus.Internal.HtmlConversionOptions()));
    }

    [Theory]
    [MemberData(nameof(Defects))]
    public void The_compatibility_probe_opens(string defect) =>
        Assert.NotNull(DocxDiffCompatibility.Inspect(Defective(defect)));

    [Fact]
    public void A_picture_of_an_absent_image_lists_as_broken()
    {
        using var session = new DocxSession(Defective("absent referenced image"));

        var images = session.ListImages();

        Assert.Equal(2, images.Count);
        Assert.All(images, image => Assert.True(image.IsBroken));
    }

    [Fact]
    public void Compare_keeps_an_untyped_font()
    {
        var original = UntypedFont();

        var redline = DocxCompare.Compare(
            new WmlDocument("original.docx", original),
            new WmlDocument("revised.docx", WithAddedParagraph(original)));

        Assert.Equal(FontBytes, Part(redline.DocumentByteArray, "/" + FontEntry)?.Bytes);
    }

    [Fact]
    public void An_internal_hyperlink_to_a_file_outside_the_package_is_kept()
    {
        var pristine = Pristine(siblingLink: true);
        using var session = new DocxSession(pristine);
        var redline = DocxCompare.Compare(
            new WmlDocument("original.docx", pristine),
            new WmlDocument("revised.docx", WithAddedParagraph(pristine)));
        using var compared = new DocxSession(redline.DocumentByteArray);

        Assert.Contains(session.ListHyperlinks(), l => l.Target == SiblingTarget);
        Assert.Contains(compared.ListHyperlinks(), l => l.Target == SiblingTarget);
    }

    [Theory]
    [InlineData(ImageEntry, "image/jpeg")]
    [InlineData(ThumbnailEntry, "image/jpeg")]
    public void An_untyped_image_keeps_its_bytes_and_gains_the_type_its_extension_names(string entry, string contentType)
    {
        using var session = new DocxSession(Untyped(Pristine(), "/" + entry));

        var image = Part(session.Save(), "/" + entry);

        Assert.Equal(contentType, image?.ContentType);
        Assert.Equal(ImageBytes, image?.Bytes);
    }

    [Fact]
    public void An_untyped_part_nothing_types_is_left_for_the_open_to_reject()
    {
        var oleEntry = EntryNamed(new ZipArchive(new MemoryStream(Pristine())), ".bin");
        var package = Untyped(Pristine(), "/" + oleEntry, "bin");

        var read = StrictOoxmlNormalizer.NormalizeForReading(package);

        Assert.Contains(EntryNamed(new ZipArchive(new MemoryStream(read)), ".bin"), oleEntry);
        Assert.Contains(RelationshipTargets(read, "word/_rels/document.xml.rels"), t => t.EndsWith(Path.GetFileName(oleEntry), StringComparison.Ordinal));
        Assert.ThrowsAny<Exception>(() =>
        {
            using var session = new DocxSession(package);
            session.Save();
        });
    }

    // A header, footer or any kind other than a font, image or thumbnail: the package does not say what
    // a document missing that part should become, so the open fails as it always has.
    public static TheoryData<string> PartsAbsentWithoutARepair => new() { "header1.xml", "footer1.xml", "glossary/document.xml" };

    [Theory]
    [MemberData(nameof(PartsAbsentWithoutARepair))]
    public void An_absent_part_of_another_kind_is_left_for_the_open_to_reject(string partSuffix)
    {
        var package = Without(Pristine(), EntryNamed(new ZipArchive(new MemoryStream(Pristine())), partSuffix));

        Assert.Same(package, StrictOoxmlNormalizer.NormalizeForReading(package));
        Assert.ThrowsAny<Exception>(() =>
        {
            using var session = new DocxSession(package);
            session.Save();
        });
    }

    [Fact]
    public void Two_relationships_referenced_by_one_id_are_left_for_the_open_to_reject()
    {
        var package = HeaderAndFooterSharingAnId();

        Assert.Same(package, StrictOoxmlNormalizer.NormalizeForReading(package));
        Assert.Throws<System.Xml.XmlException>(() => new DocxSession(package));
    }

    // Dropping the absent member would hand the shared Id, and so the other story's reference, to the
    // survivor: LibreOffice then shows the footer's text as the header too.
    [Fact]
    public void A_shared_id_left_alone_is_not_resolved_by_dropping_an_absent_member()
    {
        var package = Without(HeaderAndFooterSharingAnId(), HeaderEntry);

        Assert.Same(package, StrictOoxmlNormalizer.NormalizeForReading(package));
    }

    [Fact]
    public void Two_images_sharing_an_id_one_of_them_absent_are_left()
    {
        var package = Without(Rewrite(WithPicturesOfTheImage(), zip =>
        {
            var copy = zip.CreateEntry("word/media/image2.jpeg");
            using (var output = copy.Open())
                output.Write(ImageBytes);
            EditXml(zip, "word/_rels/document.xml.rels", rels => rels.Root!.Add(new XElement(PackageRels + "Relationship",
                new XAttribute("Id", "rIdImg"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                new XAttribute("Target", "/word/media/image2.jpeg"))));
            EditXml(zip, "[Content_Types].xml", types => types.Root!.Add(new XElement(ContentTypes + "Override",
                new XAttribute("PartName", "/word/media/image2.jpeg"), new XAttribute("ContentType", "image/jpeg"))));
        }), ImageEntry);

        Assert.Same(package, StrictOoxmlNormalizer.NormalizeForReading(package));
    }

    // Targets System.IO.Packaging and the SDK resolve to the present image, however spelled, and the
    // image stored whole or interleaved (name/[0].piece ... name/[n].last.piece, which OPC allows).
    public static TheoryData<string, string, int> PresentImageSpellings => new()
    {
        { "/word/media/image1.jpeg", ImageEntry, 0 }, { "/word/media/image1.jpeg#frag", ImageEntry, 0 },
        { "/word/media/image1.jpeg?x", ImageEntry, 0 }, { "/word/media/image1.jpeg#", ImageEntry, 0 },
        { " /word/media/image1.jpeg", ImageEntry, 0 }, { "/word/media/image1.jpeg ", ImageEntry, 0 },
        { "./media/image1.jpeg", ImageEntry, 0 }, { "../word/media/image1.jpeg", ImageEntry, 0 },
        { "media\\image1.jpeg", ImageEntry, 0 }, { "/word/media/image1.jpeg", ImageEntry, 1 },
        { "/word/media/image1.jpeg", ImageEntry, 2 },
        // The entry's name differs from the target only in case, which the packaging layer ignores.
        { "/word/media/image1.jpeg", "word/media/IMAGE1.JPEG", 0 }, { "/word/media/i.jpeg", "Word/Media/i.jpeg", 0 },
        { "/word/media/image1.jpeg", "WORD/MEDIA/IMAGE1.JPEG", -2 },
    };

    /// <summary>The image of <see cref="WithPicturesOfTheImage"/> moved to <paramref name="entry"/>, whole
    /// or in <paramref name="pieces"/> interleaved pieces (negative for upper-case piece names), its
    /// Override following it, and the picture relationship's target set to <paramref name="target"/>.</summary>
    private static byte[] WithTheImageAt(string target, string entry, int pieces) => Rewrite(WithPicturesOfTheImage(), zip =>
    {
        EditXml(zip, "word/_rels/document.xml.rels", rels =>
            rels.Root!.Elements().Single(r => (string?)r.Attribute("Id") == "rIdImg").SetAttributeValue("Target", target));
        EditXml(zip, "[Content_Types].xml", types =>
            types.Root!.Elements(ContentTypes + "Override").Single(o => (string?)o.Attribute("PartName") == "/" + ImageEntry)
                .SetAttributeValue("PartName", "/" + entry));
        zip.GetEntry(ImageEntry)!.Delete();
        var split = pieces == 0 ? Array.Empty<byte[]>() : Math.Abs(pieces) == 1 ? new[] { ImageBytes } : new[] { ImageBytes[..10], ImageBytes[10..] };
        if (split.Length == 0)
        {
            using var whole = zip.CreateEntry(entry).Open();
            whole.Write(ImageBytes);
        }
        for (var i = 0; i < split.Length; i++)
        {
            var name = $"[{i}]{(i == split.Length - 1 ? ".last" : "")}.piece";
            using var piece = zip.CreateEntry($"{entry}/{(pieces < 0 ? name.ToUpperInvariant() : name)}").Open();
            piece.Write(split[i]);
        }
    });

    [Theory]
    [MemberData(nameof(PresentImageSpellings))]
    public void An_image_the_package_holds_is_never_dropped_however_its_target_is_spelled(string target, string entry, int pieces)
    {
        var package = WithTheImageAt(target, entry, pieces);

        Assert.Same(package, StrictOoxmlNormalizer.NormalizeForReading(package));
        using var session = new DocxSession(package);
        Assert.Equal(ImageBytes, Part(session.Save(), "/" + entry)?.Bytes);
    }

    // The packaging layer matches an Override to the entry's own spelling, case and all, so the added
    // Override must name the entry, and an Override that differs from it in case does not type it.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void An_untyped_image_whose_entry_differs_in_case_is_typed_under_the_entry_name(bool caseDifferingOverride)
    {
        var package = WithTheImageAt("/word/media/image1.jpeg", "word/media/IMAGE1.JPEG", 0);
        package = caseDifferingOverride
            ? Rewrite(package, zip => EditXml(zip, "[Content_Types].xml", types =>
                types.Root!.Elements(ContentTypes + "Override").Single(o => (string?)o.Attribute("PartName") == "/word/media/IMAGE1.JPEG")
                    .SetAttributeValue("PartName", "/word/media/image1.jpeg")))
            : Untyped(package, "/word/media/IMAGE1.JPEG");

        var read = StrictOoxmlNormalizer.NormalizeForReading(package);

        Assert.Equal("image/jpeg", ContentTypeOverride(read, "/word/media/IMAGE1.JPEG"));
        Assert.Null(ContentTypeOverride(read, "/word/media/image1.jpeg")); // rewritten, not joined by a second
        using var session = new DocxSession(package);
        Assert.Equal(ImageBytes, Part(session.Save(), "/word/media/IMAGE1.JPEG")?.Bytes);
    }

    // A part name more than one zip entry forms is ambiguous: the repair must neither type nor drop
    // it, so every front door behaves as on main.
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Case_variant_entries_of_one_part_are_left_alone(bool untypedFirst)
    {
        var package = Rewrite(WithTheImageAt("media/image1.jpeg", "word/media/IMAGE1.JPEG", 0), zip =>
        {
            var typed = zip.GetEntry("word/media/IMAGE1.JPEG")!;
            typed.Delete();
            foreach (var name in untypedFirst ? new[] { "word/media/image1.jpeg", "word/media/IMAGE1.JPEG" } : new[] { "word/media/IMAGE1.JPEG", "word/media/image1.jpeg" })
            {
                using var output = zip.CreateEntry(name).Open();
                output.Write(ImageBytes);
            }
        });

        Assert.Same(package, StrictOoxmlNormalizer.NormalizeForReading(package));
        AssertEveryFrontDoorOpens(package, sessionOpens: true);
    }

    [Fact]
    public void A_whole_entry_beside_interleaved_pieces_of_one_part_is_left_alone()
    {
        var package = Rewrite(Untyped(WithPicturesOfTheImage(), "/" + ImageEntry), zip =>
        {
            using var piece = zip.CreateEntry(ImageEntry + "/[0].last.piece").Open();
            piece.Write(ImageBytes);
        });

        Assert.Same(package, StrictOoxmlNormalizer.NormalizeForReading(package));
        AssertEveryFrontDoorOpens(package, sessionOpens: false);
    }

    /// <summary>Compare, accept and HTML conversion of a WmlDocument open the package, as on main; the
    /// session opens and saves it only when <paramref name="sessionOpens"/>, and otherwise fails as on main.</summary>
    private static void AssertEveryFrontDoorOpens(byte[] package, bool sessionOpens)
    {
        var document = new WmlDocument("ambiguous.docx", package);
        Assert.NotNull(DocxCompare.Compare(document, new WmlDocument("revised.docx", WithAddedParagraph(package))));
        Assert.NotNull(RevisionProcessor.AcceptRevisions(document));
        Assert.Contains("Alpha paragraph.", WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings()).Value);
        void Session()
        {
            using var session = new DocxSession(package);
            session.Save();
        }
        if (sessionOpens)
            Session();
        else
            Assert.ThrowsAny<Exception>(Session);
    }

    // System.IO.Packaging reads a TargetMode with surrounding whitespace as that mode; any other value
    // (another case, an unknown word) it refuses. Only an absent or "Internal" mode is internal.
    [Theory]
    [InlineData(" External")]
    [InlineData("External ")]
    [InlineData("External&#9;")]
    [InlineData("&#10;External")]
    [InlineData("EXTERNAL")]
    [InlineData("external")]
    [InlineData("Linked")]
    public void A_relationship_whose_target_mode_is_not_plainly_internal_is_never_dropped(string mode)
    {
        var package = Relaid(Pristine(), "word/_rels/document.xml.rels", "</Relationships>",
            "<Relationship Id=\"rIdLinkedPicture\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" " +
            $"Target=\"../pictures/linked.jpeg\" TargetMode=\"{mode}\"/></Relationships>");

        var read = StrictOoxmlNormalizer.NormalizeForReading(package);

        Assert.Contains(RelationshipTargets(read, "word/_rels/document.xml.rels"), t => t == "../pictures/linked.jpeg");
    }

    [Theory]
    [InlineData("media/")]
    [InlineData(".")]
    public void A_target_that_names_no_part_is_never_dropped(string target)
    {
        var package = Without(WithTheImageAt(target, ImageEntry, 0), ImageEntry);

        Assert.Same(package, StrictOoxmlNormalizer.NormalizeForReading(package));
    }

    [Fact]
    public void An_absent_image_reached_through_a_dot_segment_is_dropped()
    {
        var package = Without(Rewrite(WithPicturesOfTheImage(), zip => EditXml(zip, "word/_rels/document.xml.rels", rels =>
            rels.Root!.Elements().Single(r => (string?)r.Attribute("Id") == "rIdImg").SetAttributeValue("Target", "../word/media/image1.jpeg"))), ImageEntry);

        var read = StrictOoxmlNormalizer.NormalizeForReading(package);

        Assert.DoesNotContain(RelationshipTargets(read, "word/_rels/document.xml.rels"), t => t.Contains("image1.jpeg", StringComparison.Ordinal));
    }

    // Every well-formed document in the corpus is read as the same bytes: a repair that fires here is
    // a lookup that disagrees with the packaging layer.
    [Fact]
    public void Every_corpus_document_is_read_as_the_same_bytes()
    {
        var repaired = new List<string>();
        foreach (var file in Directory.EnumerateFiles("../../../../TestFiles", "*.docx", SearchOption.AllDirectories))
        {
            var bytes = File.ReadAllBytes(file);
            byte[] read;
            try
            {
                read = PackageOpenRepair.Repair(bytes);
            }
            catch (InvalidDataException)
            {
                continue; // not a zip: the read step leaves it to the open
            }
            if (!ReferenceEquals(bytes, read))
                repaired.Add(file);
        }

        Assert.Empty(repaired);
    }

    [Fact]
    public void An_absent_thumbnail_is_dropped_and_the_session_saves()
    {
        using var session = new DocxSession(Without(Pristine(), ThumbnailEntry));

        Assert.DoesNotContain(RelationshipTargets(session.Save(), "_rels/.rels"), t => t.Contains("thumbnail", StringComparison.Ordinal));
    }

    [Fact]
    public void Dropping_an_absent_image_drops_its_override_too()
    {
        var read = StrictOoxmlNormalizer.NormalizeForReading(Defective("absent referenced image"));

        Assert.Null(ContentTypeOverride(read, "/" + ImageEntry));
    }

    [Fact]
    public void An_untyped_part_is_typed_under_its_entry_name()
    {
        // The entry holds an escaped space and the relationship spells it raw.
        const string escaped = "word/media/image%201.jpeg";
        var package = Untyped(Rewrite(Pristine(), zip =>
        {
            Rename(zip, ImageEntry, escaped);
            EditXml(zip, "word/_rels/document.xml.rels", rels =>
                rels.Root!.Elements().Single(r => (string?)r.Attribute("Id") == "rIdImg").SetAttributeValue("Target", "/word/media/image 1.jpeg"));
        }), "/" + escaped);

        var read = StrictOoxmlNormalizer.NormalizeForReading(package);

        Assert.Equal("image/jpeg", ContentTypeOverride(read, "/" + escaped));
        Assert.Contains(RelationshipTargets(read, "word/_rels/document.xml.rels"), t => t == "/word/media/image 1.jpeg");
    }

    [Fact]
    public void A_malformed_content_types_part_leaves_the_package_for_the_open_to_reject()
    {
        var package = Relaid(UntypedFont(), "[Content_Types].xml", "</Types>", "<Default Extension=\"x\"</Types>");

        Assert.Same(package, StrictOoxmlNormalizer.NormalizeForReading(package));
    }

    [Fact]
    public void The_read_step_keeps_an_internal_attached_template_outside_the_package()
    {
        var package = Rewrite(Pristine(), zip =>
        {
            var rels = zip.CreateEntry("word/_rels/settings.xml.rels");
            using var writer = new StreamWriter(rels.Open());
            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rIdTemplate\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate\" Target=\"../../Templates/Example.dotm\"/></Relationships>");
        });

        Assert.Same(package, StrictOoxmlNormalizer.NormalizeForReading(package));
    }

    /// <summary>The package thumbnail relationship given the Id of a custom package relationship, ahead of it.</summary>
    private static byte[] ThumbnailSharingAnIdWithACustomRelationship() => Rewrite(Pristine(), zip =>
    {
        using (var data = zip.CreateEntry("docProps/custom-data.xml").Open())
            data.Write(Encoding.UTF8.GetBytes("<data xmlns=\"urn:example:custom\"/>"));
        EditXml(zip, "[Content_Types].xml", types => types.Root!.Add(new XElement(ContentTypes + "Override",
            new XAttribute("PartName", "/docProps/custom-data.xml"), new XAttribute("ContentType", "application/xml"))));
        EditXml(zip, "_rels/.rels", rels =>
        {
            var thumbnail = rels.Root!.Elements().Single(r => ((string?)r.Attribute("Type"))!.EndsWith("thumbnail", StringComparison.Ordinal));
            thumbnail.Remove();
            thumbnail.SetAttributeValue("Id", "rIdCustom");
            rels.Root.AddFirst(thumbnail);
            rels.Root.Add(new XElement(PackageRels + "Relationship", new XAttribute("Id", "rIdCustom"),
                new XAttribute("Type", "http://example.com/relationships/custom"), new XAttribute("Target", "docProps/custom-data.xml")));
        });
    });

    [Fact]
    public void A_package_level_shared_id_opens_on_every_front_door()
    {
        var package = ThumbnailSharingAnIdWithACustomRelationship();
        var document = new WmlDocument("shared.docx", package);

        using (var session = new DocxSession(package))
            Assert.Contains("Alpha paragraph.", session.Project().Markdown);
        Assert.NotNull(DocxCompare.Compare(document, new WmlDocument("revised.docx", WithAddedParagraph(package))));
        Assert.NotNull(RevisionProcessor.AcceptRevisions(document));
        Assert.Contains("Alpha paragraph.", WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings()).Value);
    }

    [Fact]
    public void An_untyped_font_keeps_its_bytes_and_gains_the_font_content_type()
    {
        using var session = new DocxSession(UntypedFont());

        var font = Part(session.Save(), "/" + FontEntry);

        Assert.NotNull(font);
        Assert.Equal(TtfContentType, font.Value.ContentType);
        Assert.Equal(FontBytes, font.Value.Bytes);
    }

    [Fact]
    public void A_relationship_to_an_absent_font_is_dropped()
    {
        using var session = new DocxSession(AbsentFont());

        Assert.Empty(RelationshipIds(session.Save(), "word/_rels/fontTable.xml.rels"));
    }

    [Theory]
    [InlineData("glossary/document.xml", GlossaryContentType)]
    [InlineData("itemProps1.xml", CustomXmlPropsContentType)]
    public void A_generically_typed_part_gains_the_content_type_its_relationship_requires(string partSuffix, string contentType)
    {
        var package = GenericXml(partSuffix);
        var partName = "/" + EntryNamed(new ZipArchive(new MemoryStream(package)), partSuffix);
        using var session = new DocxSession(package);

        Assert.Equal(contentType, Part(session.Save(), partName)?.ContentType);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Duplicate_ids_stay_on_the_relationship_looked_up_by_id(bool themeFirst)
    {
        using var session = new DocxSession(DuplicateIds(themeFirst));

        var saved = session.Save();
        var ids = RelationshipIds(saved, "word/_rels/document.xml.rels");
        Assert.Equal(ids.Length, ids.Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(LinkTarget, session.ListHyperlinks().First(l => l.Text == "Linked text").Target);
        using var reopened = WordprocessingDocument.Open(new MemoryStream(saved), false);
        Assert.NotNull(reopened.MainDocumentPart!.ThemePart);
    }

    [Fact]
    public void The_fixed_content_types_are_the_ones_the_sdk_requires()
    {
        using var ms = new MemoryStream();
        using var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        OpenXmlPart[] parts =
        {
            main.AddNewPart<StyleDefinitionsPart>(), main.AddNewPart<StylesWithEffectsPart>(),
            main.AddNewPart<DocumentSettingsPart>(), main.AddNewPart<WebSettingsPart>(), main.AddNewPart<FontTablePart>(),
            main.AddNewPart<NumberingDefinitionsPart>(), main.AddNewPart<FootnotesPart>(), main.AddNewPart<EndnotesPart>(),
            main.AddNewPart<WordprocessingCommentsPart>(), main.AddNewPart<WordprocessingCommentsExPart>(),
            main.AddNewPart<WordprocessingCommentsIdsPart>(), main.AddNewPart<WordprocessingPeoplePart>(),
            main.AddNewPart<HeaderPart>(), main.AddNewPart<FooterPart>(), main.AddNewPart<GlossaryDocumentPart>(),
            main.AddNewPart<ThemePart>(), main.AddCustomXmlPart(CustomXmlPartType.CustomXml).AddNewPart<CustomXmlPropertiesPart>(),
            doc.AddExtendedFilePropertiesPart(), doc.AddCustomFilePropertiesPart(), doc.AddCoreFilePropertiesPart(),
        };

        Assert.Equal(
            parts.ToDictionary(p => p.RelationshipType, p => p.ContentType).OrderBy(e => e.Key),
            PackageOpenRepair.FixedContentTypes.OrderBy(e => e.Key));
    }

    /// <summary>Rewrites a part's text, as a writer with another layout would have written it.</summary>
    private static byte[] Relaid(byte[] package, string entryName, string from, string to) => Rewrite(package, zip =>
    {
        var entry = zip.GetEntry(entryName)!;
        string xml;
        using (var reader = new StreamReader(entry.Open()))
            xml = reader.ReadToEnd();
        entry.Delete();
        using var writer = new StreamWriter(zip.CreateEntry(entryName).Open());
        writer.Write(xml.Replace(from, to));
    });

    [Fact]
    public void A_relationships_part_laid_out_one_attribute_per_line_loses_only_the_dangling_relationship()
    {
        var package = Without(Relaid(Pristine(), "word/_rels/document.xml.rels", "<Relationship ", "<Relationship\n    "), ImageEntry);

        var read = StrictOoxmlNormalizer.NormalizeForReading(package);

        var targets = RelationshipTargets(read, "word/_rels/document.xml.rels");
        Assert.DoesNotContain(targets, t => t.EndsWith("image1.jpeg", StringComparison.Ordinal));
        Assert.Contains(targets, t => t.EndsWith("styles.xml", StringComparison.Ordinal));
        Assert.Equal(RelationshipTargets(package, "word/_rels/document.xml.rels").Length - 1, targets.Length);
    }

    [Fact]
    public void A_pristine_package_laid_out_one_attribute_per_line_is_read_as_the_same_bytes()
    {
        var package = Relaid(Relaid(Pristine(), "[Content_Types].xml", "<Default ", "<Default\n    "),
            "word/_rels/document.xml.rels", "<Relationship ", "<Relationship\n    ");

        Assert.Same(package, StrictOoxmlNormalizer.NormalizeForReading(package));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void A_pristine_package_is_read_as_the_same_bytes(bool siblingLink)
    {
        var pristine = Pristine(siblingLink);

        Assert.Same(pristine, StrictOoxmlNormalizer.NormalizeForReading(pristine));
    }
}
