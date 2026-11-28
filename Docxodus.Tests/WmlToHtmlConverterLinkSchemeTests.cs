using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>A document's link target goes into an <c>href</c> unless a browser would run it as script or open it
/// as inline content: a <c>javascript:</c>, <c>vbscript:</c> or <c>data:</c> URL (issue #1880). The link keeps its
/// text; only the <c>href</c> goes.</summary>
public class WmlToHtmlConverterLinkSchemeTests
{
    private const string Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";

    private static string Run(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    /// <summary>XML for a value, with tabs and line breaks as character references so the parser keeps them.</summary>
    private static string Xml(string value) => value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
        .Replace("\"", "&quot;").Replace("\t", "&#9;").Replace("\n", "&#10;").Replace("\r", "&#13;");

    private static string Complex(string instruction, string text) => $"""
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> {Xml(instruction)} </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>{Run(text)}<w:r><w:fldChar w:fldCharType="end"/></w:r>
        """;

    private static string Simple(string instruction, string text) =>
        $"<w:fldSimple w:instr=\"{Xml(" " + instruction + " ")}\">{Run(text)}</w:fldSimple>";

    private const string Picture =
        "<w:r><w:drawing><wp:inline><wp:extent cx=\"91440\" cy=\"91440\"/>" +
        "<wp:docPr id=\"1\" name=\"Picture 1\"><a:hlinkClick r:id=\"rIdPicture\"/></wp:docPr>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic>" +
        "<pic:nvPicPr><pic:cNvPr id=\"1\" name=\"Picture 1\"/><pic:cNvPicPr/></pic:nvPicPr>" +
        "<pic:blipFill><a:blip r:embed=\"rIdPng\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
        "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"91440\" cy=\"91440\"/></a:xfrm><a:prstGeom prst=\"rect\"/></pic:spPr>" +
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

    /// <summary>A document whose links all go to <paramref name="target"/>: a w:hyperlink, a complex and a simple
    /// HYPERLINK field, and a picture's click link.</summary>
    /// <param name="relationshipTarget">When given, the text the <c>.rels</c> part spells both relationship targets
    /// with, written as is (XML escapes included), for targets a <see cref="Uri"/> would not keep as written.</param>
    private static XElement Render(string target, bool relationshipForms = true, string? relationshipTarget = null)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var quoted = "\"" + target.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
            var body = $"<w:p>{Complex("HYPERLINK " + quoted, "ComplexLink")}</w:p><w:p>{Simple("HYPERLINK " + quoted, "SimpleLink")}</w:p>";
            if (relationshipForms)
            {
                var uri = new Uri(target, UriKind.RelativeOrAbsolute);
                main.AddHyperlinkRelationship(uri, uri.IsAbsoluteUri, "rIdLink");
                main.AddHyperlinkRelationship(uri, uri.IsAbsoluteUri, "rIdPicture");
                var image = main.AddImagePart(ImagePartType.Png, "rIdPng");
                using (var png = new MemoryStream(Convert.FromBase64String(Png)))
                    image.FeedData(png);
                body += $"<w:p><w:hyperlink r:id=\"rIdLink\">{Run("ElementLink")}</w:hyperlink></w:p><w:p>{Picture}</w:p>";
            }
            main.PutXDocument(XDocument.Parse($"""
                <w:document xmlns:w="{W.w}" xmlns:r="{R.r}" xmlns:wp="{WP.wp}" xmlns:a="{A.a}" xmlns:pic="{Pic.pic}"><w:body>{body}
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
                </w:body></w:document>
                """));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"<w:styles xmlns:w=\"{W.w}\"/>"));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
        }
        var bytes = stream.ToArray();
        if (relationshipTarget != null)
            bytes = WithRelationshipTargets(bytes, relationshipTarget);
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("links.docx", bytes),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, ImageHandler = _ => new XElement("img") });
    }

    /// <summary>Respells the targets of the two hyperlink relationships in the package's raw <c>.rels</c> XML.</summary>
    private static byte[] WithRelationshipTargets(byte[] package, string target)
    {
        using var ms = new MemoryStream();
        ms.Write(package);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("word/_rels/document.xml.rels")!;
            string xml;
            using (var reader = new StreamReader(entry.Open()))
                xml = reader.ReadToEnd();
            entry.Delete();
            foreach (var id in new[] { "rIdLink", "rIdPicture" })
                xml = System.Text.RegularExpressions.Regex.Replace(xml,
                    $"(<Relationship [^>]*Id=\"{id}\"[^>]*Target=\")[^\"]*(\")|(<Relationship [^>]*Target=\")[^\"]*(\"[^>]*Id=\"{id}\")",
                    m => m.Groups[1].Success ? m.Groups[1].Value + target + m.Groups[2].Value : m.Groups[3].Value + target + m.Groups[4].Value);
            using var output = zip.CreateEntry("word/_rels/document.xml.rels").Open();
            output.Write(Encoding.UTF8.GetBytes(xml));
        }
        return ms.ToArray();
    }

    private static XElement LinkHolding(XElement html, string text) =>
        Assert.Single(html.Descendants(), e => e.Name.LocalName == "a" && e.Value == text);

    private static XElement PictureLink(XElement html) =>
        Assert.Single(html.Descendants(), e => e.Name.LocalName == "img").Parent!;

    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("JavaScript:alert(1)")]
    [InlineData("vbscript:MsgBox(1)")]
    [InlineData("data:text/html,%3Cb%3Ehi%3C/b%3E")]
    public void AScriptTargetLeavesEveryFormOfLinkWithoutAnHref(string target)
    {
        var html = Render(target);
        foreach (var text in new[] { "ElementLink", "ComplexLink", "SimpleLink" })
            Assert.Null(LinkHolding(html, text).Attribute("href"));
        Assert.Null(PictureLink(html).Attribute("href"));
        Assert.DoesNotContain(html.Descendants().Attributes("href"), href => href.Value.Contains(':') && !href.Value.StartsWith("#"));
    }

    // A browser's URL parser drops surrounding spaces and control characters, and tabs and line breaks anywhere, before
    // it reads the scheme, so these are script targets too. The field forms carry them in their code; the relationship
    // forms below spell them in the .rels part.
    [Theory]
    [InlineData("  javascript:alert(1)")]
    [InlineData("java\tscript:alert(1)")]
    [InlineData("java\nscript:alert(1)")]
    [InlineData("javascript\r:alert(1)")]
    public void AScriptTargetSpelledWithBlanksIsAScriptTarget(string target)
    {
        var html = Render(target, relationshipForms: false);
        Assert.Null(LinkHolding(html, "ComplexLink").Attribute("href"));
        Assert.Null(LinkHolding(html, "SimpleLink").Attribute("href"));
    }

    // The w:hyperlink and picture links read their target from the relationship, which the package reader decodes:
    // blanks spelled as character references, and a percent-encoded letter, which .NET's Uri turns back into the letter.
    // The check reads the target the link would carry, so each is still a script target.
    [Theory]
    [InlineData("java&#9;script:alert(1)")]
    [InlineData("&#10;javascript:alert(1)")]
    [InlineData("javascript&#13;:alert(1)")]
    [InlineData(" javascript:alert(1)")]
    [InlineData("%6Aavascript:alert(1)")]
    public void AScriptTargetSpelledAnyWayInARelationshipIsAScriptTarget(string spelled)
    {
        var html = Render("https://example.test/placeholder", relationshipTarget: spelled);
        Assert.Null(LinkHolding(html, "ElementLink").Attribute("href"));
        Assert.Null(PictureLink(html).Attribute("href"));
    }

    // Control: every other target is the link's, as written.
    [Theory]
    [InlineData("https://example.test/page")]
    [InlineData("mailto:someone@example.test")]
    [InlineData("file:///C:/shared/report.docx")]
    [InlineData("tel:+15550100")]
    [InlineData("custom-app:open?id=7")]
    [InlineData("javascript-guide.html")]
    [InlineData("data/report.html")]
    public void AnyOtherTargetIsTheLinksHref(string target)
    {
        var html = Render(target);
        foreach (var text in new[] { "ElementLink", "ComplexLink", "SimpleLink" })
            Assert.Equal(target, (string?)LinkHolding(html, text).Attribute("href"));
        Assert.Equal(target, (string?)PictureLink(html).Attribute("href"));
    }
}
