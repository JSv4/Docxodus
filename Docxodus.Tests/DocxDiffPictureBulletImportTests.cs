using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A list level can draw a picture as its bullet: <c>w:lvlPicBulletId</c> names a <c>w:numPicBullet</c> in the
/// same numbering part, whose <c>v:imagedata</c> names an image through that part's relationships. A list
/// definition imported from the other document has to arrive with all three, or Word refuses the output
/// (issue #1700).
/// </summary>
public class DocxDiffPictureBulletImportTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";

    // Two 1x1 PNGs.
    private static readonly byte[] Teal = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNgaGgAAAGEAQFWjyAjAAAAAElFTkSuQmCC");
    private static readonly byte[] Red = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGM4IScHAAK2AQU0pnWqAAAAAElFTkSuQmCC");

    /// <summary>One plain paragraph, then one list item per entry of <paramref name="items"/> under numId 1.
    /// With a <paramref name="picture"/>, both levels of the list use it as picture bullet 0, named from the
    /// list definition or, with <paramref name="inLevelOverride"/>, from level overrides of the list instance;
    /// with <paramref name="danglingBullet"/>, the levels name bullet 0 and nothing defines it.</summary>
    private static WmlDocument Doc(string intro, string[] items, byte[]? picture = null, int indent = 720,
        bool danglingBullet = false, bool inLevelOverride = false)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            Write(main, $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:r><w:t>{intro}</w:t></w:r></w:p>" +
                string.Concat(items.Select(text =>
                    "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>" +
                    $"<w:r><w:t>{text}</w:t></w:r></w:p>")) +
                "</w:body></w:document>");
            Write(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles xmlns:w=\"{W}\"/>");
            Write(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings xmlns:w=\"{W}\"/>");
            if (items.Length > 0)
            {
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                var reference = picture != null || danglingBullet ? "<w:lvlPicBulletId w:val=\"0\"/>" : "";
                string Level(int ilvl, bool named) =>
                    $"<w:lvl w:ilvl=\"{ilvl}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"o\"/>" +
                    $"{(named ? reference : "")}<w:lvlJc w:val=\"left\"/>" +
                    $"<w:pPr><w:ind w:left=\"{indent * (ilvl + 1)}\" w:hanging=\"360\"/></w:pPr></w:lvl>";
                var bullet = "";
                if (picture != null)
                {
                    using var image = new MemoryStream(picture);
                    numbering.AddImagePart(ImagePartType.Png, "rIdBullet").FeedData(image);
                    bullet = "<w:numPicBullet w:numPicBulletId=\"0\"><w:pict>" +
                        "<v:shape id=\"_x0000_i1025\" style=\"width:9pt;height:9pt\" o:bullet=\"t\">" +
                        "<v:imagedata r:id=\"rIdBullet\" o:title=\"\"/></v:shape></w:pict></w:numPicBullet>";
                }
                Write(numbering,
                    $"<w:numbering xmlns:w=\"{W}\" xmlns:r=\"{R}\" xmlns:v=\"{V}\" " +
                    "xmlns:o=\"urn:schemas-microsoft-com:office:office\">" + bullet +
                    $"<w:abstractNum w:abstractNumId=\"0\">{Level(0, !inLevelOverride)}{Level(1, !inLevelOverride)}" +
                    "</w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/>" +
                    (inLevelOverride
                        ? string.Concat(new[] { 0, 1 }.Select(ilvl =>
                            $"<w:lvlOverride w:ilvl=\"{ilvl}\">{Level(ilvl, true)}</w:lvlOverride>"))
                        : "") +
                    "</w:num></w:numbering>");
            }
        }
        return new WmlDocument("d.docx", stream.ToArray());

        static void Write(OpenXmlPart part, string xml)
        {
            using var content = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            part.FeedData(content);
        }
    }

    /// <summary>What <paramref name="doc"/> holds about list bullets: its numbering XML, and for each body
    /// paragraph with text in a list, the bytes of the picture its level draws (null for an ordinary bullet).
    /// Fails on a level naming a bullet, or a bullet naming an image, that is not there.</summary>
    private static (XElement Numbering, (string Text, byte[]? Picture)[] Items) Bullets(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var main = wdoc.MainDocumentPart!;
        var part = main.NumberingDefinitionsPart!;
        var numbering = Load(part);

        byte[]? PictureOf(string numId, string ilvl)
        {
            var num = numbering.Elements(W + "num").Single(n => (string?)n.Attribute(W + "numId") == numId);
            var abstractId = (string?)num.Element(W + "abstractNumId")!.Attribute(W + "val");
            var level = num.Elements(W + "lvlOverride").Elements(W + "lvl")
                .Concat(numbering.Elements(W + "abstractNum")
                    .Single(a => (string?)a.Attribute(W + "abstractNumId") == abstractId).Elements(W + "lvl"))
                .First(l => (string?)l.Attribute(W + "ilvl") == ilvl);
            if (level.Element(W + "lvlPicBulletId") is not { } reference)
                return null;
            var bullet = numbering.Elements(W + "numPicBullet")
                .Single(b => (string?)b.Attribute(W + "numPicBulletId") == (string?)reference.Attribute(W + "val"));
            var image = part.GetPartById((string)bullet.Descendants(V + "imagedata").Single().Attribute(R + "id")!);
            using var bytes = new MemoryStream();
            using (var content = image.GetStream())
                content.CopyTo(bytes);
            return bytes.ToArray();
        }

        var items = Load(main).Descendants(W + "p")
            .Where(p => p.Element(W + "pPr")?.Element(W + "numPr") != null)
            .Select(p => (
                Text: string.Concat(p.Descendants().Where(e => e.Name == W + "t" || e.Name == W + "delText")
                    .Select(e => e.Value)),
                NumPr: p.Element(W + "pPr")!.Element(W + "numPr")!))
            .Where(p => p.Text.Length > 0)
            .Select(p => (p.Text, PictureOf(
                (string)p.NumPr.Element(W + "numId")!.Attribute(W + "val")!,
                (string)p.NumPr.Element(W + "ilvl")!.Attribute(W + "val")!)))
            .ToArray();
        return (numbering, items);

        static XElement Load(OpenXmlPart part)
        {
            using var content = part.GetStream();
            return XDocument.Load(content).Root!;
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InsertedPictureBulletList_ArrivesWithItsBulletAndImage(bool inLevelOverride)
    {
        var left = Doc("Morning checklist.", Array.Empty<string>());
        var right = Doc("Morning checklist.", new[] { "Inspect the valves", "Log the temperature" }, Teal,
            inLevelOverride: inLevelOverride);

        var redline = DocxDiff.Compare(left, right);

        foreach (var doc in new[] { redline, RevisionProcessor.AcceptRevisions(redline) })
        {
            var (numbering, items) = Bullets(doc);
            Assert.Equal(new[] { "Inspect the valves", "Log the temperature" }, items.Select(i => i.Text));
            Assert.All(items, item => Assert.Equal(Teal, item.Picture));
            // Both levels of the list use the bullet; it is copied once.
            Assert.Single(numbering.Elements(W + "numPicBullet"));
            Assert.Equal(2, numbering.Descendants(W + "lvlPicBulletId").Count());
        }
    }

    [Fact]
    public void BulletIdAlreadyNamingAnotherPicture_EachListKeepsItsOwnPicture()
    {
        // Unrelated lists, both under numId 1 and both calling their picture "bullet 0".
        var left = Doc("Checklist.", new[] { "alpha bravo charlie", "delta echo foxtrot" }, Red);
        var right = Doc("Checklist.", new[] { "neun zehn elf", "zwoelf dreizehn vierzehn" }, Teal, indent: 900);

        var redline = DocxDiff.Compare(left, right);

        var bullets = Bullets(redline).Numbering.Elements(W + "numPicBullet").ToList();
        Assert.Equal(new[] { "0", "1" }, bullets.Select(b => (string?)b.Attribute(W + "numPicBulletId")));
        // Both shapes were _x0000_i1025, and a part may not hold one shape id twice.
        Assert.Distinct(bullets.Descendants(V + "shape").Select(shape => (string?)shape.Attribute("id")));
        var accepted = Bullets(RevisionProcessor.AcceptRevisions(redline)).Items;
        Assert.Equal(new[] { "neun zehn elf", "zwoelf dreizehn vierzehn" }, accepted.Select(i => i.Text));
        Assert.All(accepted, item => Assert.Equal(Teal, item.Picture));
        var rejected = Bullets(RevisionProcessor.RejectRevisions(redline)).Items;
        Assert.Equal(new[] { "alpha bravo charlie", "delta echo foxtrot" }, rejected.Select(i => i.Text));
        Assert.All(rejected, item => Assert.Equal(Red, item.Picture));
    }

    [Fact]
    public void ListPresentOnBothSides_BulletIsNotCopiedAgain()
    {
        var left = Doc("Checklist.", new[] { "Inspect the valves", "Log the temperature" }, Teal);
        var right = Doc("Checklist.", new[] { "Inspect the valves", "Log the coolant temperature" }, Teal);

        var redline = DocxDiff.Compare(left, right);

        var (numbering, items) = Bullets(redline);
        Assert.All(items, item => Assert.Equal(Teal, item.Picture));
        Assert.Single(numbering.Elements(W + "numPicBullet"));
        using var stream = new MemoryStream(redline.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        Assert.Single(wdoc.MainDocumentPart!.NumberingDefinitionsPart!.Parts);
    }

    [Fact]
    public void LevelNamingABulletTheSourceLacks_IsImportedWithoutTheReference()
    {
        // Word refuses a document whose level names a missing picture bullet, so the redline must not carry one.
        var left = Doc("Checklist.", Array.Empty<string>());
        var right = Doc("Checklist.", new[] { "Inspect the valves" }, danglingBullet: true);

        var (numbering, items) = Bullets(DocxDiff.Compare(left, right));

        Assert.Null(Assert.Single(items).Picture);
        Assert.Empty(numbering.Descendants(W + "lvlPicBulletId"));
    }

    [Fact]
    public void Consolidate_ReviewerPictureBulletList_ArrivesWithItsBulletAndImage()
    {
        var original = Doc("Morning checklist.", Array.Empty<string>());
        var reviewed = Doc("Morning checklist.", new[] { "Inspect the valves" }, Teal);

        var consolidated = DocxDiff.Consolidate(
            original, new[] { new DocxDiffReviewer { Author = "Reviewer", Document = reviewed } });

        Assert.Equal(Teal, Assert.Single(Bullets(consolidated).Items).Picture);
    }
}
