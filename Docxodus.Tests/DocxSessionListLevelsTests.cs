#nullable enable

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A list's levels (Word's Define New Multilevel List) and Continue Numbering:
/// <see cref="DocxSession.GetListLevels"/>, <see cref="DocxSession.SetListLevels"/>,
/// <see cref="DocxSession.ContinueNumbering"/>. IDs LL###.
/// </summary>
public class DocxSessionListLevelsTests
{
    private static byte[] BuildDoc(params string[] paragraphs)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document(new Body(paragraphs.Select(t => new Paragraph(new Run(new Text(t))))));
            // (the list-number resolver reads the styles part)
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static (DocxSession s, string[] anchors) Open(params string[] paragraphs)
    {
        var s = new DocxSession(BuildDoc(paragraphs));
        return (s, s.Project().AnchorIndex.Keys.ToArray());
    }

    /// <summary>The current anchor of the paragraph whose text starts with <paramref name="text"/>.</summary>
    private static string AnchorOf(DocxSession s, string text) =>
        s.Project().AnchorIndex.First(kv => s.Raw.GetXml(kv.Key).Contains(">" + text)).Key;

    private static string? Label(DocxSession s, string text) => s.GetListMembership(AnchorOf(s, text))?.GeneratedLabel;

    /// <summary>Two items, the second a level in: "One" (1.) and "One-one" (level 2).</summary>
    private static DocxSession NestedList()
    {
        var (s, a) = Open("One", "One-one", "Two");
        Assert.True(s.ApplyListFormatRange(a[0], a[2], ListFormat.Decimal).Success);
        Assert.True(s.SetListLevel(AnchorOf(s, "One-one"), 1).Success);
        return s;
    }

    [Fact]
    public void LL001_GetListLevels_ReadsNineLevels_AndNullOffList()
    {
        var (s, a) = Open("Item", "Plain");
        using (s)
        {
            Assert.True(s.ApplyListFormat(a[0], ListFormat.Decimal).Success);
            var levels = s.GetListLevels(AnchorOf(s, "Item"))!;
            Assert.Equal(9, levels.Count);
            Assert.Equal("decimal", levels[0].Format);
            Assert.Equal("%1.", levels[0].Text);
            Assert.Equal(1, levels[0].Start);
            Assert.Equal(720, levels[0].TextPositionTwips);
            Assert.Equal(360, levels[0].NumberPositionTwips);
            Assert.Null(s.GetListLevels(AnchorOf(s, "Plain")));
        }
    }

    [Fact]
    public void LL002_LegalNumbering_ChangesOnlyThisList()
    {
        using var s = NestedList();
        // Another list of another format elsewhere: untouched by the change.
        var plain = s.InsertParagraph(AnchorOf(s, "Two"), Position.After, "Other");
        Assert.True(plain.Success);
        Assert.True(s.ApplyListFormat(AnchorOf(s, "Other"), ListFormat.LowerLetter).Success);

        var r = s.SetListLevels(AnchorOf(s, "One"), new[]
        {
            new ListLevelSpec { Level = 0, Text = "%1." },
            new ListLevelSpec { Level = 1, Format = "decimal", Text = "%1.%2", Legal = true },
        });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal("1.1", Label(s, "One-one"));
        Assert.Equal("2.", Label(s, "Two"));
        Assert.Equal("a.", Label(s, "Other"));

        // A NEW decimal list still gets Docxodus's shared definition, not this list's copy.
        var fresh = s.InsertParagraph(AnchorOf(s, "Other"), Position.After, "Fresh");
        Assert.True(fresh.Success);
        Assert.True(s.ApplyListFormat(AnchorOf(s, "Fresh"), ListFormat.Decimal).Success);
        var freshLevels = s.GetListLevels(AnchorOf(s, "Fresh"))!;
        Assert.Equal("%2.", freshLevels[1].Text);
    }

    [Fact]
    public void LL003_Positions_WriteLeftAndHanging()
    {
        using var s = NestedList();
        Assert.True(s.SetListLevels(AnchorOf(s, "One"), new[]
        {
            new ListLevelSpec { Level = 0, NumberPositionTwips = 360, TextPositionTwips = 1080 },
            new ListLevelSpec { Level = 1, NumberPositionTwips = 1440, TextPositionTwips = 1440 },
        }).Success);
        var levels = s.GetListLevels(AnchorOf(s, "One"))!;
        Assert.Equal((360, 1080), (levels[0].NumberPositionTwips, levels[0].TextPositionTwips));
        Assert.Equal((1440, 1440), (levels[1].NumberPositionTwips, levels[1].TextPositionTwips));
        // Only one of the two given: the other stays.
        Assert.True(s.SetListLevels(AnchorOf(s, "One"), new[] { new ListLevelSpec { Level = 0, TextPositionTwips = 1440 } }).Success);
        levels = s.GetListLevels(AnchorOf(s, "One"))!;
        Assert.Equal((360, 1440), (levels[0].NumberPositionTwips, levels[0].TextPositionTwips));
    }

    [Fact]
    public void LL004_BulletToNumber_DropsSymbolFont_AndTakesANumberText()
    {
        var (s, a) = Open("Point");
        using (s)
        {
            Assert.True(s.ApplyListFormat(a[0], ListFormat.Bullet).Success);
            Assert.True(s.SetListLevels(AnchorOf(s, "Point"), new[] { new ListLevelSpec { Level = 0, Format = "upperRoman" } }).Success);
            var level = s.GetListLevels(AnchorOf(s, "Point"))![0];
            Assert.Equal(("upperRoman", "%1."), (level.Format, level.Text));
            Assert.Equal("I.", Label(s, "Point"));
            using var saved = new MemoryStream(s.Save());
            using var doc = WordprocessingDocument.Open(saved, false);
            var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering.OuterXml;
            // The list's own copy has no Symbol font on its first level any more.
            var own = doc.MainDocumentPart.NumberingDefinitionsPart.Numbering.Elements<AbstractNum>()
                .Single(an => an.Descendants<NumberingFormat>().Any(f => f.Val?.Value == NumberFormatValues.UpperRoman));
            Assert.DoesNotContain("Symbol", own.Elements<Level>().First().OuterXml);
            Assert.Contains("Symbol", numbering);    // (the shared bullet definition keeps it)
        }
    }

    [Fact]
    public void LL005_BadLevels_AreRefused_WithoutAnUndoStep()
    {
        using var s = NestedList();
        var anchor = AnchorOf(s, "One");
        foreach (var bad in new[]
        {
            new ListLevelSpec { Level = 9 },
            new ListLevelSpec { Level = 0, Format = "sparkles" },
            new ListLevelSpec { Level = 1, Text = "%3." },
            new ListLevelSpec { Level = 0, Start = -1 },
        })
        {
            var r = s.SetListLevels(anchor, new[] { bad });
            Assert.False(r.Success);
            Assert.Equal(EditErrorCode.InvalidListLevel, r.Error!.Code);
        }
        // The last real edit (the nesting) is still what one Undo takes back.
        Assert.True(s.Undo());
        Assert.Equal("2.", Label(s, "One-one"));
    }

    [Fact]
    public void LL006_Undo_RestoresTheLevels_AndStartDropsARestart()
    {
        using var s = NestedList();
        Assert.True(s.SetListLevels(AnchorOf(s, "One"), new[] { new ListLevelSpec { Level = 0, Start = 5 } }).Success);
        Assert.Equal("5.", Label(s, "One"));
        Assert.True(s.Undo());
        Assert.Equal("1.", Label(s, "One"));

        // A restart (start override) at the level would hide a new start: setting Start drops it.
        Assert.True(s.SetListStartOverride(AnchorOf(s, "Two"), 7).Success);
        Assert.Equal("7.", Label(s, "Two"));
        Assert.True(s.SetListLevels(AnchorOf(s, "Two"), new[] { new ListLevelSpec { Level = 0, Start = 3 } }).Success);
        Assert.Equal("3.", Label(s, "Two"));
    }

    [Fact]
    public void LL007_ContinueNumbering_JoinsThePreviousList()
    {
        var (s, a) = Open("A1", "A2", "Between", "B1", "B2");
        using (s)
        {
            Assert.True(s.ApplyListFormatRange(a[0], a[1], ListFormat.Decimal).Success);
            Assert.True(s.ApplyListFormatRange(AnchorOf(s, "B1"), AnchorOf(s, "B2"), ListFormat.Decimal).Success);
            // Make B a list of its own, numbered from 1.
            Assert.True(s.SetListStartOverride(AnchorOf(s, "B1"), 1).Success);
            Assert.Equal("1.", Label(s, "B1"));

            var r = s.ContinueNumbering(AnchorOf(s, "B1"));
            Assert.True(r.Success, r.Error?.Message);
            Assert.Equal("3.", Label(s, "B1"));
            Assert.Equal("4.", Label(s, "B2"));

            var first = s.ContinueNumbering(AnchorOf(s, "A1"));
            Assert.False(first.Success);
            Assert.Equal(EditErrorCode.AnchorWrongKind, first.Error!.Code);
        }
    }

    [Fact]
    public void LL008_Json_RoundTrip()
    {
        var specs = Docxodus.Internal.DocxSessionJson.ParseListLevelSpecs(
            "[{\"level\":1,\"format\":\"decimal\",\"text\":\"%1.%2\",\"legal\":true,\"start\":2," +
            "\"numberPositionTwips\":720,\"textPositionTwips\":1440}]");
        var spec = Assert.Single(specs);
        Assert.Equal((1, "decimal", "%1.%2", true, 2, 720, 1440),
            (spec.Level, spec.Format, spec.Text, spec.Legal, spec.Start, spec.NumberPositionTwips, spec.TextPositionTwips));
        Assert.Empty(Docxodus.Internal.DocxSessionJson.ParseListLevelSpecs("{\"levels\":[]}"));

        using var s = NestedList();
        var json = Docxodus.Internal.DocxSessionJson.SerializeListLevelsOrNull(s.GetListLevels(AnchorOf(s, "One")));
        Assert.StartsWith("[{\"level\":0,\"format\":\"decimal\",\"text\":\"%1.\"", json);
        Assert.Equal("null", Docxodus.Internal.DocxSessionJson.SerializeListLevelsOrNull(null));
    }
}
