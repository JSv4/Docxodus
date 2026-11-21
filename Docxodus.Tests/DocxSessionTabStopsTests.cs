using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tab stops on <see cref="DocxSession"/>: GetTabStops, SetTabStops (Word's way: own stops plus
/// clears of unwanted style stops) and SetDefaultTabStop. Test IDs use the TS0xx range.
/// </summary>
public class DocxSessionTabStopsTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Style "Tabbed" sets a left stop at 1 inch; the first paragraph uses it and adds a
    /// dotted right stop at 4 inches, the second has no style and no tabs.</summary>
    private static byte[] Document(string settingsXml = "")
    {
        const string ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            void Write(OpenXmlPart part, string xml)
            {
                using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write(xml);
            }
            Write(main, $"<w:document {ns}><w:body>"
                + "<w:p><w:pPr><w:pStyle w:val=\"Tabbed\"/><w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"5760\"/></w:tabs></w:pPr>"
                + "<w:r><w:t>One</w:t></w:r></w:p>"
                + "<w:p><w:r><w:t>Two</w:t></w:r></w:p>"
                + "</w:body></w:document>");
            Write(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings {ns}>{settingsXml}</w:settings>");
            Write(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {ns}>"
                + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>"
                + "<w:style w:type=\"paragraph\" w:styleId=\"Tabbed\"><w:name w:val=\"Tabbed\"/><w:basedOn w:val=\"Normal\"/>"
                + "<w:pPr><w:tabs><w:tab w:val=\"left\" w:pos=\"1440\"/></w:tabs></w:pPr></w:style>"
                + "</w:styles>");
        }
        return stream.ToArray();
    }

    private static string[] Paragraphs(DocxSession session) =>
        session.Project().AnchorIndex.Values.Where(t => t.Anchor.Kind == "p").Select(t => t.Anchor.Id).ToArray();

    private static XElement Body(DocxSession session)
    {
        using var ms = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!;
    }

    private static string OwnTabs(DocxSession session, int paragraph) =>
        string.Join(" ", Body(session).Elements(W + "p").ElementAt(paragraph)
            .Element(W + "pPr")?.Element(W + "tabs")?.Elements(W + "tab")
            .Select(t => $"{(string?)t.Attribute(W + "val")}{((string?)t.Attribute(W + "leader") is { } l ? "/" + l : "")}@{(string?)t.Attribute(W + "pos")}")
            ?? Enumerable.Empty<string>());

    private static ParagraphTabStop Stop(int pos, TabStopAlignment alignment = TabStopAlignment.Left,
        TabStopLeader leader = TabStopLeader.None) => new() { PositionTwips = pos, Alignment = alignment, Leader = leader };

    [Fact]
    public void TS001_GetTabStops_ReadsStopsInEffect_StyleStops_AndTheDefaultInterval()
    {
        var session = new DocxSession(Document());
        var p = Paragraphs(session);
        var info = session.GetTabStops(p[0])!;
        Assert.Equal(new[] { Stop(1440), Stop(5760, TabStopAlignment.Right, TabStopLeader.Dot) }, info.Stops);
        Assert.Equal(new[] { Stop(1440) }, info.StyleStops);
        Assert.Equal(720, info.DefaultTabStopTwips);

        var plain = session.GetTabStops(p[1])!;
        Assert.Empty(plain.Stops);
        Assert.Empty(plain.StyleStops);
    }

    [Fact]
    public void TS002_SetTabStops_WritesOwnStops_AndClearsUnwantedStyleStops_OneUndo()
    {
        var session = new DocxSession(Document());
        var p = Paragraphs(session)[0];
        var before = Body(session);
        var wanted = new[] { Stop(2880, TabStopAlignment.Center), Stop(4320, TabStopAlignment.Decimal, TabStopLeader.Hyphen) };
        var r = session.SetTabStops(p, wanted);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal("clear@1440 center@2880 decimal/hyphen@4320", OwnTabs(session, 0));
        Assert.Equal(wanted, session.GetTabStops(p)!.Stops);
        PackageValidation.AssertValid(session.Save(), DocumentFormat.OpenXml.FileFormatVersions.Office2007);

        Assert.True(session.Undo());
        Assert.True(XNode.DeepEquals(before, Body(session)));
    }

    [Fact]
    public void TS003_AStyleStopKeptIsNotWrittenAgain_AndNoStopsWritesNoTabs()
    {
        var session = new DocxSession(Document());
        var p = Paragraphs(session);
        Assert.True(session.SetTabStops(p[0], new[] { Stop(1440), Stop(3600, TabStopAlignment.Bar) }).Success);
        Assert.Equal("bar@3600", OwnTabs(session, 0));

        // Keeping exactly the style's stops leaves no w:tabs of its own.
        Assert.True(session.SetTabStops(p[0], new[] { Stop(1440) }).Success);
        Assert.Null(Body(session).Elements(W + "p").First().Element(W + "pPr")!.Element(W + "tabs"));

        Assert.True(session.SetTabStops(p[1], new[] { Stop(720, TabStopAlignment.Right) }).Success);
        Assert.Equal("right@720", OwnTabs(session, 1));
        Assert.True(session.SetTabStops(p[1], System.Array.Empty<ParagraphTabStop>()).Success);
        Assert.Equal("", OwnTabs(session, 1));
    }

    [Fact]
    public void TS004_InvalidStopsAndAnchors_AreRefused_AndChangeNothing()
    {
        var session = new DocxSession(Document());
        var p = Paragraphs(session)[1];
        long version = session.Version;
        var duplicate = session.SetTabStops(p, new[] { Stop(720), Stop(720, TabStopAlignment.Right) });
        Assert.Equal(EditErrorCode.InvalidParagraphFormat, duplicate.Error!.Code);
        var far = session.SetTabStops(p, new[] { Stop(40000) });
        Assert.Equal(EditErrorCode.InvalidParagraphFormat, far.Error!.Code);
        Assert.False(session.SetTabStops("no-such-anchor", new[] { Stop(720) }).Success);
        Assert.Equal(version, session.Version);
        Assert.Null(session.GetTabStops("no-such-anchor"));
    }

    [Fact]
    public void TS005_Tracked_RecordsOnePPrChange()
    {
        var session = new DocxSession(Document(), new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Bob",
        });
        Assert.True(session.SetTabStops(Paragraphs(session)[1], new[] { Stop(1440, TabStopAlignment.Center) }).Success);
        var pPr = Body(session).Elements(W + "p").ElementAt(1).Element(W + "pPr")!;
        var change = pPr.Element(W + "pPrChange")!;
        Assert.Equal("Bob", (string?)change.Attribute(W + "author"));
        Assert.Null(change.Element(W + "pPr")!.Element(W + "tabs"));
        PackageValidation.AssertValid(session.Save(), DocumentFormat.OpenXml.FileFormatVersions.Office2007);
    }

    [Fact]
    public void TS006_SetDefaultTabStop_WritesSettings_ReadsBack_AndUndoes()
    {
        var session = new DocxSession(Document("<w:zoom w:percent=\"100\"/><w:characterSpacingControl w:val=\"doNotCompress\"/>"));
        var p = Paragraphs(session)[1];
        var r = session.SetDefaultTabStop(567);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(567, session.GetTabStops(p)!.DefaultTabStopTwips);
        using (var ms = new MemoryStream(session.Save()))
        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            var settings = doc.MainDocumentPart!.DocumentSettingsPart!.GetXDocument().Root!;
            // In its CT_Settings slot: after w:zoom, before w:characterSpacingControl.
            Assert.Equal(new[] { "zoom", "defaultTabStop", "characterSpacingControl" },
                settings.Elements().Select(e => e.Name.LocalName));
        }
        PackageValidation.AssertValid(session.Save(), DocumentFormat.OpenXml.FileFormatVersions.Office2007);

        Assert.True(session.Undo());
        Assert.Equal(720, session.GetTabStops(p)!.DefaultTabStopTwips);
        Assert.False(session.SetDefaultTabStop(0).Success);
    }

    [Fact]
    public void TS007_Json_ParsesStops_AndSerializesTheInfo()
    {
        var stops = DocxSessionJson.ParseTabStops(
            "[{\"positionTwips\":1440,\"alignment\":\"decimal\",\"leader\":\"middleDot\"},{\"positionTwips\":2880}]");
        Assert.Equal(new[] { Stop(1440, TabStopAlignment.Decimal, TabStopLeader.MiddleDot), Stop(2880) }, stops);

        var session = new DocxSession(Document());
        using var json = JsonDocument.Parse(DocxSessionJson.SerializeTabStopsInfo(session.GetTabStops(Paragraphs(session)[0])));
        var root = json.RootElement;
        Assert.Equal(2, root.GetProperty("stops").GetArrayLength());
        Assert.Equal("right", root.GetProperty("stops")[1].GetProperty("alignment").GetString());
        Assert.Equal("dot", root.GetProperty("stops")[1].GetProperty("leader").GetString());
        Assert.Equal(1440, root.GetProperty("styleStops")[0].GetProperty("positionTwips").GetInt32());
        Assert.Equal(720, root.GetProperty("defaultTabStopTwips").GetInt32());
        Assert.Equal("null", DocxSessionJson.SerializeTabStopsInfo(null));
    }
}
