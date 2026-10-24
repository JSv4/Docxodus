using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class LiteralTabTextTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var story in new[] { "body", "header", "footer" })
        foreach (var text in new[] { "\tHarbor", "North\tHarbor", " North\t\t  Harbor\t" })
        foreach (var literal in new[] { false, true })
            yield return new object[] { story, text, literal };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void Text_tabs_match_explicit_tab_elements(string story, string text, bool literal)
    {
        var expected = Convert(Fixture(story, text, literal: false));
        var actual = Convert(Fixture(story, text, literal));

        Assert.NotEmpty(Tabs(expected));
        Assert.Equal(Tabs(expected), Tabs(actual));
        Assert.Equal(expected.Element(Xhtml.body)!.Value, actual.Element(Xhtml.body)!.Value);
        Assert.DoesNotContain('\t', actual.Element(Xhtml.body)!.Value);
        Assert.Contains("Harbor", actual.Value);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Deleted_text_tabs_match_explicit_tab_elements(bool literal, bool hidden)
    {
        var expected = Convert(Fixture("body", "North\tHarbor", literal: false, deleted: true), hidden);
        var actual = Convert(Fixture("body", "North\tHarbor", literal, deleted: true), hidden);

        if (hidden) Assert.Empty(Tabs(expected));
        else Assert.NotEmpty(Tabs(expected));
        Assert.Equal(Tabs(expected), Tabs(actual));
        Assert.Equal(expected.Element(Xhtml.body)!.Value, actual.Element(Xhtml.body)!.Value);
        Assert.DoesNotContain('\t', actual.Element(Xhtml.body)!.Value);
    }

    [Fact]
    public void A_tab_in_field_instructions_does_not_become_a_layout_tab()
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture("body", "Harbor", literal: true));
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, true);
        var main = doc.MainDocumentPart!;
        var field = new Wp.Paragraph(
            new Wp.Run(new Wp.FieldChar { FieldCharType = Wp.FieldCharValues.Begin }),
            new Wp.Run(new Wp.FieldCode("PAGE\t") { Space = SpaceProcessingModeValues.Preserve }),
            new Wp.Run(new Wp.FieldChar { FieldCharType = Wp.FieldCharValues.Separate }),
            new Wp.Run(new Wp.Text("7")),
            new Wp.Run(new Wp.FieldChar { FieldCharType = Wp.FieldCharValues.End }));
        var document = main.Document!;
        var body = document.Body!;
        body.InsertBefore(field, body.GetFirstChild<Wp.SectionProperties>());
        document.Save();
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc));

        var html = WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings
            { RenderPagination = PaginationMode.Paginated, FabricateCssClasses = false });

        Assert.Empty(Tabs(html));
        Assert.Contains(html.Descendants(), e => (string?)e.Attribute("data-field") == "PAGE");
        Assert.Equal("PAGE\t", main.GetXDocument().Descendants(W.instrText).Single().Value);
    }

    private static string[] Tabs(XElement html) => html.Descendants()
        .Where(e => e.Attribute("data-docx-tab") != null)
        .Select(e => e.ToString(SaveOptions.DisableFormatting)).ToArray();

    private static XElement Convert(byte[] bytes, bool hidden = false) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("literal-tabs.docx", bytes),
            new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = false,
                RenderHeadersAndFooters = true,
                RenderPagination = PaginationMode.Paginated,
                RenderTrackedChanges = true,
                ShowDeletedContent = !hidden,
            });

    private static byte[] Fixture(string story, string text, bool literal, bool deleted = false)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(new Wp.DocDefaults(
                new Wp.RunPropertiesDefault(new Wp.RunPropertiesBaseStyle(
                    new Wp.RunFonts { Ascii = "Cedar Unavailable Serif", HighAnsi = "Cedar Unavailable Serif" },
                    new Wp.FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings(new Wp.DefaultTabStop { Val = 720 });
            var paragraph = new Wp.Paragraph(new Wp.ParagraphProperties(new Wp.Tabs(
                new Wp.TabStop { Val = Wp.TabStopValues.Right, Position = 9360, Leader = Wp.TabStopLeaderCharValues.Dot })),
                new Wp.Run(new Wp.Text("")));
            var run = new Wp.Run(new Wp.RunProperties(new Wp.Bold(), new Wp.Color { Val = "225588" }));
            // The docx-js shape has an empty run, a tab-only text run, then its label.
            var separateLabel = text == "\tHarbor";
            var content = separateLabel ? "\t" : text;
            void AddText(string value)
            {
                if (value.Length == 0) return;
                if (deleted) run.Append(new Wp.DeletedText(value) { Space = SpaceProcessingModeValues.Preserve });
                else run.Append(new Wp.Text(value) { Space = SpaceProcessingModeValues.Preserve });
            }
            if (literal) AddText(content);
            else
            {
                var segments = content.Split('\t');
                for (var i = 0; i < segments.Length; i++)
                {
                    if (i > 0) run.Append(new Wp.TabChar());
                    AddText(segments[i]);
                }
            }
            if (deleted) paragraph.Append(new Wp.DeletedRun(run)
                { Id = "7", Author = "Ferry reviewer", Date = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc) });
            else paragraph.Append(run);
            if (separateLabel) paragraph.Append(new Wp.Run(
                new Wp.RunProperties(new Wp.Bold(), new Wp.Color { Val = "225588" }), new Wp.Text("Harbor")));

            var section = new Wp.SectionProperties();
            if (story == "header")
            {
                var part = main.AddNewPart<HeaderPart>();
                part.Header = new Wp.Header(paragraph);
                section.Append(new Wp.HeaderReference { Type = Wp.HeaderFooterValues.Default, Id = main.GetIdOfPart(part) });
            }
            else if (story == "footer")
            {
                var part = main.AddNewPart<FooterPart>();
                part.Footer = new Wp.Footer(paragraph);
                section.Append(new Wp.FooterReference { Type = Wp.HeaderFooterValues.Default, Id = main.GetIdOfPart(part) });
            }
            section.Append(new Wp.PageSize { Width = 12240, Height = 15840 },
                new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720 });
            main.Document = new Wp.Document(new Wp.Body(story == "body" ? paragraph :
                new Wp.Paragraph(new Wp.Run(new Wp.Text("Body control."))), section));
            doc.Save();
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc));
        }
        return stream.ToArray();
    }
}
