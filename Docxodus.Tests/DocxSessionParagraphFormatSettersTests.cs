#nullable enable

// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The Paragraph dialog's setters on <see cref="DocxSession.SetParagraphFormat"/>: absolute
/// left/right indent and the on/off properties keepNext, keepLines, widowControl and
/// contextualSpacing (written only when they differ from the style chain). IDs PF###.
/// </summary>
public class DocxSessionParagraphFormatSettersTests
{
    /// <summary>Two paragraphs: the first in Normal, the second in "Keepy", a style that turns on
    /// keepNext and keepLines and turns OFF widow control.</summary>
    private static byte[] BuildDoc()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style(new StyleName { Val = "Normal" })
                { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
                new Style(
                    new StyleName { Val = "Keepy" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(
                        new KeepNext(), new KeepLines(), new WidowControl { Val = false }))
                { Type = StyleValues.Paragraph, StyleId = "Keepy" });
            var settingsPart = main.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings();
            body.Append(new Paragraph(new Run(new Text("Plain paragraph."))));
            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Keepy" }),
                new Run(new Text("Styled paragraph."))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>The anchor's XML without the engine's Unid marks, self-closing tags spelled one way.</summary>
    private static string Xml(DocxSession s, string anchor) =>
        System.Text.RegularExpressions.Regex.Replace(s.Raw.GetXml(anchor), " p1:Unid=\"[0-9a-f]+\"", "")
            .Replace(" />", "/>");

    private static (DocxSession s, string plain, string styled) Open()
    {
        var s = new DocxSession(BuildDoc());
        var anchors = s.Project().AnchorIndex.Keys.ToList();
        return (s, anchors[0], anchors[1]);
    }

    [Fact]
    public void PF001_LeftRightIndent_Absolute_AndNegative()
    {
        var (s, plain, _) = Open();
        using (s)
        {
            var r = s.SetParagraphFormat(plain, new ParagraphFormatOp { LeftIndent = 720, RightIndent = 360 });
            Assert.True(r.Success, r.Error?.Message);
            var xml = Xml(s, plain);
            Assert.Contains("w:left=\"720\"", xml);
            Assert.Contains("w:right=\"360\"", xml);

            // Absolute, not a delta; and Word allows text in the margin.
            Assert.True(s.SetParagraphFormat(plain, new ParagraphFormatOp { LeftIndent = -360 }).Success);
            xml = Xml(s, plain);
            Assert.Contains("w:left=\"-360\"", xml);
            Assert.Contains("w:right=\"360\"", xml);

            var f = s.GetFormatting(plain)!;
            Assert.Equal(-360, f.EffectiveParagraph.LeftIndentTwips);
            Assert.Equal(360, f.EffectiveParagraph.RightIndentTwips);
        }
    }

    [Fact]
    public void PF002_LeftIndent_KeepsStartSpelling_AndExcludesIndentDelta()
    {
        var (s, plain, _) = Open();
        using (s)
        {
            // A paragraph written with the bidi-neutral spelling keeps it. Its indents in character
            // units, which Word would show instead of the twips, go.
            Assert.True(s.Raw.ReplaceXml(plain,
                "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                "<w:pPr><w:ind w:start=\"100\" w:startChars=\"200\" w:end=\"50\" w:endChars=\"100\"/></w:pPr>" +
                "<w:r><w:t>Plain paragraph.</w:t></w:r></w:p>").Success);
            plain = s.Project().AnchorIndex.Keys.First();
            Assert.True(s.SetParagraphFormat(plain, new ParagraphFormatOp { LeftIndent = 1440, RightIndent = 0 }).Success);
            var xml = Xml(s, plain);
            Assert.Contains("w:start=\"1440\"", xml);
            Assert.Contains("w:end=\"0\"", xml);
            Assert.DoesNotContain("w:left=", xml);
            Assert.DoesNotContain("Chars=", xml);

            var both = s.SetParagraphFormat(plain, new ParagraphFormatOp { LeftIndent = 0, IndentDelta = 360 });
            Assert.False(both.Success);
            Assert.Equal(EditErrorCode.InvalidParagraphFormat, both.Error!.Code);
        }
    }

    [Fact]
    public void PF003_Toggles_OnPlainParagraph_WriteOnlyDifferences()
    {
        var (s, plain, _) = Open();
        using (s)
        {
            var r = s.SetParagraphFormat(plain, new ParagraphFormatOp
            {
                KeepNext = true, KeepLines = true, ContextualSpacing = true, WidowControl = true,
            });
            Assert.True(r.Success, r.Error?.Message);
            var xml = Xml(s, plain);
            Assert.Contains("<w:keepNext/>", xml);
            Assert.Contains("<w:keepLines/>", xml);
            Assert.Contains("<w:contextualSpacing/>", xml);
            // Widow control is already on when nothing mentions it: no element.
            Assert.DoesNotContain("widowControl", xml);

            var f = s.GetFormatting(plain)!.EffectiveParagraph;
            Assert.True(f.KeepNext);
            Assert.True(f.KeepLines);
            Assert.True(f.ContextualSpacing);
            Assert.True(f.WidowControl);

            // Off again: the direct elements go away (the style says nothing).
            Assert.True(s.SetParagraphFormat(plain, new ParagraphFormatOp
            {
                KeepNext = false, KeepLines = false, ContextualSpacing = false, WidowControl = false,
            }).Success);
            xml = Xml(s, plain);
            Assert.DoesNotContain("keepNext", xml);
            Assert.DoesNotContain("keepLines", xml);
            Assert.DoesNotContain("contextualSpacing", xml);
            Assert.Contains("<w:widowControl w:val=\"0\"/>", xml);
            Assert.False(s.GetFormatting(plain)!.EffectiveParagraph.WidowControl);
        }
    }

    [Fact]
    public void PF004_Toggles_AgainstStyle_WriteExplicitOverrides()
    {
        var (s, _, styled) = Open();
        using (s)
        {
            var before = s.GetFormatting(styled)!.EffectiveParagraph;
            Assert.True(before.KeepNext);
            Assert.False(before.WidowControl);

            Assert.True(s.SetParagraphFormat(styled, new ParagraphFormatOp { KeepNext = false, WidowControl = true }).Success);
            var xml = Xml(s, styled);
            Assert.Contains("<w:keepNext w:val=\"0\"/>", xml);
            Assert.Contains("<w:widowControl/>", xml);
            var after = s.GetFormatting(styled)!.EffectiveParagraph;
            Assert.False(after.KeepNext);
            Assert.True(after.KeepLines); // untouched, still from the style
            Assert.True(after.WidowControl);

            // Back to the style's values: the overrides are removed, not flipped.
            Assert.True(s.SetParagraphFormat(styled, new ParagraphFormatOp { KeepNext = true, WidowControl = false }).Success);
            xml = Xml(s, styled);
            Assert.DoesNotContain("keepNext", xml);
            Assert.DoesNotContain("widowControl", xml);
        }
    }

    [Fact]
    public void PF005_WholeDialog_IsOneUndoStep()
    {
        var (s, plain, _) = Open();
        using (s)
        {
            var original = Xml(s, plain);
            var r = s.SetParagraphFormat(plain, new ParagraphFormatOp
            {
                Alignment = ParagraphAlignment.Center,
                LeftIndent = 720, RightIndent = 720, HangingIndent = 360,
                SpacingBefore = 120, SpacingAfter = 240, LineSpacing = 360,
                KeepNext = true, PageBreakBefore = true,
            });
            Assert.True(r.Success, r.Error?.Message);
            Assert.NotEqual(original, Xml(s, plain));
            Assert.True(s.Undo());
            Assert.Equal(original, Xml(s, plain));
        }
    }

    [Fact]
    public void PF006_Json_RoundTrip()
    {
        var op = Docxodus.Internal.DocxSessionJson.ParseParagraphFormatOp(
            "{\"leftIndent\":-200,\"rightIndent\":100,\"keepNext\":true,\"keepLines\":false," +
            "\"widowControl\":false,\"contextualSpacing\":true}");
        Assert.Equal(-200, op.LeftIndent);
        Assert.Equal(100, op.RightIndent);
        Assert.True(op.KeepNext);
        Assert.False(op.KeepLines);
        Assert.False(op.WidowControl);
        Assert.True(op.ContextualSpacing);

        var (s, plain, _) = Open();
        using (s)
        {
            Assert.True(s.SetParagraphFormat(plain, op).Success);
            var json = Docxodus.Internal.DocxSessionJson.SerializeFormattingInspectionOrNull(s.GetFormatting(plain)!);
            Assert.Contains("\"widowControl\":false", json);
            Assert.Contains("\"contextualSpacing\":true", json);
        }
    }
}
