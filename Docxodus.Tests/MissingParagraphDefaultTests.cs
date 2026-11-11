using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1652: Word 2024 lays out a document whose styles part has no <c>w:pPrDefault</c> with 8 pt
/// after and 1.16 lines (<c>w:after="160" w:line="278"</c>), while a <c>w:pPrDefault</c> that is
/// present, even empty, means what it says. Every expectation here was measured in Word. See
/// docs/ooxml_corner_cases.md, "A styles part without <c>w:pPrDefault</c>".
/// </summary>
public class MissingParagraphDefaultTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string RunDefault =
        "<w:rPrDefault><w:rPr><w:rFonts w:ascii='Liberation Serif' w:hAnsi='Liberation Serif'/><w:sz w:val='24'/></w:rPr></w:rPrDefault>";
    private const string Normal = "<w:style w:type='paragraph' w:default='1' w:styleId='Normal'><w:name w:val='Normal'/>";
    private const string Explicit = "<w:spacing w:after='0' w:line='240' w:lineRule='auto'/>";

    /// <summary>A three-paragraph document; <paramref name="styles"/> is the content of
    /// <c>w:styles</c>, or null for a package with no styles part.</summary>
    private static WmlDocument Doc(string? styles, string extraBody = "")
    {
        using var memory = new MemoryStream();
        using (var package = WordprocessingDocument.Create(memory, WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using (var stream = main.GetStream())
                XDocument.Parse($"<w:document xmlns:w='{W}'><w:body>" +
                    "<w:p><w:r><w:t>Quartz01</w:t></w:r></w:p><w:p><w:r><w:t>Quartz02</w:t></w:r></w:p>" +
                    $"<w:p><w:r><w:t>Quartz03</w:t></w:r></w:p>{extraBody}</w:body></w:document>").Save(stream);
            if (styles is not null)
                using (var stream = main.AddNewPart<StyleDefinitionsPart>().GetStream())
                    XDocument.Parse($"<w:styles xmlns:w='{W}'>{styles}</w:styles>").Save(stream);
        }
        return new WmlDocument("missing-paragraph-default.docx", memory.ToArray());
    }

    /// <summary>Each paragraph's effective <c>after/line</c> as the renderer resolves it.</summary>
    private static string Spacing(WmlDocument document)
    {
        var assembled = FormattingAssembler.AssembleFormatting(document, new FormattingAssemblerSettings());
        using var memory = new MemoryStream(assembled.DocumentByteArray);
        using var package = WordprocessingDocument.Open(memory, false);
        using var stream = package.MainDocumentPart!.GetStream();
        return string.Join(" ", XDocument.Load(stream).Descendants(W + "p").Select(paragraph =>
        {
            var spacing = paragraph.Element(W + "pPr")?.Element(W + "spacing");
            return $"{(string?)spacing?.Attribute(W + "after") ?? "0"}/{(string?)spacing?.Attribute(W + "line") ?? "240"}";
        }).Distinct());
    }

    [Theory]
    // Missing: Word supplies 8 pt after and 1.16 lines.
    [InlineData("<w:docDefaults>" + RunDefault + "</w:docDefaults>" + Normal + "</w:style>", "160/278")]
    [InlineData("<w:docDefaults/>" + Normal + "</w:style>", "160/278")]
    [InlineData(Normal + "</w:style>", "160/278")]
    [InlineData("", "160/278")]
    [InlineData(null, "160/278")]
    // A style overrides the supplied default one attribute at a time, like any declared default.
    [InlineData("<w:docDefaults>" + RunDefault + "</w:docDefaults>" + Normal + "<w:pPr><w:spacing w:after='0'/></w:pPr></w:style>", "0/278")]
    [InlineData("<w:docDefaults>" + RunDefault + "</w:docDefaults>" + Normal + "<w:pPr><w:spacing w:line='240' w:lineRule='auto'/></w:pPr></w:style>", "160/240")]
    // Present: taken as written, so an empty default is single spacing.
    [InlineData("<w:docDefaults>" + RunDefault + "<w:pPrDefault/></w:docDefaults>" + Normal + "</w:style>", "0/240")]
    [InlineData("<w:docDefaults>" + RunDefault + "<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>" + Normal + "</w:style>", "0/240")]
    [InlineData("<w:docDefaults>" + RunDefault + "<w:pPrDefault><w:pPr><w:jc w:val='left'/></w:pPr></w:pPrDefault></w:docDefaults>" + Normal + "</w:style>", "0/240")]
    [InlineData("<w:docDefaults>" + RunDefault + "<w:pPrDefault><w:pPr>" + Explicit + "</w:pPr></w:pPrDefault></w:docDefaults>" + Normal + "</w:style>", "0/240")]
    public void EffectiveSpacingMatchesWord(string? styles, string expected)
    {
        var document = Doc(styles);
        Assert.Equal(expected, Spacing(document));

        // The read-only resolver behind session introspection and automatic spacing agrees with the renderer.
        using var memory = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(memory, false);
        var spacing = FormattingAssembler.ResolveEffectiveParagraphProperties(package,
            package.MainDocumentPart!.GetXDocument().Descendants(W + "p").First()).Element(W + "spacing");
        Assert.Equal(expected, $"{(string?)spacing?.Attribute(W + "after") ?? "0"}/{(string?)spacing?.Attribute(W + "line") ?? "240"}");
    }

    [Theory]
    [InlineData("", "8pt", "1.158")]
    [InlineData("<w:pPrDefault/>", ".001pt", null)]
    public void HtmlRendersTheSpacing(string paragraphDefault, string? marginBottom, string? lineSpacing)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(
            Doc($"<w:docDefaults>{RunDefault}{paragraphDefault}</w:docDefaults>{Normal}</w:style>"),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        var style = ((string?)html.Descendants().First(e => e.Name.LocalName == "p").Attribute("style") ?? "")
            .Split(';', StringSplitOptions.RemoveEmptyEntries).Select(d => d.Split(':', 2))
            .ToDictionary(d => d[0].Trim(), d => d[1].Trim());
        Assert.Equal(marginBottom, style.GetValueOrDefault("margin-bottom"));
        Assert.Equal(lineSpacing, style.GetValueOrDefault("--docx-auto-line-spacing"));
    }

    /// <summary>A document without the default and one that states Word's values lay out identically
    /// in Word, so comparing them finds no formatting change; against explicit single spacing it does.</summary>
    [Theory]
    [InlineData("<w:spacing w:after='160' w:line='278' w:lineRule='auto'/>", false)]
    [InlineData(Explicit, true)]
    public void CompareTreatsTheMissingDefaultAsWordsValues(string statedSpacing, bool changed)
    {
        var missing = Doc($"<w:docDefaults>{RunDefault}</w:docDefaults>{Normal}</w:style>");
        var stated = Doc($"<w:docDefaults>{RunDefault}<w:pPrDefault><w:pPr>{statedSpacing}</w:pPr></w:pPrDefault></w:docDefaults>{Normal}</w:style>");

        var redline = DocxCompare.Compare(missing, stated);

        using var memory = new MemoryStream(redline.DocumentByteArray);
        using var package = WordprocessingDocument.Open(memory, false);
        using var stream = package.MainDocumentPart!.StyleDefinitionsPart!.GetStream();
        Assert.Equal(changed, XDocument.Load(stream).Descendants(W + "pPrChange").Any());
        Assert.Equal(Spacing(missing), Spacing(RevisionProcessor.RejectRevisions(redline)));
        Assert.Equal(Spacing(stated), Spacing(RevisionProcessor.AcceptRevisions(redline)));
    }

    /// <summary>With a table present the defaults difference is recorded on each paragraph, a path that
    /// starts from the two documents' parsed defaults. Missing and empty must not read as the same.</summary>
    [Theory]
    [InlineData("", "<w:pPrDefault/>")]
    [InlineData("<w:pPrDefault/>", "")]
    public void CompareSeparatesMissingFromEmptyWhenATableIsPresent(string leftDefault, string rightDefault)
    {
        const string table = "<w:tbl><w:tblPr><w:tblW w:w='0' w:type='auto'/></w:tblPr><w:tblGrid><w:gridCol w:w='2400'/></w:tblGrid>" +
            "<w:tr><w:tc><w:tcPr><w:tcW w:w='2400' w:type='dxa'/></w:tcPr><w:p><w:r><w:t>Quartz04</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>";
        var left = Doc($"<w:docDefaults>{RunDefault}{leftDefault}</w:docDefaults>{Normal}</w:style>", table);
        var right = Doc($"<w:docDefaults>{RunDefault}{rightDefault}</w:docDefaults>{Normal}</w:style>", table);

        var redline = DocxCompare.Compare(left, right);

        Assert.NotEqual(Spacing(left), Spacing(right));
        Assert.Equal(Spacing(left), Spacing(RevisionProcessor.RejectRevisions(redline)));
        Assert.Equal(Spacing(right), Spacing(RevisionProcessor.AcceptRevisions(redline)));
    }
}
