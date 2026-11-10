using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1576: which side's default paragraph spacing a redline keeps when the two documents declare
/// different spacing in <c>w:docDefaults</c> or in the default paragraph style. Whatever the review view
/// shows, accepting every change must give the revised document's spacing and rejecting every change
/// the original's. See docs/ooxml_corner_cases.md, "Compare: documents with different default
/// paragraph spacing".
/// </summary>
public class DocxCompareDefaultSpacingTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string Wide = "<w:spacing w:after='200' w:line='276' w:lineRule='auto'/>";
    private const string Tight = "<w:spacing w:after='0' w:line='240' w:lineRule='auto'/>";

    private static readonly string[] Original =
    {
        "The ferry leaves the east dock at seven each morning.",
        "Tickets can be bought from the kiosk beside the gate.",
        "Bicycles travel free on the lower deck.",
    };

    private static readonly string[] Revised =
    {
        Original[0],
        "Tickets can be bought from the booth beside the gate.",
        Original[2],
    };

    /// <summary>Where a document declares its default paragraph spacing.</summary>
    public enum Store { DocDefaults, NormalStyle, None }

    private static WmlDocument Doc(string? spacing, Store store, IEnumerable<string> paragraphs)
    {
        var defaults = spacing is not null && store == Store.DocDefaults
            ? $"<w:pPrDefault><w:pPr>{spacing}</w:pPr></w:pPrDefault>"
            : "";
        var normal = spacing is not null && store == Store.NormalStyle ? $"<w:pPr>{spacing}</w:pPr>" : "";
        var styles = "<w:docDefaults><w:rPrDefault><w:rPr>" +
            "<w:rFonts w:ascii='Liberation Serif' w:hAnsi='Liberation Serif' w:cs='Liberation Serif'/>" +
            $"<w:sz w:val='24'/></w:rPr></w:rPrDefault>{defaults}</w:docDefaults>" +
            $"<w:style w:type='paragraph' w:default='1' w:styleId='Normal'><w:name w:val='Normal'/>{normal}</w:style>";
        var body = string.Concat(paragraphs.Select(text => $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>")) +
            "<w:sectPr><w:pgSz w:w='12240' w:h='15840'/>" +
            "<w:pgMar w:top='1440' w:right='1440' w:bottom='1440' w:left='1440' w:header='720' w:footer='720' w:gutter='0'/></w:sectPr>";
        using var memory = new MemoryStream();
        using (var package = WordprocessingDocument.Create(memory, WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using (var stream = main.GetStream())
                XDocument.Parse($"<w:document xmlns:w='{W}'><w:body>{body}</w:body></w:document>").Save(stream);
            using (var stream = main.AddNewPart<StyleDefinitionsPart>().GetStream())
                XDocument.Parse($"<w:styles xmlns:w='{W}'>{styles}</w:styles>").Save(stream);
            using (var stream = main.AddNewPart<DocumentSettingsPart>().GetStream())
                XDocument.Parse($"<w:settings xmlns:w='{W}'/>").Save(stream);
        }
        return new WmlDocument("default-spacing.docx", memory.ToArray());
    }

    /// <summary>Each body paragraph's effective after/line spacing, resolved through docDefaults and the
    /// style chain by <see cref="FormattingAssembler"/>, so the check is independent of where the spacing
    /// is stored.</summary>
    private static string[] Spacing(WmlDocument document)
    {
        var assembled = FormattingAssembler.AssembleFormatting(document, new FormattingAssemblerSettings
        {
            CreateHtmlConverterAnnotationAttributes = false,
        });
        using var memory = new MemoryStream(assembled.DocumentByteArray);
        using var package = WordprocessingDocument.Open(memory, false);
        using var stream = package.MainDocumentPart!.GetStream();
        return XDocument.Load(stream).Descendants(W + "p").Select(paragraph =>
        {
            var spacing = paragraph.Element(W + "pPr")?.Element(W + "spacing");
            return $"{(string?)spacing?.Attribute(W + "after") ?? "0"}/{(string?)spacing?.Attribute(W + "line") ?? "240"}";
        }).ToArray();
    }

    private static XElement Normal(WmlDocument document)
    {
        using var memory = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(memory, false);
        using var stream = package.MainDocumentPart!.StyleDefinitionsPart!.GetStream();
        return XDocument.Load(stream).Root!.Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == "Normal");
    }

    private static void AssertRoundTrips(WmlDocument left, WmlDocument right)
    {
        var redline = DocxCompare.Compare(left, right);
        // Word's review view shows the revised document's spacing; so does the redline's current formatting.
        Assert.Equal(Spacing(right), Spacing(redline));
        Assert.Equal(Spacing(right), Spacing(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(Spacing(left), Spacing(RevisionProcessor.RejectRevisions(redline)));
        using var stream = new MemoryStream(redline.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package));
    }

    public static IEnumerable<object?[]> Pairs() => new[]
    {
        new object?[] { Wide, Store.DocDefaults, Tight, Store.DocDefaults },
        new object?[] { Tight, Store.DocDefaults, Wide, Store.DocDefaults },
        new object?[] { Wide, Store.DocDefaults, null, Store.None },
        new object?[] { null, Store.None, Wide, Store.DocDefaults },
        new object?[] { Wide, Store.NormalStyle, Tight, Store.NormalStyle },
        new object?[] { Tight, Store.NormalStyle, Wide, Store.NormalStyle },
    };

    [Theory]
    [MemberData(nameof(Pairs))]
    public void AcceptGivesRevisedSpacing_RejectGivesOriginal(string? left, Store leftStore, string? right, Store rightStore) =>
        AssertRoundTrips(Doc(left, leftStore, Original), Doc(right, rightStore, Revised));

    /// <summary>
    /// Word's Reject All restores a style from BOTH of its property-change markers. Measured on these
    /// pairs: when the Normal style's spacing revision had no <c>w:rPrChange</c>, Word reset Normal's run
    /// properties to its built-in 10 pt on Reject All, so the original came back with the right spacing
    /// but smaller text. The revision must therefore archive the original's run properties as well.
    /// The Normal-style pairs are the control: their differing definitions always archived both.
    /// </summary>
    [Theory]
    [MemberData(nameof(Pairs))]
    public void SpacingRevisionArchivesTheOriginalsRunPropertiesToo(string? left, Store leftStore, string? right, Store rightStore)
    {
        var normal = Normal(DocxCompare.Compare(Doc(left, leftStore, Original), Doc(right, rightStore, Revised)));
        Assert.NotNull(normal.Element(W + "pPr")?.Element(W + "pPrChange"));
        Assert.Equal("24", (string?)normal.Element(W + "rPr")?.Element(W + "rPrChange")?.Element(W + "rPr")
            ?.Element(W + "sz")?.Attribute(W + "val"));
    }

    [Theory]
    [MemberData(nameof(Pairs))]
    public void RoundTripsWhenTheRevisedDocumentAddsAParagraph(string? left, Store leftStore, string? right, Store rightStore) =>
        AssertRoundTrips(Doc(left, leftStore, Original),
            Doc(right, rightStore, Revised.Append("Dogs must be kept on a lead at all times.")));
}
