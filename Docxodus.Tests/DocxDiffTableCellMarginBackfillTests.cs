#nullable enable

using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Word's compare output normalizes a FIXED-WIDTH table (<c>w:tblW w:type="dxa"</c>) that lacks
/// explicit cell margins: it materializes a hairline <c>w:tblCellMar</c> (left/right) plus a matching
/// <c>w:tblInd</c>, so the declared column widths render without a renderer's default ~108-twip cell
/// margin overflowing the fixed layout (LibreOffice otherwise insets cell text by its own default and
/// the whole table shifts vs Word). The inset equals the table's border width — a 0.5pt
/// (<c>w:sz="4"</c>) border → 10 twips — the value observed verbatim across every fixed-width table in
/// the Word-compare corpus. AUTO-width tables (<c>type="auto"</c>) get no such normalization, and a
/// table that already declares cell margins is left untouched. This mirrors the docDefaults backfill
/// (see <see cref="DocxDiffDocDefaultsBackfillTests"/>) — replicate what Word's compare emits.
/// </summary>
public class DocxDiffTableCellMarginBackfillTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static WmlDocument BuildDoc(string cellText, string tblWType, string tblWval,
        bool withCellMar, string borderSz = "4")
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var cellMar = withCellMar
                ? "<w:tblCellMar><w:left w:w=\"55\" w:type=\"dxa\"/><w:right w:w=\"55\" w:type=\"dxa\"/></w:tblCellMar>"
                : "";
            var borders = borderSz == "none"
                ? ""
                : $"<w:tblBorders><w:top w:val=\"single\" w:sz=\"{borderSz}\" w:color=\"000000\"/>" +
                  $"<w:left w:val=\"single\" w:sz=\"{borderSz}\" w:color=\"000000\"/>" +
                  $"<w:bottom w:val=\"single\" w:sz=\"{borderSz}\" w:color=\"000000\"/>" +
                  $"<w:right w:val=\"single\" w:sz=\"{borderSz}\" w:color=\"000000\"/>" +
                  $"<w:insideH w:val=\"single\" w:sz=\"{borderSz}\" w:color=\"000000\"/>" +
                  $"<w:insideV w:val=\"single\" w:sz=\"{borderSz}\" w:color=\"000000\"/></w:tblBorders>";
            var xml =
                $"<w:document xmlns:w=\"{W}\"><w:body>" +
                $"<w:tbl><w:tblPr><w:tblW w:w=\"{tblWval}\" w:type=\"{tblWType}\"/>{borders}{cellMar}</w:tblPr>" +
                "<w:tblGrid><w:gridCol w:w=\"3120\"/></w:tblGrid>" +
                "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"3120\" w:type=\"dxa\"/></w:tcPr>" +
                $"<w:p><w:r><w:t>{cellText}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                "<w:p><w:r><w:t>trailer</w:t></w:r></w:p>" +
                "</w:body></w:document>";
            using var s = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(s, new UTF8Encoding(false));
            writer.Write(xml);
        }
        return new WmlDocument("d.docx", stream.ToArray());
    }

    private static XElement OutputTblPr(WmlDocument result)
    {
        using var s = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(s, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.GetStream());
        return XDocument.Parse(reader.ReadToEnd()).Descendants(W + "tbl").First().Element(W + "tblPr")!;
    }

    [Fact]
    public void FixedWidthTable_NoCellMargin_BackfillsBorderWidthInset()
    {
        var left = BuildDoc("9:00 AM", "dxa", "3120", withCellMar: false);
        var right = BuildDoc("9:30 AM", "dxa", "3120", withCellMar: false);

        var tblPr = OutputTblPr(DocxDiff.Compare(left, right));

        var cm = tblPr.Element(W + "tblCellMar");
        Assert.NotNull(cm);
        Assert.Equal("10", (string?)cm!.Element(W + "left")?.Attribute(W + "w"));
        Assert.Equal("10", (string?)cm.Element(W + "right")?.Attribute(W + "w"));
        Assert.Equal("10", (string?)tblPr.Element(W + "tblInd")?.Attribute(W + "w"));
    }

    [Fact]
    public void FixedWidthTable_HeavierBorder_InsetTracksBorderWidth()
    {
        // 1pt border (w:sz="8") → 20 twips. The inset is derived from the border width, not a constant.
        var left = BuildDoc("9:00 AM", "dxa", "3120", withCellMar: false, borderSz: "8");
        var right = BuildDoc("9:30 AM", "dxa", "3120", withCellMar: false, borderSz: "8");

        var tblPr = OutputTblPr(DocxDiff.Compare(left, right));

        Assert.Equal("20", (string?)tblPr.Element(W + "tblCellMar")?.Element(W + "left")?.Attribute(W + "w"));
        Assert.Equal("20", (string?)tblPr.Element(W + "tblInd")?.Attribute(W + "w"));
    }

    [Fact]
    public void AutoWidthTable_NoCellMargin_LeftUntouched()
    {
        var left = BuildDoc("9:00 AM", "auto", "0", withCellMar: false);
        var right = BuildDoc("9:30 AM", "auto", "0", withCellMar: false);

        var tblPr = OutputTblPr(DocxDiff.Compare(left, right));

        Assert.Null(tblPr.Element(W + "tblCellMar"));
        Assert.Null(tblPr.Element(W + "tblInd"));
    }

    [Fact]
    public void FixedWidthTable_ExistingCellMargin_LeftUntouched()
    {
        var left = BuildDoc("9:00 AM", "dxa", "3120", withCellMar: true);
        var right = BuildDoc("9:30 AM", "dxa", "3120", withCellMar: true);

        var tblPr = OutputTblPr(DocxDiff.Compare(left, right));

        // The declared 55-twip margins are preserved, never overwritten with the inset.
        Assert.Equal("55", (string?)tblPr.Element(W + "tblCellMar")?.Element(W + "left")?.Attribute(W + "w"));
        Assert.Null(tblPr.Element(W + "tblInd"));
    }

    [Fact]
    public void FixedWidthTable_NoBorders_LeftUntouched()
    {
        // With no border width to derive the inset from, we do not guess — leave the table alone.
        var left = BuildDoc("9:00 AM", "dxa", "3120", withCellMar: false, borderSz: "none");
        var right = BuildDoc("9:30 AM", "dxa", "3120", withCellMar: false, borderSz: "none");

        var tblPr = OutputTblPr(DocxDiff.Compare(left, right));

        Assert.Null(tblPr.Element(W + "tblCellMar"));
        Assert.Null(tblPr.Element(W + "tblInd"));
    }
}
