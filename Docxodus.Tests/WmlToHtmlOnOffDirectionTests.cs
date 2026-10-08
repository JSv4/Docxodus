// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <c>w:rtl</c> and <c>w:bidiVisual</c> are ST_OnOff properties: <c>w:val="0"</c>, <c>"false"</c> or
/// <c>"off"</c> switches them off, and Word lays the content out left to right (issue #1011). Google Docs writes
/// <c>&lt;w:rtl w:val="0"/&gt;</c> on almost every run it exports, so reading the element's presence as "on"
/// wrapped ordinary text in right-to-left marks and let the browser reorder its digits and punctuation.
/// </summary>
public class WmlToHtmlOnOffDirectionTests
{
    private const string Rlm = "‏";

    private static byte[] Docx(string bodyXml)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                    bodyXml + "<w:sectPr/></w:body></w:document>");
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            doc.Save();
        }
        return stream.ToArray();
    }

    private static XElement Convert(string bodyXml) =>
        XElement.Parse(HtmlConversionOps.ConvertToHtml(Docx(bodyXml), new HtmlConversionOptions { FabricateCssClasses = false }));

    private static string RunText(string rtl) =>
        string.Concat(Convert($"<w:p><w:r><w:rPr>{rtl}</w:rPr><w:t>2026. 10. 07.</w:t></w:r></w:p>")
            .Descendants().Where(e => e.Name.LocalName == "p").Single().DescendantNodes().OfType<XText>()
            .Select(t => t.Value));

    [Theory]
    [InlineData("0")]
    [InlineData("false")]
    [InlineData("off")]
    public void RunWithRtlSwitchedOff_GetsNoRightToLeftMarks(string val) =>
        Assert.Equal("2026. 10. 07.", RunText($"<w:rtl w:val=\"{val}\"/>"));

    [Theory]
    [InlineData("<w:rtl/>")]
    [InlineData("<w:rtl w:val=\"1\"/>")]
    [InlineData("<w:rtl w:val=\"true\"/>")]
    [InlineData("<w:rtl w:val=\"on\"/>")]
    public void RunWithRtlOn_IsWrappedInRightToLeftMarks(string rtl) =>
        Assert.Equal(Rlm + "2026. 10. 07." + Rlm, RunText(rtl));

    private static string? TableDir(string bidiVisual) =>
        (string?)Convert(
                $"<w:tbl><w:tblPr>{bidiVisual}</w:tblPr><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
                "<w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>")
            .Descendants().Single(e => e.Name.LocalName == "table").Attribute("dir");

    [Theory]
    [InlineData("0")]
    [InlineData("false")]
    [InlineData("off")]
    public void TableWithBidiVisualSwitchedOff_IsLeftToRight(string val) =>
        Assert.Equal("ltr", TableDir($"<w:bidiVisual w:val=\"{val}\"/>"));

    [Theory]
    [InlineData("<w:bidiVisual/>")]
    [InlineData("<w:bidiVisual w:val=\"1\"/>")]
    public void TableWithBidiVisualOn_IsRightToLeft(string bidiVisual) =>
        Assert.Equal("rtl", TableDir(bidiVisual));
}
