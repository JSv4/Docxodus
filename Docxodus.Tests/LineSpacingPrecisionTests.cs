// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Exact and at-least line heights keep their full twentieth-of-a-point value (issue #850): a
/// one-decimal format turned <c>w:line="253"</c> (12.65 pt) into 12.7 pt, 0.05 pt per line that
/// accumulates down a page.
/// </summary>
public class LineSpacingPrecisionTests
{
    [Theory]
    [InlineData("exact", 253, "line-height: 12.65pt")]
    [InlineData("atLeast", 301, "line-height: 15.05pt")]
    [InlineData("exact", 200, "line-height: 10.0pt")]
    public void ExplicitLineHeight_KeepsTwentiethsOfAPoint(string rule, int line, string expected)
    {
        Assert.Contains(expected, Html(rule, line));
    }

    private static string Html(string rule, int line)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using var writer = new StreamWriter(main.GetStream(FileMode.Create));
            writer.Write($"<w:document xmlns:w='{w}'><w:body><w:p><w:pPr><w:spacing w:line='{line}' w:lineRule='{rule}'/></w:pPr>" +
                         "<w:r><w:t>Line.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("spacing.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false }).ToString();
    }
}
