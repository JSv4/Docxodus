// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

/// <summary>
/// A <c>w:cnfStyle</c> or <c>w:color</c> keeps its required <c>w:val</c> through a comparison, together with
/// the attributes Word writes beside it: the 2010 per-flag attributes on a table-conditional paragraph's
/// <c>w:cnfStyle</c>, and the theme attributes on a <c>w:color</c> (issue #839). Every such element in the
/// redline is an attribute-for-attribute copy of one in an input, through the two-way and the consolidate
/// renderer, and the redline adds no Open XML validator errors absent from the inputs.
/// </summary>
public class DocxDiffRequiredAttributeTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string ConditionalCell =
        "w:val=\"101000000100\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\"" +
        " w:oddVBand=\"0\" w:evenVBand=\"0\" w:oddHBand=\"0\" w:evenHBand=\"0\" w:firstRowFirstColumn=\"1\"" +
        " w:firstRowLastColumn=\"0\" w:lastRowFirstColumn=\"0\" w:lastRowLastColumn=\"0\"";

    private const string DocumentPart = "word/document.xml";

    private static readonly string[] DocumentAndStyles = { DocumentPart, "word/styles.xml" };

    private const string ThemeColor = "w:val=\"2F5496\" w:themeColor=\"accent1\" w:themeShade=\"BF\"";

    private const string TableStyle =
        "<w:style w:type=\"table\" w:styleId=\"TableGrid\"><w:name w:val=\"Table Grid\"/></w:style>";

    private const string Intro = "<w:p><w:r><w:t>Intro.</w:t></w:r></w:p>";

    private static string Table(string text) =>
        "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/><w:tblW w:w=\"0\" w:type=\"auto\"/>" +
        "<w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\"" +
        " w:noHBand=\"0\" w:noVBand=\"1\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>" +
        "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>" +
        $"<w:p><w:pPr><w:cnfStyle {ConditionalCell}/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>" +
        "</w:tc></w:tr></w:tbl><w:p/>";

    private static string HeadingStyle(string color) =>
        "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/>" +
        $"<w:rPr><w:color {color}/></w:rPr></w:style>";

    private static string Heading(string text) =>
        $"<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

    private static string ColoredRun(string color) =>
        $"<w:p><w:r><w:rPr><w:color {color}/></w:rPr><w:t>Colored text.</w:t></w:r></w:p>";

    private static readonly Dictionary<string, (WmlDocument Original, WmlDocument Revised)> Scenarios = new()
    {
        ["TextEditInConditionalCell"] = (
            IrTestDocuments.FromParts(Table("Hello world"), TableStyle),
            IrTestDocuments.FromParts(Table("Hello brave world"), TableStyle)),
        ["NewTable"] = (
            IrTestDocuments.FromParts(Intro, TableStyle),
            IrTestDocuments.FromParts(Intro + Table("Hello"), TableStyle)),
        ["StyleColorPlainToTheme"] = (
            IrTestDocuments.FromParts(Heading("Title"), HeadingStyle("w:val=\"FF0000\"")),
            IrTestDocuments.FromParts(Heading("Title text"), HeadingStyle(ThemeColor))),
        ["RunColorPlainToTheme"] = (
            IrTestDocuments.FromParts(ColoredRun("w:val=\"FF0000\"")),
            IrTestDocuments.FromParts(ColoredRun(ThemeColor))),
    };

    public static TheoryData<string> ScenarioNames => new(Scenarios.Keys);

    [Theory]
    [MemberData(nameof(ScenarioNames))]
    public void Compare_KeepsValAndEveryOtherAttribute(string scenario)
    {
        var (original, revised) = Scenarios[scenario];
        AssertCopiedFromInputs(DocxCompare.Compare(original, revised), original, revised, DocumentAndStyles);
    }

    [Theory]
    [MemberData(nameof(ScenarioNames))]
    public void Consolidate_KeepsValAndEveryOtherAttribute(string scenario)
    {
        var (original, revised) = Scenarios[scenario];
        // Consolidate keeps the base's existing style definitions (it copies only the styles the base
        // lacks), so only the revised body's elements must reach the redline.
        AssertCopiedFromInputs(Consolidate(original, revised), original, revised, new[] { DocumentPart });
    }

    [Theory]
    [MemberData(nameof(ScenarioNames))]
    public void Compare_AddsNoValidationErrors(string scenario)
    {
        var (original, revised) = Scenarios[scenario];
        var output = DocxCompare.Compare(original, revised).DocumentByteArray;

        NoNewValidationErrors(original.DocumentByteArray, output);
        NoNewValidationErrors(revised.DocumentByteArray, output);
    }

    [Theory]
    [MemberData(nameof(ScenarioNames))]
    public void Consolidate_AddsNoValidationErrors(string scenario)
    {
        var (original, revised) = Scenarios[scenario];
        var output = Consolidate(original, revised).DocumentByteArray;

        NoNewValidationErrors(original.DocumentByteArray, output);
        NoNewValidationErrors(revised.DocumentByteArray, output);
    }

    private static WmlDocument Consolidate(WmlDocument original, WmlDocument revised) =>
        DocxDiff.Consolidate(original, new[] { new DocxDiffReviewer { Author = "Reviewer", Document = revised } });

    /// <summary>
    /// The redline carries every <c>w:cnfStyle</c>/<c>w:color</c> of the revised document's
    /// <paramref name="carriedParts"/>, and each such element in it has a <c>w:val</c> and exactly the attributes
    /// of one in an input — none dropped, none invented.
    /// </summary>
    private static void AssertCopiedFromInputs(
        WmlDocument redline, WmlDocument original, WmlDocument revised, string[] carriedParts)
    {
        var inputShapes = Properties(original).Concat(Properties(revised)).Select(Shape).ToHashSet();
        var output = Properties(redline).ToList();

        Assert.Subset(output.Select(Shape).ToHashSet(), Properties(revised, carriedParts).Select(Shape).ToHashSet());
        Assert.All(output, element =>
        {
            Assert.NotNull(element.Attribute(W + "val"));
            Assert.Contains(Shape(element), inputShapes);
        });
    }

    /// <summary>Every <c>w:cnfStyle</c> and <c>w:color</c> in the given parts (default: document and styles), as serialized.</summary>
    private static IEnumerable<XElement> Properties(WmlDocument document, string[]? parts = null)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        return (parts ?? DocumentAndStyles)
            .Select(name => package.GetEntry(name))
            .Where(entry => entry is not null)
            .SelectMany(entry =>
            {
                using var part = entry!.Open();
                return XElement.Load(part).Descendants().Where(e => e.Name == W + "cnfStyle" || e.Name == W + "color").ToList();
            })
            .ToList();
    }

    /// <summary>An element's name and its non-declaration attributes, in a stable order.</summary>
    private static string Shape(XElement element) =>
        element.Name + string.Concat(element.Attributes()
            .Where(a => !a.IsNamespaceDeclaration)
            .OrderBy(a => a.Name.ToString())
            .Select(a => $" {a.Name}={a.Value}"));
}
