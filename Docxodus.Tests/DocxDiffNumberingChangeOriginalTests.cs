// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

/// <summary>
/// A comparison records the label a list item displayed before an insertion, deletion or move shifted
/// its counter in <c>w:numberingChange/@w:original</c>, which the schema caps at 15 characters
/// (issue #861). Spelled-out numbers, letter-repeated numbers and long level texts exceed that, so the
/// stored label is shortened to 14 characters and an ellipsis.
/// </summary>
public class DocxDiffNumberingChangeOriginalTests
{
    /// <summary>List levels whose labels run past 15 characters, each with the old label of the first item.</summary>
    public static TheoryData<string, string, string> LongLabelLevels => new()
    {
        { "cardinalText", "<w:start w:val=\"3500\"/><w:numFmt w:val=\"cardinalText\"/><w:lvlText w:val=\"%1\"/>", "Three thousand…" },
        { "lowerLetter", "<w:start w:val=\"390\"/><w:numFmt w:val=\"lowerLetter\"/><w:lvlText w:val=\"%1.\"/>", "zzzzzzzzzzzzzz…" },
        { "long lvlText", "<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"Paragraph number %1.\"/>", "Paragraph numb…" },
    };

    [Theory]
    [MemberData(nameof(LongLabelLevels))]
    public void Compare_InsertedItem_StoresAShortenedOriginalLabel(string format, string level, string expected)
    {
        _ = format;
        var redline = DocxCompare.Compare(List(level, "Alpha", "Bravo"), List(level, "New", "Alpha", "Bravo"));

        var originals = Originals(redline);
        Assert.Equal(expected, originals[0]);
        Assert.All(originals, original => Assert.InRange(original.Length, 1, 15));
    }

    [Theory]
    [MemberData(nameof(LongLabelLevels))]
    public void Compare_LongLabels_AddNoValidationErrors(string format, string level, string expected)
    {
        _ = (format, expected);
        var original = List(level, "Alpha", "Bravo", "Charlie", "Delta");
        var revised = List(level, "New", "Alpha", "Charlie", "Delta", "Bravo"); // insert, delete-or-move, shift

        NoNewValidationErrors(original.DocumentByteArray, DocxCompare.Compare(original, revised).DocumentByteArray);
    }

    [Theory]
    [MemberData(nameof(LongLabelLevels))]
    public void Consolidate_LongLabels_AddNoValidationErrors(string format, string level, string expected)
    {
        _ = (format, expected);
        var original = List(level, "Alpha", "Bravo", "Charlie");
        var reviewer = new DocxDiffReviewer { Author = "Reviewer", Document = List(level, "New", "Alpha", "Charlie") };

        var consolidated = DocxDiff.Consolidate(original, new[] { reviewer });

        Assert.NotEmpty(Originals(consolidated));
        NoNewValidationErrors(original.DocumentByteArray, consolidated.DocumentByteArray);
    }

    [Fact]
    public void Compare_ShortLabels_AreStoredUnchanged()
    {
        const string level = "<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/>";

        var redline = DocxCompare.Compare(List(level, "Alpha", "Bravo"), List(level, "New", "Alpha", "Bravo"));

        Assert.Equal(new[] { "1.", "2." }, Originals(redline));
    }

    [Theory]
    [InlineData("", "")]
    [InlineData("1.", "1.")]
    [InlineData("fifteen chars!!", "fifteen chars!!")]
    [InlineData("sixteen chars!!!", "sixteen chars!…")]
    [InlineData("Three thousand five hundred", "Three thousand…")]
    [InlineData("1234567890123\U0001F600xyz", "1234567890123…")] // a surrogate pair straddling the cut is dropped whole
    public void NumberingChangeOriginal_FitsTheSchemaLimit(string label, string expected) =>
        Assert.Equal(expected, IrMarkupRenderer.NumberingChangeOriginal(label));

    private static WmlDocument List(string level, params string[] items) => IrTestDocuments.FromParts(
        string.Concat(items.Select(text =>
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>" +
            $"<w:r><w:t>{text}</w:t></w:r></w:p>")),
        numberingInnerXml:
            $"<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\">{level}</w:lvl></w:abstractNum>" +
            "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>");

    private static string[] Originals(WmlDocument document)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(document.DocumentByteArray), false);
        return package.MainDocumentPart!.GetXDocument()
            .Descendants(W.numberingChange)
            .Select(change => (string)change.Attribute(W.original)!)
            .ToArray();
    }
}
