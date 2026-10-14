// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
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
/// Content a comparison copies out of the revised document keeps its Word extension markup valid in
/// the output (issue #836). The original's part roots declare only <c>w:</c>. The revised document's
/// roots declare Word's extension namespaces and list them in <c>mc:Ignorable</c>, as Word writes
/// them, and a new list, a new style and a new paragraph use those namespaces — so every output part
/// that receives one of them must declare the namespaces too, ignorable and under Word's prefixes.
/// </summary>
public class DocxDiffImportedNamespaceTests
{
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private static readonly Dictionary<string, XNamespace> WordExtensions = new()
    {
        ["w14"] = "http://schemas.microsoft.com/office/word/2010/wordml",
        ["w15"] = "http://schemas.microsoft.com/office/word/2012/wordml",
        ["w16cid"] = "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    };

    private static readonly string WordRootAttributes =
        $" xmlns:mc=\"{Mc}\"" +
        string.Concat(WordExtensions.Select(extension => $" xmlns:{extension.Key}=\"{extension.Value}\"")) +
        " mc:Ignorable=\"w14 w15 w16cid\"";

    private const string ExistingItem =
        "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>" +
        "<w:r><w:t>Existing item</w:t></w:r></w:p>";

    private const string DecimalDefinition =
        "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/>" +
        "<w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>";

    private const string DecimalInstance = "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>";

    /// <summary>A new list as Word writes it: a w15 attribute, and a level whose number format needs w14.</summary>
    private const string BulletDefinition =
        "<w:abstractNum w:abstractNumId=\"1\" w15:restartNumberingAfterBreak=\"0\">" +
        "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"•\"/></w:lvl>" +
        "<w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><mc:AlternateContent><mc:Choice Requires=\"w14\">" +
        "<w:numFmt w:val=\"custom\" w:format=\"001, 002, 003, ...\"/></mc:Choice>" +
        "<mc:Fallback><w:numFmt w:val=\"decimal\"/></mc:Fallback></mc:AlternateContent>" +
        "<w:lvlText w:val=\"%2\"/></w:lvl></w:abstractNum>";

    private const string BulletInstance =
        "<w:num w:numId=\"2\" w16cid:durableId=\"1234567890\"><w:abstractNumId w:val=\"1\"/></w:num>";

    private static readonly WmlDocument Original = IrTestDocuments.FromParts(
        ExistingItem,
        numberingInnerXml: DecimalDefinition + DecimalInstance);

    private static readonly WmlDocument Revised = IrTestDocuments.FromParts(
        ExistingItem +
        "<w:p w14:paraId=\"1A2B3C4D\" w14:textId=\"77777777\"><w:pPr><w:numPr><w:ilvl w:val=\"0\"/>" +
        "<w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>New bullet</w:t></w:r></w:p>",
        stylesInnerXml:
            "<w:style w:type=\"character\" w:styleId=\"Ligatures\"><w:name w:val=\"Ligatures\"/>" +
            "<w:rPr><w14:ligatures w14:val=\"standard\"/></w:rPr></w:style>",
        numberingInnerXml: DecimalDefinition + BulletDefinition + DecimalInstance + BulletInstance,
        rootAttributes: WordRootAttributes);

    [Theory]
    [InlineData("word/numbering.xml", "w15 w16cid")] // the new list's definition and instance
    [InlineData("word/styles.xml", "w14")] // the new style
    [InlineData("word/document.xml", "w14")] // the new paragraph
    public void Compare_DeclaresImportedExtensionNamespacesOnThePartRootAsIgnorable(string partName, string prefixes)
    {
        var root = PartRoot(DocxCompare.Compare(Original, Revised), partName);

        foreach (var prefix in prefixes.Split(' '))
        {
            Assert.Equal(WordExtensions[prefix], root.GetNamespaceOfPrefix(prefix));
            Assert.Contains(prefix, IgnorablePrefixes(root));
        }
    }

    [Fact]
    public void Compare_ImportedAlternateContent_RequiresAPrefixTheOutputDeclares()
    {
        var numbering = PartRoot(DocxCompare.Compare(Original, Revised), "word/numbering.xml");
        var choice = numbering.Descendants(Mc + "Choice").Single();

        Assert.Equal(WordExtensions["w14"], choice.GetNamespaceOfPrefix((string)choice.Attribute("Requires")!));
    }

    [Fact]
    public void Compare_AddsNoValidationErrors() =>
        NoNewValidationErrors(Original.DocumentByteArray, DocxCompare.Compare(Original, Revised).DocumentByteArray);

    [Fact]
    public void Consolidate_AddsNoValidationErrors()
    {
        var reviewer = new DocxDiffReviewer { Author = "Reviewer", Document = Revised };

        NoNewValidationErrors(
            Original.DocumentByteArray, DocxDiff.Consolidate(Original, new[] { reviewer }).DocumentByteArray);
    }

    /// <summary>The part's root as serialized, so its declarations are the ones a consumer reads.</summary>
    private static XElement PartRoot(WmlDocument document, string partName)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        using var part = package.GetEntry(partName)!.Open();
        return XElement.Load(part);
    }

    private static string[] IgnorablePrefixes(XElement root) =>
        ((string?)root.Attribute(Mc + "Ignorable") ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries);
}
