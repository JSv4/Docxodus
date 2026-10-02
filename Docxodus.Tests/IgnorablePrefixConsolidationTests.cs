// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DocumentBuilder"/> and <see cref="FormattingAssembler"/> make the PowerTools namespaces
/// ignorable through <see cref="PartNamespaces.EnsureIgnorablePrefix"/> (issue #857): the namespace is
/// declared, listed once, and <c>mc:Ignorable</c> is created when the part root has none.
/// </summary>
public class IgnorablePrefixConsolidationTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void DocumentBuilder_SourceWithoutIgnorable_ListsThePowerToolsNamespaceItKeeps()
    {
        // A source main part with no mc:Ignorable whose paragraph carries a PowerTools attribute. The
        // builder's output root always starts with mc:Ignorable (FreshNamespaceAttributes), so this
        // guards the refactor rather than reproducing a defect.
        var source = Docx(
            $"<w:document xmlns:w='{WNs}' xmlns:pt='{PtOpenXml.pt.NamespaceName}'><w:body>" +
            "<w:p pt:Unid='a1'><w:r><w:t>Hello</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");

        var built = DocumentBuilder.BuildDocument(new List<Source> { new(source, true) });

        var root = MainRoot(built);
        Assert.True(root.Descendants().Attributes().Any(a => a.Name.Namespace == PtOpenXml.pt),
            "fixture precondition: the builder keeps the PowerTools attribute");
        Assert.Contains(root.GetPrefixOfNamespace(PtOpenXml.pt)!, IgnorablePrefixes(root));
    }

    [Fact]
    public void DocumentBuilder_TwoSourcesWithPowerToolsContent_ListEachPrefixOnce()
    {
        var xml =
            $"<w:document xmlns:w='{WNs}' xmlns:pt='{PtOpenXml.pt.NamespaceName}'><w:body>" +
            "<w:p pt:Unid='a1'><w:r><w:t>Hello</w:t></w:r></w:p><w:sectPr/></w:body></w:document>";

        var built = DocumentBuilder.BuildDocument(new List<Source> { new(Docx(xml), true), new(Docx(xml), true) });

        var prefixes = IgnorablePrefixes(MainRoot(built));
        Assert.Equal(prefixes.Distinct().Count(), prefixes.Count);
    }

    [Fact]
    public void FormattingAssembler_IgnorableAlreadyNamingALongerToken_StillListsPt14()
    {
        // "pt14x" contains "pt14" as a substring but is a different token.
        var root = XElement.Parse(
            $"<w:document xmlns:w='{WNs}' xmlns:mc='{MC.mc.NamespaceName}' xmlns:pt14x='urn:other' " +
            "mc:Ignorable='pt14x'><w:body/></w:document>");

        FormattingAssembler.NormalizePropsForPart(new XDocument(root),
            new FormattingAssemblerSettings { CreateHtmlConverterAnnotationAttributes = true });

        Assert.Contains("pt14", IgnorablePrefixes(root));
        Assert.Contains("pt14x", IgnorablePrefixes(root));
    }

    [Fact]
    public void EnsureIgnorablePrefix_NamespaceAlreadyDeclaredUnderAnotherPrefix_ListsThatPrefix()
    {
        var root = XElement.Parse($"<w:document xmlns:w='{WNs}' xmlns:p='{PtOpenXml.pt.NamespaceName}'/>");

        PartNamespaces.EnsureIgnorablePrefix(root, "pt", PtOpenXml.pt);

        Assert.Equal(new[] { "p" }, IgnorablePrefixes(root));
        Assert.Null(root.Attribute(XNamespace.Xmlns + "pt"));
    }

    [Fact]
    public void EnsureIgnorablePrefix_PrefixBoundToAnotherNamespace_IsNotRebound()
    {
        var root = XElement.Parse($"<w:document xmlns:w='{WNs}' xmlns:pt14='urn:other'/>");

        PartNamespaces.EnsureIgnorablePrefix(root, "pt14", PtOpenXml.pt);

        Assert.Equal("urn:other", root.GetNamespaceOfPrefix("pt14")!.NamespaceName);
        var listed = Assert.Single(IgnorablePrefixes(root));
        Assert.Equal(PtOpenXml.pt, root.GetNamespaceOfPrefix(listed));
    }

    private static List<string> IgnorablePrefixes(XElement root) =>
        ((string?)root.Attribute(MC.Ignorable) ?? string.Empty)
            .Split(' ', System.StringSplitOptions.RemoveEmptyEntries).ToList();

    private static XElement MainRoot(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        return XDocument.Load(package.MainDocumentPart!.GetStream()).Root!;
    }

    private static WmlDocument Docx(string documentXml)
    {
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create)))
                writer.Write(documentXml);
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            using (var writer = new StreamWriter(styles.GetStream(FileMode.Create)))
                writer.Write($"<w:styles xmlns:w='{WNs}'/>");
            var settings = main.AddNewPart<DocumentSettingsPart>();
            using (var writer = new StreamWriter(settings.GetStream(FileMode.Create)))
                writer.Write($"<w:settings xmlns:w='{WNs}'/>");
        }
        return new WmlDocument("source.docx", stream.ToArray());
    }
}
