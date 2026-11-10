// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Library output breaks lines with <c>"\n"</c> on every OS (issue #1659). The output cases fail
/// on Windows without the fix; on Linux, whose platform newline is already <c>"\n"</c>, the
/// guard is the <c>BannedSymbols.txt</c> analyzer ban that the last case pins. The helper cases
/// hold <see cref="LineEndings"/> to the .NET serializer it replaces, on both.
/// </summary>
public class LineEndingTests
{
    private static readonly byte[] Fixture = File.ReadAllBytes("../../../../TestFiles/HC031-Complicated-Document.docx");

    private static void AssertLfOnly(string output)
    {
        Assert.Contains('\n', output);
        Assert.DoesNotContain('\r', output);
    }

    [Fact]
    public void MarkdownProjection_BreaksLinesWithLf()
    {
        using var session = new DocxSession(Fixture);
        AssertLfOnly(session.Project().Markdown);
    }

    [Fact]
    public void Html_BreaksLinesWithLf()
    {
        // The CSS block is multi-line, and XML serialization rewrites line breaks in text too.
        AssertLfOnly(HtmlConversionOps.ConvertToHtml(Fixture, new HtmlConversionOptions()));
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("a.docx", Fixture), new WmlToHtmlConverterSettings());
        AssertLfOnly(WmlToHtmlConverter.ToHtmlString(html, indent: true));
    }

    [Fact]
    public void IrSnapshotJson_BreaksLinesWithLf() =>
        AssertLfOnly(IrDiagnosticJson.Write(IrReader.Read(new WmlDocument("a.docx", Fixture))));

    [Fact]
    public void RawXml_BreaksLinesWithLf()
    {
        using var session = new DocxSession(Fixture);
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Id.StartsWith("p:", StringComparison.Ordinal));
        AssertLfOnly(session.Raw.GetXml(anchor.Anchor.Id));
    }

    // Line breaks of every kind in text, a comment and an attribute, under a declaration.
    private static XDocument Sample() => XDocument.Parse(
        "<?xml version=\"1.0\" encoding=\"utf-8\"?><!--one\ntwo--><r xmlns:a=\"urn:a\"><a:c a:k=\"x&#xA;y\">t1&#xD;&#xA;t2&#xD;t3\nt4</a:c><a:c xmlns:a=\"urn:a\"/></r>",
        LoadOptions.PreserveWhitespace);

    public static TheoryData<SaveOptions> AllSaveOptions() => new()
    {
        SaveOptions.None,
        SaveOptions.DisableFormatting,
        SaveOptions.OmitDuplicateNamespaces,
        SaveOptions.DisableFormatting | SaveOptions.OmitDuplicateNamespaces,
    };

    [Theory]
    [MemberData(nameof(AllSaveOptions))]
    public void ToXmlString_IsXNodeToString_WithLfLineBreaks(SaveOptions options)
    {
        var document = Sample();
        XNode[] nodes = [document, document.Root!, document.Root!.Elements().First(), document.Root!.Elements().First().FirstNode!];
        foreach (var node in nodes)
        {
            var lf = node.ToXmlString(options);
            Assert.Equal(node.ToString(options).ReplaceLineEndings("\n"), lf);
            Assert.DoesNotContain('\r', lf);
        }
    }

    [Theory]
    [MemberData(nameof(AllSaveOptions))]
    public void SaveLf_IsXDocumentSave_WithLfLineBreaks(SaveOptions options)
    {
        var document = Sample();
        using var expectedStream = new MemoryStream();
        document.Save(expectedStream, options);
        using var actualStream = new MemoryStream();
        document.SaveLf(actualStream, options);
        var expectedWriter = new StringWriter();
        document.Save(expectedWriter, options);
        var actualWriter = new StringWriter();
        document.SaveLf(actualWriter, options);

        Assert.Equal(Encoding.UTF8.GetString(expectedStream.ToArray()).ReplaceLineEndings("\n"), Encoding.UTF8.GetString(actualStream.ToArray()));
        Assert.Equal(expectedWriter.ToString().ReplaceLineEndings("\n"), actualWriter.ToString());
        Assert.DoesNotContain((byte)'\r', actualStream.ToArray());
    }

    [Fact]
    public void AppendLf_EndsEachLineWithLf()
    {
        var sb = new StringBuilder().AppendLf("a").AppendLf($"{1}b").AppendLf();
        Assert.Equal("a\n1b\n\n", sb.ToString());
    }

    /// <summary>
    /// On Linux the output cases above pass whatever the library does, so this pins the
    /// compile-time guard instead: the library bans every API that defaults to the platform
    /// newline, and the ban fails the build in every configuration.
    /// </summary>
    [Fact]
    public void Library_BansThePlatformNewlineApis()
    {
        const string library = "../../../../Docxodus";
        var banned = File.ReadAllLines(Path.Combine(library, "BannedSymbols.txt"))
            .Where(l => !l.StartsWith("//", StringComparison.Ordinal))
            .Select(l => l.Split(';')[0])
            .ToHashSet(StringComparer.Ordinal);
        string[] required =
        [
            "M:System.Text.StringBuilder.AppendLine(System.String)",
            "M:System.Text.StringBuilder.AppendLine(System.Text.StringBuilder.AppendInterpolatedStringHandler@)",
            "P:System.Environment.NewLine",
            "M:System.Xml.Linq.XNode.ToString",
            "M:System.Xml.Linq.XNode.ToString(System.Xml.Linq.SaveOptions)",
            "M:System.Xml.Linq.XDocument.Save(System.IO.Stream)",
            "M:System.Xml.XmlWriterSettings.#ctor",
            "P:System.Text.Json.JsonSerializerOptions.WriteIndented",
            "P:System.Text.Json.JsonWriterOptions.Indented",
        ];
        Assert.Empty(required.Where(symbol => !banned.Contains(symbol)));

        var project = File.ReadAllText(Path.Combine(library, "Docxodus.csproj"));
        Assert.Contains("Microsoft.CodeAnalysis.BannedApiAnalyzers", project);
        Assert.Contains("<AdditionalFiles Include=\"BannedSymbols.txt\" />", project);
        Assert.Contains("<WarningsAsErrors>$(WarningsAsErrors);RS0030</WarningsAsErrors>", project);
    }
}
