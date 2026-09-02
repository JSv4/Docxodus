// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class WmlDocumentTests
{
    private static WmlDocument CreateMinimalDocument()
    {
        using var stream = new MemoryStream();
        using (var wordDocument = WordprocessingDocument.Create(
            stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Wp.Document(
                new Wp.Body(new Wp.Paragraph(new Wp.Run(new Wp.Text("original")))));
        }
        return new WmlDocument("test.docx", stream.ToArray());
    }

    [Fact]
    public void ReplacementPartConstructorRoundTripsAnUnmodifiedPart()
    {
        var original = CreateMinimalDocument();
        var mainDocumentPart = original.MainDocumentPart;

        var replaced = new WmlDocument(original, mainDocumentPart);

        using var streamDoc = new OpenXmlMemoryStreamDocument(replaced);
        using var reopened = streamDoc.GetWordprocessingDocument();
        XElement? root = reopened.MainDocumentPart?.GetXDocument().Root;
        Assert.NotNull(root);
        Assert.Equal("original", root!.Value);
    }

    [Fact]
    public void ReplacementPartConstructorThrowsClearlyWhenUriMatchesNoPart()
    {
        var original = CreateMinimalDocument();

        // A replacement part whose pt:Uri attribute does not correspond to any part in the
        // package must fail with a clear error rather than a NullReferenceException.
        var bogusPart = new XElement(
            W.body,
            new XAttribute(PtOpenXml.Uri, new Uri("/word/does-not-exist.xml", UriKind.Relative)));

        var ex = Assert.Throws<DocxodusException>(() => new WmlDocument(original, bogusPart));
        Assert.Contains("does not match any part", ex.Message);
    }
}
