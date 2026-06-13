#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Tests.Ir;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Small DOCX fixtures + readback helpers for the composite-merger tests. <see cref="Para"/>
/// builds an IrReader-clean one-section document (one single-run paragraph per supplied string),
/// delegating to <see cref="IrTestDocuments.Create"/> so the required StyleDefinitionsPart /
/// DocumentSettingsPart are present. <see cref="PlainText"/> and <see cref="MainPartXml"/> read a
/// document back for assertions used by later composite-merge tasks.
/// </summary>
internal static class Docs
{
    /// <summary>A one-section DOCX whose body holds one single-run paragraph per supplied string.</summary>
    public static WmlDocument Para(params string[] paragraphs) => IrTestDocuments.Create(paragraphs);

    /// <summary>Body paragraph text, paragraphs joined by newline (run text concatenated per paragraph).</summary>
    public static string PlainText(WmlDocument d)
    {
        var ns = (XNamespace)IrTestDocuments.W;
        var doc = XDocument.Parse(MainPartXml(d));
        var body = doc.Root?.Element(ns + "body");
        if (body is null)
            return string.Empty;
        var paras = body.Elements(ns + "p")
            .Select(p => string.Concat(p.Descendants(ns + "t").Select(t => t.Value)));
        return string.Join("\n", paras);
    }

    /// <summary>The main document part XML as a string.</summary>
    public static string MainPartXml(WmlDocument d)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var main = wDoc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no MainDocumentPart.");
        using var partStream = main.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(partStream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
