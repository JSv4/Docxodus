// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;

namespace Docxodus.Tests.Eval;

/// <summary>
/// Programmatic fixture builders, for the one corpus need the step format cannot express: the
/// tool surface fills content controls but has no action that creates one, so a template
/// fixture is authored here instead of as a step script. Builders must be deterministic —
/// EV003 builds every fixture twice and requires the same document — and, like the step
/// scripts, they keep committed document bytes out of the corpus.
/// </summary>
internal static class EvalFixtureBuilders
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

    public static byte[] Build(string builder) => builder switch
    {
        "content-control-template" => ContentControlTemplate(),
        _ => throw new InvalidOperationException($"unknown fixture builder '{builder}'"),
    };

    /// <summary>
    /// An engagement-letter template: prose around a plain-text control (tag
    /// <c>client-name</c>), a drop-down (tag <c>governing-law</c>), and a checkbox (tag
    /// <c>include-arbitration</c>), each carrying visible placeholder content a fill must
    /// replace.
    /// </summary>
    private static byte[] ContentControlTemplate()
    {
        using var stream = new MemoryStream();
        stream.Write(DocxSessionOps.CreateBlankDocx());
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W + "body")!;
            body.Elements().Where(element => element.Name != W + "sectPr").Remove();
            body.AddFirst(
                Paragraph("ENGAGEMENT LETTER TEMPLATE"),
                Paragraph("This engagement letter is entered into by the client named below."),
                BlockSdt("201", "client-name", "Client name",
                    new XElement(W + "text"), "[CLIENT NAME]"),
                BlockSdt("202", "governing-law", "Governing law",
                    new XElement(
                        W + "dropDownList",
                        Item("New York"),
                        Item("Delaware")),
                    "[GOVERNING LAW]"),
                BlockSdt("203", "include-arbitration", "Include arbitration",
                    new XElement(
                        W14 + "checkbox",
                        new XElement(W14 + "checked", new XAttribute(W14 + "val", "0")),
                        new XElement(W14 + "checkedState", new XAttribute(W14 + "val", "2612")),
                        new XElement(W14 + "uncheckedState", new XAttribute(W14 + "val", "2610"))),
                    "☐"),
                Paragraph("Signature follows on the final page."));
            main.PutXDocument();
        }

        // One session round-trip puts the hand-authored XML into saved-normal form. Without it,
        // the scenario's own save would renormalize every paragraph and the change set would
        // report the whole template as modified by a three-control fill.
        using var session = new DocxSession(stream.ToArray());
        return session.Save();
    }

    private static XElement Paragraph(string text) =>
        new(W + "p", new XElement(W + "r", new XElement(W + "t", text)));

    private static XElement BlockSdt(
        string id, string tag, string alias, XElement type, string placeholder) =>
        new(W + "sdt",
            new XElement(W + "sdtPr",
                new XElement(W + "id", new XAttribute(W + "val", id)),
                new XElement(W + "tag", new XAttribute(W + "val", tag)),
                new XElement(W + "alias", new XAttribute(W + "val", alias)),
                type),
            new XElement(W + "sdtContent", Paragraph(placeholder)));

    private static XElement Item(string display) =>
        new(W + "listItem",
            new XAttribute(W + "displayText", display),
            new XAttribute(W + "value", display));
}
