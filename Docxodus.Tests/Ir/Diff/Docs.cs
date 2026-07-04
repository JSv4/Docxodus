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

    /// <summary>
    /// A normalized, table-aware structural projection of the MAIN document body, in document order: one
    /// tag per block — <c>"P:"+paragraph text</c> for each body <c>w:p</c>, and for each body <c>w:tbl</c>
    /// a <c>"TBL"</c> marker followed by a <c>"TC:"+cell text</c> tag per cell (descending row → cell →
    /// the cell's paragraph text). Unlike <see cref="PlainText"/> (which reads only the body's direct-child
    /// <c>w:p</c> and silently skips tables), this captures table presence and per-cell content, so a
    /// consolidate that corrupts or drops a table on the reject path produces a different projection.
    /// Walks only direct children of <c>w:body</c> so nested tables surface through their parent cell's text.
    /// Footnote text is intentionally NOT included here (footnote round-trip is asserted separately).
    /// </summary>
    public static string StructuralBody(WmlDocument d)
    {
        var ns = (XNamespace)IrTestDocuments.W;
        var doc = XDocument.Parse(MainPartXml(d));
        var body = doc.Root?.Element(ns + "body");
        if (body is null)
            return string.Empty;

        var sb = new StringBuilder();
        foreach (var block in body.Elements())
        {
            if (block.Name == ns + "p")
            {
                sb.Append("P:")
                  .Append(string.Concat(block.Descendants(ns + "t").Select(t => t.Value)))
                  .Append('\n');
            }
            else if (block.Name == ns + "tbl")
            {
                sb.Append("TBL\n");
                foreach (var tr in block.Elements(ns + "tr"))
                    foreach (var tc in tr.Elements(ns + "tc"))
                        sb.Append("TC:")
                          .Append(string.Concat(tc.Descendants(ns + "t").Select(t => t.Value)))
                          .Append('\n');
            }
            // Other body-level blocks (e.g. sectPr) are ignored — they are not part of the
            // content structure this oracle compares.
        }
        return sb.ToString();
    }

    /// <summary>
    /// Body text INCLUDING table cell text, in document order: each body <c>w:p</c>'s run text, and each
    /// body <c>w:tbl</c>'s cell text (row → cell → the cell's paragraph text), all joined by newlines.
    /// Unlike <see cref="PlainText"/> (which skips tables) this surfaces table content, so an accept that
    /// composes table-cell edits is observable. Reads only direct children of <c>w:body</c>; nested tables
    /// surface through their parent cell's text.
    /// </summary>
    public static string PlainTextWithTables(WmlDocument d)
    {
        var ns = (XNamespace)IrTestDocuments.W;
        var doc = XDocument.Parse(MainPartXml(d));
        var body = doc.Root?.Element(ns + "body");
        if (body is null)
            return string.Empty;

        var sb = new StringBuilder();
        foreach (var block in body.Elements())
        {
            if (block.Name == ns + "p")
            {
                sb.Append(string.Concat(block.Descendants(ns + "t").Select(t => t.Value))).Append('\n');
            }
            else if (block.Name == ns + "tbl")
            {
                foreach (var tr in block.Elements(ns + "tr"))
                    foreach (var tc in tr.Elements(ns + "tc"))
                        sb.Append(string.Concat(tc.Descendants(ns + "t").Select(t => t.Value))).Append('\n');
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Apply all tracked revisions (accept) to <paramref name="merged"/> and project the resulting body to
    /// the SAME table-aware text shape <see cref="IrCompositeVerifier"/> reconstructs from the composite
    /// script — so the verifier's apply-reconstruction can be checked against the rendered accepted body for
    /// table-bearing documents. The shape: one fragment per body block (a paragraph's text; a table's cell
    /// text in row → cell order), joined by newlines. Whitespace is later collapsed by
    /// <see cref="RevisionEquivalence.Normalize"/> on both sides, so intra-table delimiters need only match
    /// loosely.
    /// </summary>
    public static string AcceptStructuralBody(WmlDocument merged) =>
        PlainTextWithTables(RevisionAccepter.AcceptRevisions(merged));

    // ------------------------------------------------------------------ block-format (shell/section) projection

    private static readonly XNamespace WNs = IrTestDocuments.W;

    /// <summary>
    /// A canonical, document-order projection of every body block-format SHELL and the section properties —
    /// the byte-level oracle the format-blind text projections (<see cref="PlainText"/>/<see cref="StructuralBody"/>)
    /// lack. Emits, per body table, its <c>w:tblPr</c> and <c>w:tblGrid</c>; per row, its <c>w:trPr</c> and
    /// <c>w:tblPrEx</c>; per cell, its <c>w:tcPr</c>; per paragraph with an inline <c>w:pPr/w:sectPr</c>, that
    /// section's properties; and the trailing body <c>w:sectPr</c>'s properties. Each shell is normalized by
    /// <see cref="NormShell"/> (rsid/unid stripped, <c>w:*Change</c> markers removed, attributes sorted) so that
    /// two documents with the same block formatting project to the same string regardless of revision markup or
    /// non-semantic id noise. A dropped, retained-but-should-be-reverted, or corrupted shell/section therefore
    /// changes the projection — which the text projections cannot see. Used for byte-level
    /// <c>reject ≡ base</c> / <c>accept ≡ winner</c> assertions in the consolidate block-format tests.
    /// </summary>
    public static string ShellSection(WmlDocument d)
    {
        var body = XDocument.Parse(MainPartXml(d)).Root?.Element(WNs + "body");
        if (body is null)
            return string.Empty;

        var sb = new StringBuilder();
        foreach (var block in body.Elements())
        {
            if (block.Name == WNs + "p")
            {
                var inlineSect = block.Element(WNs + "pPr")?.Element(WNs + "sectPr");
                if (inlineSect != null)
                    sb.Append("PSECT{").Append(NormSectProps(inlineSect)).Append("}\n");
            }
            else if (block.Name == WNs + "tbl")
            {
                sb.Append("TBL{tblPr:").Append(NormShell(block.Element(WNs + "tblPr")))
                  .Append(";grid:").Append(NormShell(block.Element(WNs + "tblGrid"))).Append("}\n");
                foreach (var tr in block.Elements(WNs + "tr"))
                {
                    var trPr = tr.Element(WNs + "trPr");
                    sb.Append("TR{trPr:").Append(NormShell(trPr))
                      .Append(";ex:").Append(NormShell(trPr?.Element(WNs + "tblPrEx") ?? tr.Element(WNs + "tblPrEx")))
                      .Append("}\n");
                    foreach (var tc in tr.Elements(WNs + "tc"))
                        sb.Append("TC{").Append(NormShell(tc.Element(WNs + "tcPr"))).Append("}\n");
                }
            }
            else if (block.Name == WNs + "sectPr")
            {
                sb.Append("SECT{").Append(NormSectProps(block)).Append("}\n");
            }
        }
        return sb.ToString();
    }

    /// <summary>Canonicalize a block-format shell element for byte-level comparison: null/empty → <c>"∅"</c>;
    /// otherwise a recursively rsid/unid-stripped, <c>w:*Change</c>-free, attribute-sorted rendering. Empty ≡
    /// absent (mirrors the reader's shell-children digest), so a render→reject cycle's empty <c>&lt;w:trPr/&gt;</c>
    /// equals base's absent one.</summary>
    private static string NormShell(XElement? shell)
    {
        if (shell is null)
            return "∅";
        var norm = Canonicalize(shell);
        return norm is null || !norm.HasElements && !norm.HasAttributes ? "∅" : norm.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Canonicalize the SECTION properties (a <c>w:sectPr</c>) excluding header/footer references and
    /// the change marker — matching the two-way engine's <c>IsSectPrProp</c> contract, since those references are
    /// owned by the header/footer machinery and are outside the tracked <c>w:sectPrChange</c>.</summary>
    private static string NormSectProps(XElement sectPr)
    {
        var props = new XElement(WNs + "sectPr",
            sectPr.Elements().Where(e =>
                e.Name != WNs + "headerReference" &&
                e.Name != WNs + "footerReference" &&
                e.Name != WNs + "sectPrChange"));
        var norm = Canonicalize(props);
        return norm is null ? "∅" : norm.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Recursively strip <c>w:rsid*</c> / <c>pt:*</c>(unid) attributes and <c>w:*Change</c> child
    /// elements, and sort each element's attributes by name — the minimal canonical form for comparing two
    /// shells for property-byte equality.</summary>
    private static XElement Canonicalize(XElement el)
    {
        var keptAttrs = el.Attributes()
            .Where(a => !a.IsNamespaceDeclaration
                && !(a.Name.Namespace == WNs && a.Name.LocalName.StartsWith("rsid", System.StringComparison.Ordinal))
                && a.Name.NamespaceName != "http://powertools.codeplex.com/2011")
            .OrderBy(a => a.Name.NamespaceName, System.StringComparer.Ordinal)
            .ThenBy(a => a.Name.LocalName, System.StringComparer.Ordinal)
            .Select(a => new XAttribute(a.Name, a.Value));

        var keptChildren = el.Elements()
            .Where(c => !c.Name.LocalName.EndsWith("Change", System.StringComparison.Ordinal))
            .Select(Canonicalize);

        return new XElement(el.Name, keptAttrs, keptChildren);
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
