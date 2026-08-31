#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Builds the three reference fields Word generates a table from — TOC, TOF (table of figures) and
/// TOA (table of authorities) — from typed options rather than a hand-written switch string
/// (issue #607).
/// </summary>
/// <remarks>
/// <para>A reference field is a <c>w:fldChar</c> begin/separate/end sequence around a
/// <c>w:instrText</c> carrying the instruction. It is written <b>dirty</b>
/// (<c>w:fldChar w:dirty="true"</c>) and the document asks for a field update on open
/// (<c>w:updateFields</c>), so Word paginates and fills the table itself rather than the library
/// shipping a cached result that is wrong the moment anything above it moves.</para>
/// <para>The switch strings are the entire reason this type exists. <c>\o "1-3"</c>, <c>\c
/// "Figure"</c>, <c>\e</c> and their kin are exactly the class of OOXML detail the library is for
/// hiding, and a malformed instruction renders as <em>nothing</em> in Word — silently.</para>
/// </remarks>
internal static class ReferenceFieldOps
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Word's own paragraph style for each table's entries, and for the TOC's title.</summary>
    internal const string TocEntryStyleId = "TOC1";
    internal const string TocHeadingStyleId = "TOCHeading";
    internal const string TofEntryStyleId = "TableofFigures";
    internal const string ToaEntryStyleId = "TableofAuthorities";

    /// <summary>Word's table-of-authorities categories, in their fixed numeric order. A TOA's
    /// <c>\c</c> switch selects one; the numbers are Word's, not ours.</summary>
    internal static readonly string[] AuthorityCategories =
    {
        "Cases", "Statutes", "Other Authorities", "Rules", "Treatises", "Regulations",
        "Constitutional Provisions",
    };

    /// <summary>Validate a level range such as <c>1-3</c> (or a single level such as <c>2</c>) and
    /// return the canonical form, or <c>null</c> with a reason.</summary>
    internal static string? NormalizeLevels(string levels, out string? error)
    {
        error = null;
        var trimmed = (levels ?? string.Empty).Trim();
        var parts = trimmed.Split('-').Select(part => part.Trim()).ToArray();
        bool Ok(string s, out int v) =>
            int.TryParse(s, NumberStyles.None, CultureInfo.InvariantCulture, out v) && v is >= 1 and <= 9;

        if (parts.Length == 1 && Ok(parts[0], out var only))
            return $"{only}-{only}";
        if (parts.Length == 2 && Ok(parts[0], out var from) && Ok(parts[1], out var to) && from <= to)
            return $"{from}-{to}";

        error = $"levels must be a heading level or range within 1-9 (e.g. \"1-3\"); got \"{levels}\"";
        return null;
    }

    /// <summary>The <c>TOC</c> instruction for the given options. <paramref name="levels"/> is
    /// already normalized by <see cref="NormalizeLevels"/>.</summary>
    internal static string TocInstruction(string levels, bool hyperlinks, bool hideTabAndPageNumbersInWeb,
        bool useOutlineLevels)
    {
        var sb = new StringBuilder("TOC");
        sb.Append(" \\o \"").Append(levels).Append('"');
        if (hyperlinks) sb.Append(" \\h");
        if (hideTabAndPageNumbersInWeb) sb.Append(" \\z");
        if (useOutlineLevels) sb.Append(" \\u");
        return sb.ToString();
    }

    /// <summary>The <c>TOC</c> instruction for a table of FIGURES: same field, selecting entries by
    /// caption label rather than by outline level. That is Word's own encoding — a table of figures
    /// is a TOC with <c>\c</c>, not a field of its own.</summary>
    internal static string TofInstruction(string captionLabel, bool hyperlinks)
    {
        var sb = new StringBuilder("TOC");
        sb.Append(" \\c \"").Append(EscapeQuotes(captionLabel)).Append('"');
        if (hyperlinks) sb.Append(" \\h");
        return sb.ToString();
    }

    /// <summary>The <c>TOA</c> instruction. <paramref name="category"/> is a 1-based index into
    /// <see cref="AuthorityCategories"/>.</summary>
    internal static string ToaInstruction(int category, bool hyperlinks, string? entryPageSeparator)
    {
        var sb = new StringBuilder("TOA");
        sb.Append(" \\c \"").Append(category.ToString(CultureInfo.InvariantCulture)).Append('"');
        if (hyperlinks) sb.Append(" \\h");
        if (entryPageSeparator is { Length: > 0 } separator)
            sb.Append(" \\e \"").Append(EscapeQuotes(separator)).Append('"');
        return sb.ToString();
    }

    /// <summary>
    /// A paragraph carrying the whole field: begin (dirty) → instruction → separate → end. There is
    /// no cached result between separate and end on purpose — an empty result is what Word writes
    /// for a field it has not evaluated, and it is what makes "update this table" the reader's first
    /// action rather than "why is this table wrong".
    /// </summary>
    internal static XElement FieldParagraph(string styleId, string instruction, int rightTabPos) =>
        new XElement(W + "p",
            new XElement(W + "pPr",
                new XElement(W + "pStyle", new XAttribute(W + "val", styleId)),
                new XElement(W + "tabs",
                    new XElement(W + "tab",
                        new XAttribute(W + "val", "right"),
                        new XAttribute(W + "leader", "dot"),
                        new XAttribute(W + "pos",
                            rightTabPos.ToString(CultureInfo.InvariantCulture)))),
                new XElement(W + "rPr", new XElement(W + "noProof"))),
            new XElement(W + "r",
                new XElement(W + "fldChar",
                    new XAttribute(W + "fldCharType", "begin"),
                    new XAttribute(W + "dirty", "true"))),
            new XElement(W + "r",
                new XElement(W + "instrText",
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    $" {instruction} ")),
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));

    /// <summary>The heading paragraph above a table of contents.</summary>
    internal static XElement TitleParagraph(string title) =>
        new XElement(W + "p",
            new XElement(W + "pPr",
                new XElement(W + "pStyle", new XAttribute(W + "val", TocHeadingStyleId))),
            new XElement(W + "r", new XElement(W + "t", title)));

    /// <summary>
    /// Wrap a table of contents in the <c>w:sdt</c> Word puts around one. The <c>docPartObj</c>
    /// gallery declaration is what gives the table its "Update Table" control in Word's UI; without
    /// it the field still renders and still updates, but only through the generic field-update path.
    /// </summary>
    internal static XElement TableOfContentsControl(IEnumerable<XElement> blocks) =>
        new XElement(W + "sdt",
            new XElement(W + "sdtPr",
                new XElement(W + "docPartObj",
                    new XElement(W + "docPartGallery",
                        new XAttribute(W + "val", "Table of Contents")),
                    new XElement(W + "docPartUnique"))),
            new XElement(W + "sdtContent", blocks));

    /// <summary>
    /// Ask Word to update every field when the document is opened. A reference field ships without a
    /// cached result, so without this the reader sees an empty table until they update it by hand.
    /// Idempotent: an existing declaration is set rather than duplicated.
    /// </summary>
    internal static void RequestFieldUpdateOnOpen(MainDocumentPart main)
    {
        var settingsPart = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
        var xDoc = settingsPart.GetXDocument();
        var root = xDoc.Root;
        if (root is null)
        {
            root = new XElement(W + "settings", new XAttribute(XNamespace.Xmlns + "w", W));
            xDoc.Add(root);
        }

        if (root.Element(W + "updateFields") is { } existing)
        {
            existing.SetAttributeValue(W + "val", "true");
            settingsPart.PutXDocument();
            return;
        }

        // Schema position, not append: CT_Settings is a sequence, and Word flags an out-of-order
        // child for repair. Same helper the note-separator declaration uses.
        if (WordprocessingMLUtil.EnsureSettingsChildInOrder(
                root, new XElement(W + "updateFields", new XAttribute(W + "val", "true"))))
            settingsPart.PutXDocument();
    }

    private static string EscapeQuotes(string value) => value.Replace("\"", "'", StringComparison.Ordinal);
}
