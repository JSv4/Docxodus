#nullable enable

using System;
using System.IO;
using System.Text;
using System.Text.Json;

namespace Docxodus.Ir;

/// <summary>
/// Renders an <see cref="IrDocument"/> as stable, human-readable, indented JSON for tests and
/// debugging (spec §9). The output captures anchors, hashes (lowercase hex), modeled format
/// records, text, and opaque element names.
/// </summary>
/// <remarks>
/// This is a <strong>diagnostic format, not a versioned contract</strong>. It exists for
/// golden-snapshot conformance tests and for inspecting "what the IR thinks this document is".
/// Snapshots are regenerated under review whenever the IR evolves; nothing should persist this
/// JSON or treat its shape as stable across library versions (spec §9, §11).
/// <para/>
/// Determinism is the design requirement, so the writer is <em>hand-written</em> (one method per
/// node type) rather than reflection-based: field order is fixed in declaration order, hashes are
/// lowercase hex, numbers use the invariant culture, and there are no timestamps or machine paths.
/// Two reads of the same bytes therefore produce byte-identical JSON.
/// </remarks>
internal static class IrDiagnosticJson
{
    private static readonly JsonWriterOptions Options = new() { Indented = true };

    /// <summary>
    /// Render <paramref name="document"/>'s scopes as diagnostic JSON. The top level is
    /// <c>{"scopes":[…]}</c> with entries in a fixed order: <c>body</c>, then each header
    /// (<c>hdr1</c>…) and footer (<c>ftr1</c>…) in part order, then footnotes (<c>fn</c>) and
    /// endnotes (<c>en</c>), then comments (<c>cmt</c>). A body-only document renders just the body
    /// entry; non-body scopes appear only when they hold content.
    /// </summary>
    public static string Write(IrDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer, Options))
        {
            writer.WriteStartObject();
            writer.WriteStartArray("scopes");

            WriteBlockScope(writer, document.Body.Name, document.Body.Blocks);

            foreach (var header in document.Headers)
                WriteBlockScope(writer, header.ScopeName, header.Scope.Blocks, header.Kind);
            foreach (var footer in document.Footers)
                WriteBlockScope(writer, footer.ScopeName, footer.Scope.Blocks, footer.Kind);

            WriteNoteScope(writer, "fn", document.Footnotes);
            WriteNoteScope(writer, "en", document.Endnotes);
            WriteCommentScope(writer, document.Comments);

            writer.WriteEndArray();
            writer.WriteEndObject();
        }

        return Encoding.UTF8.GetString(buffer.ToArray());
    }

    // --- scopes -----------------------------------------------------------

    private static void WriteBlockScope(Utf8JsonWriter writer, string scope,
        IrNodeList<IrBlock> blocks, IrHeaderFooterKind? kind = null)
    {
        writer.WriteStartObject();
        writer.WriteString("scope", scope);
        if (kind is { } k)
            writer.WriteString("kind", k.ToString());
        writer.WriteStartArray("blocks");
        foreach (var block in blocks)
            WriteBlock(writer, block);
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    private static void WriteNoteScope(Utf8JsonWriter writer, string scope, IrNoteStore store)
    {
        if (store.Notes.Count == 0)
            return;

        writer.WriteStartObject();
        writer.WriteString("scope", scope);
        writer.WriteStartObject("notes");
        // Deterministic order: notes sorted by numeric id when possible, else ordinal by string.
        foreach (var id in SortNoteIds(store.Notes.Keys))
        {
            writer.WriteStartArray(id);
            foreach (var block in store.Notes[id].Blocks)
                WriteBlock(writer, block);
            writer.WriteEndArray();
        }
        writer.WriteEndObject();
        writer.WriteEndObject();
    }

    private static System.Collections.Generic.IEnumerable<string> SortNoteIds(
        System.Collections.Generic.IEnumerable<string> ids)
    {
        var list = new System.Collections.Generic.List<string>(ids);
        list.Sort((a, b) =>
        {
            bool aNum = int.TryParse(a, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var ai);
            bool bNum = int.TryParse(b, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var bi);
            if (aNum && bNum)
                return ai.CompareTo(bi);
            if (aNum != bNum)
                return aNum ? -1 : 1; // numeric ids before non-numeric.
            return string.CompareOrdinal(a, b);
        });
        return list;
    }

    private static void WriteCommentScope(Utf8JsonWriter writer, IrCommentStore store)
    {
        if (store.Comments.Count == 0)
            return;

        writer.WriteStartObject();
        writer.WriteString("scope", "cmt");
        writer.WriteStartArray("comments");
        foreach (var comment in store.Comments)
        {
            writer.WriteStartObject();
            writer.WriteString("anchor", comment.Anchor.ToString());
            writer.WriteString("author", comment.Author);
            if (comment.Initials is { } initials) writer.WriteString("initials", initials);
            if (comment.Date is { } date) writer.WriteString("date", date);
            writer.WriteStartArray("targets");
            foreach (var target in comment.Targets)
            {
                writer.WriteStartObject();
                writer.WriteString("blockAnchor", target.BlockAnchor.ToString());
                writer.WriteNumber("startChar", target.StartChar);
                writer.WriteNumber("endChar", target.EndChar);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteStartArray("blocks");
            foreach (var block in comment.Blocks)
                WriteBlock(writer, block);
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    // --- blocks -----------------------------------------------------------

    private static void WriteBlock(Utf8JsonWriter writer, IrBlock block)
    {
        writer.WriteStartObject();
        writer.WriteString("anchor", block.Anchor.ToString());

        switch (block)
        {
            case IrParagraph p:
                writer.WriteString("type", "paragraph");
                WriteHashes(writer, p);
                WriteParagraphBody(writer, p);
                break;
            case IrTable t:
                writer.WriteString("type", "table");
                WriteHashes(writer, t);
                WriteTableBody(writer, t);
                break;
            case IrSectionBreak s:
                writer.WriteString("type", "sectionBreak");
                WriteHashes(writer, s);
                writer.WritePropertyName("format");
                WriteSectionFormat(writer, s.Format);
                break;
            case IrOpaqueBlock o:
                writer.WriteString("type", "opaque");
                WriteHashes(writer, o);
                writer.WriteString("element", o.ElementName.ToString());
                break;
            default:
                // The M1.1 reader never emits other block kinds; keep the writer total.
                writer.WriteString("type", "unsupported");
                WriteHashes(writer, block);
                writer.WriteString("runtimeType", block.GetType().Name);
                break;
        }

        writer.WriteEndObject();
    }

    private static void WriteHashes(Utf8JsonWriter writer, IrBlock block)
    {
        writer.WriteString("contentHash", block.ContentHash.ToHex());
        writer.WriteString("formatFingerprint", block.FormatFingerprint.ToHex());
    }

    private static void WriteParagraphBody(Utf8JsonWriter writer, IrParagraph p)
    {
        if (p.List is { } list)
        {
            writer.WriteStartObject("list");
            writer.WriteNumber("numId", list.NumId);
            if (list.AbstractNumId is { } absId)
                writer.WriteNumber("abstractNumId", absId);
            else
                writer.WriteNull("abstractNumId");
            writer.WriteNumber("ilvl", list.Ilvl);
            writer.WriteString("numberFormat", list.NumberFormat);
            if (list.StartOverride is { } start)
                writer.WriteNumber("startOverride", start);
            writer.WriteBoolean("fromStyle", list.FromStyle);
            writer.WriteEndObject();
        }

        writer.WritePropertyName("format");
        WriteParaFormat(writer, p.Format);

        if (p.InlineSectionBreakAnchor is { } sectAnchor)
            writer.WriteString("inlineSectionBreak", sectAnchor.ToString());

        writer.WriteStartArray("inlines");
        foreach (var inline in p.Inlines)
            WriteInline(writer, inline);
        writer.WriteEndArray();
    }

    private static void WriteTableBody(Utf8JsonWriter writer, IrTable t)
    {
        writer.WriteString("unmodeledTablePropsDigest", t.UnmodeledTablePropsDigest.ToHex());
        writer.WriteStartArray("rows");
        foreach (var row in t.Rows)
        {
            writer.WriteStartObject();
            writer.WriteString("anchor", row.Anchor.ToString());
            writer.WriteString("contentHash", row.ContentHash.ToHex());
            writer.WriteStartArray("cells");
            foreach (var cell in row.Cells)
            {
                writer.WriteStartObject();
                writer.WriteString("anchor", cell.Anchor.ToString());
                writer.WriteNumber("gridSpan", cell.GridSpan);
                writer.WriteString("vMerge", cell.VMerge.ToString());
                writer.WriteString("contentHash", cell.ContentHash.ToHex());
                writer.WriteStartArray("blocks");
                foreach (var child in cell.Blocks)
                    WriteBlock(writer, child);
                writer.WriteEndArray();
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
    }

    // --- inlines ----------------------------------------------------------

    private static void WriteInline(Utf8JsonWriter writer, IrInline inline)
    {
        writer.WriteStartObject();
        switch (inline)
        {
            case IrTextRun r:
                writer.WriteString("kind", "text");
                writer.WriteString("text", r.Text);
                writer.WritePropertyName("format");
                WriteRunFormat(writer, r.Format);
                break;
            case IrTab:
                writer.WriteString("kind", "tab");
                break;
            case IrBreak b:
                writer.WriteString("kind", "break");
                writer.WriteString("breakKind", b.Kind.ToString());
                break;
            case IrFieldRun fld:
                writer.WriteString("kind", "field");
                writer.WriteString("instruction", fld.Instruction);
                writer.WriteStartArray("cachedResult");
                foreach (var child in fld.CachedResult)
                    WriteInline(writer, child);
                writer.WriteEndArray();
                break;
            case IrHyperlink h:
                writer.WriteString("kind", "hyperlink");
                if (h.Target is { } target) writer.WriteString("target", target);
                writer.WriteStartArray("inlines");
                foreach (var child in h.Inlines)
                    WriteInline(writer, child);
                writer.WriteEndArray();
                break;
            case IrNoteRef n:
                writer.WriteString("kind", "noteRef");
                writer.WriteString("noteKind", n.Kind.ToString());
                writer.WriteString("noteId", n.NoteId);
                break;
            case IrInlineImage img:
                writer.WriteString("kind", "image");
                // PartUri is a relative part URI (e.g. "/word/media/image1.png"); no absolute
                // filesystem path may leak into the diagnostic output.
                writer.WriteString("partUri", img.PartUri.ToString());
                writer.WriteString("imageBytesHash", img.ImageBytesHash.ToHex());
                writer.WriteNumber("widthEmu", img.WidthEmu);
                writer.WriteNumber("heightEmu", img.HeightEmu);
                if (img.AltText is { } altText) writer.WriteString("altText", altText);
                if (img.Unid is { } imgUnid) writer.WriteString("unid", imgUnid);
                break;
            case IrOpaqueInline o:
                writer.WriteString("kind", "opaque");
                writer.WriteString("element", o.ElementName.ToString());
                writer.WriteString("hash", o.CanonicalHash.ToHex());
                break;
            default:
                // No concrete IrInline subtype should reach here; the completeness guard test
                // (IrDiagnosticJsonTests) enforces a branch above for every concrete kind.
                writer.WriteString("kind", "unsupported");
                writer.WriteString("type", inline.GetType().Name);
                break;
        }
        writer.WriteEndObject();
    }

    // --- formats ----------------------------------------------------------

    private static void WriteRunFormat(Utf8JsonWriter writer, IrRunFormat f)
    {
        writer.WriteStartObject();
        if (f.StyleId is { } styleId) writer.WriteString("styleId", styleId);
        if (f.Bold is { } bold) writer.WriteBoolean("bold", bold);
        if (f.Italic is { } italic) writer.WriteBoolean("italic", italic);
        if (f.Underline is { } underline)
        {
            writer.WriteStartObject("underline");
            writer.WriteString("kind", underline.Kind.ToString());
            if (underline.ColorHex is { } colorHex) writer.WriteString("colorHex", colorHex);
            writer.WriteEndObject();
        }
        if (f.Strike is { } strike) writer.WriteBoolean("strike", strike);
        if (f.DoubleStrike is { } dstrike) writer.WriteBoolean("doubleStrike", dstrike);
        if (f.VertAlign is { } vertAlign) writer.WriteString("vertAlign", vertAlign.ToString());
        if (f.FontAscii is { } fontAscii) writer.WriteString("fontAscii", fontAscii);
        if (f.SizeHalfPoints is { } size) writer.WriteNumber("sizeHalfPoints", size);
        if (f.ColorHex is { } colorHex2) writer.WriteString("colorHex", colorHex2);
        if (f.Highlight is { } highlight) writer.WriteString("highlight", highlight);
        if (f.Caps is { } caps) writer.WriteBoolean("caps", caps);
        if (f.SmallCaps is { } smallCaps) writer.WriteBoolean("smallCaps", smallCaps);
        if (f.Vanish is { } vanish) writer.WriteBoolean("vanish", vanish);
        writer.WriteEndObject();
    }

    private static void WriteParaFormat(Utf8JsonWriter writer, IrParaFormat f)
    {
        writer.WriteStartObject();
        if (f.StyleId is { } styleId) writer.WriteString("styleId", styleId);
        if (f.Justification is { } justification) writer.WriteString("justification", justification.ToString());
        if (f.IndentLeftTwips is { } left) writer.WriteNumber("indentLeftTwips", left);
        if (f.IndentRightTwips is { } right) writer.WriteNumber("indentRightTwips", right);
        if (f.IndentFirstLineTwips is { } first) writer.WriteNumber("indentFirstLineTwips", first);
        if (f.SpacingBeforeTwips is { } before) writer.WriteNumber("spacingBeforeTwips", before);
        if (f.SpacingAfterTwips is { } after) writer.WriteNumber("spacingAfterTwips", after);
        if (f.LineSpacing is { } lineSpacing)
        {
            writer.WriteStartObject("lineSpacing");
            writer.WriteNumber("valueTwips", lineSpacing.ValueTwips);
            writer.WriteString("rule", lineSpacing.Rule.ToString());
            writer.WriteEndObject();
        }
        if (f.OutlineLevel is { } outline) writer.WriteNumber("outlineLevel", outline);
        if (f.KeepNext is { } keepNext) writer.WriteBoolean("keepNext", keepNext);
        if (f.KeepLines is { } keepLines) writer.WriteBoolean("keepLines", keepLines);
        if (f.PageBreakBefore is { } pageBreak) writer.WriteBoolean("pageBreakBefore", pageBreak);
        writer.WriteEndObject();
    }

    private static void WriteSectionFormat(Utf8JsonWriter writer, IrSectionFormat f)
    {
        writer.WriteStartObject();
        if (f.PageWidthTwips is { } w) writer.WriteNumber("pageWidthTwips", w);
        if (f.PageHeightTwips is { } h) writer.WriteNumber("pageHeightTwips", h);
        if (f.Landscape is { } landscape) writer.WriteBoolean("landscape", landscape);
        if (f.MarginTopTwips is { } top) writer.WriteNumber("marginTopTwips", top);
        if (f.MarginBottomTwips is { } bottom) writer.WriteNumber("marginBottomTwips", bottom);
        if (f.MarginLeftTwips is { } mleft) writer.WriteNumber("marginLeftTwips", mleft);
        if (f.MarginRightTwips is { } mright) writer.WriteNumber("marginRightTwips", mright);
        if (f.SectionType is { } sectionType) writer.WriteString("sectionType", sectionType);
        writer.WriteEndObject();
    }
}
