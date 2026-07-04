#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Docxodus.Internal;

internal enum ParserBlockKind
{
    Paragraph,
    Heading1, Heading2, Heading3, Heading4, Heading5, Heading6,
    Quote,
    Code,
    BulletItem,
    OrderedItem,
}

internal sealed record ParsedBlock(
    ParserBlockKind Kind,
    int ListLevel,
    IReadOnlyList<XElement> RunElements);

internal sealed record ParseError(EditErrorCode Code, string Message);

internal sealed class ParseResult
{
    public IReadOnlyList<ParsedBlock> Blocks { get; init; } = Array.Empty<ParsedBlock>();
    public ParseError? Error { get; init; }
    public bool Success => Error is null;

    public static ParseResult Ok(IReadOnlyList<ParsedBlock> blocks) => new() { Blocks = blocks };
    public static ParseResult Fail(EditErrorCode code, string msg) =>
        new() { Error = new ParseError(code, msg) };
}

/// <summary>
/// Hand-rolled parser for the projector-symmetric markdown subset accepted by
/// <see cref="DocxSession"/>. Block-level: paragraphs, ATX headings, bulleted +
/// ordered lists (with indent-based nesting), blockquotes, fenced code blocks.
/// Inline: bold, italic, code, strike (GFM), hyperlinks, soft breaks, backslash
/// escapes. Anything outside the subset is rejected with a typed
/// <see cref="EditErrorCode"/> so the caller can pattern-match on remediation.
/// </summary>
internal static class MarkdownPayloadParser
{
    /// <summary>
    /// Special XName used to stash a hyperlink's URL on the parsed
    /// <c>&lt;w:hyperlink&gt;</c> element. <see cref="DocxSession"/> promotes
    /// this to a real relationship before inserting the run into the document.
    /// </summary>
    internal static readonly XName HrefAttr = XName.Get("href", "docxodus:");

    public static ParseResult Parse(string markdown)
    {
        if (markdown is null) return ParseResult.Fail(EditErrorCode.MalformedMarkdown, "null payload");
        try
        {
            var rawBlocks = SplitBlocks(markdown);
            var parsed = new List<ParsedBlock>(rawBlocks.Count);
            foreach (var raw in rawBlocks)
            {
                var b = ParseBlock(raw);
                if (b is not null) parsed.Add(b);
            }
            return ParseResult.Ok(parsed);
        }
        catch (MarkdownPayloadException ex)
        {
            return ParseResult.Fail(ex.Code, ex.Message);
        }
    }

    // ─── Block splitter ─────────────────────────────────────────────────

    private static List<string> SplitBlocks(string md)
    {
        var normalized = md.Replace("\r\n", "\n");
        var lines = normalized.Split('\n');
        var result = new List<string>();
        var buf = new StringBuilder();
        bool inFence = false;

        void Flush()
        {
            if (buf.Length > 0) { result.Add(buf.ToString()); buf.Clear(); }
        }

        foreach (var line in lines)
        {
            if (line.StartsWith("```", StringComparison.Ordinal))
            {
                if (!inFence) Flush();
                inFence = !inFence;
                if (buf.Length > 0) buf.Append('\n');
                buf.Append(line);
                if (!inFence) Flush();
                continue;
            }
            if (!inFence && line.Length == 0) { Flush(); continue; }
            if (!inFence && IsListLine(line)) { Flush(); buf.Append(line); Flush(); continue; }
            if (buf.Length > 0) buf.Append('\n');
            buf.Append(line);
        }
        Flush();
        return result;
    }

    private static bool IsListLine(string line)
    {
        int i = 0;
        while (i < line.Length && line[i] == ' ') i++;
        if (i >= line.Length) return false;
        if (line[i] == '-' || line[i] == '*' || line[i] == '+')
            return i + 1 < line.Length && line[i + 1] == ' ';
        int start = i;
        while (i < line.Length && char.IsDigit(line[i])) i++;
        return i > start && i + 1 < line.Length && line[i] == '.' && line[i + 1] == ' ';
    }

    // ─── Block parser ───────────────────────────────────────────────────

    private static ParsedBlock? ParseBlock(string raw)
    {
        if (string.IsNullOrEmpty(raw)) return null;

        // Pipe table detection (early reject)
        if (raw.StartsWith("|", StringComparison.Ordinal)
            && raw.Contains("\n|", StringComparison.Ordinal)
            && raw.Contains("---", StringComparison.Ordinal))
        {
            throw new MarkdownPayloadException(
                EditErrorCode.TableInsertNotSupported,
                "Tables can't be inserted via markdown in v1. Use ReplaceCellContent(anchor, md) to edit an existing cell. InsertTable is planned for v2.");
        }

        // Fenced code
        if (raw.StartsWith("```", StringComparison.Ordinal))
        {
            var firstNl = raw.IndexOf('\n');
            var lastNl = raw.LastIndexOf('\n');
            string inner = firstNl >= 0 && lastNl > firstNl
                ? raw.Substring(firstNl + 1, lastNl - firstNl - 1)
                : "";
            return new ParsedBlock(ParserBlockKind.Code, 0,
                new[] { TextRun(inner, bold: false, italic: false) });
        }

        // ATX headings
        if (raw.StartsWith("#", StringComparison.Ordinal))
        {
            int level = 0;
            while (level < raw.Length && raw[level] == '#') level++;
            if (level >= 1 && level <= 6 && level < raw.Length && raw[level] == ' ')
            {
                var headingText = raw.Substring(level + 1).TrimEnd();
                return new ParsedBlock(
                    (ParserBlockKind)((int)ParserBlockKind.Heading1 + level - 1),
                    0,
                    ParseInline(headingText));
            }
        }

        // Blockquote
        if (raw.StartsWith("> ", StringComparison.Ordinal))
        {
            var quoteText = string.Join("\n",
                raw.Split('\n').Select(l => l.StartsWith("> ", StringComparison.Ordinal) ? l.Substring(2) : l));
            return new ParsedBlock(ParserBlockKind.Quote, 0, ParseInline(quoteText));
        }

        // Lists
        if (IsListLine(raw))
        {
            int indent = 0;
            while (indent < raw.Length && raw[indent] == ' ') indent++;
            int level = indent / 2;
            bool bullet = raw[indent] == '-' || raw[indent] == '*' || raw[indent] == '+';
            int markerEnd;
            if (bullet)
            {
                markerEnd = indent + 2;
            }
            else
            {
                markerEnd = indent;
                while (markerEnd < raw.Length && char.IsDigit(raw[markerEnd])) markerEnd++;
                markerEnd += 2;
            }
            var itemText = raw.Substring(markerEnd).TrimEnd();
            return new ParsedBlock(
                bullet ? ParserBlockKind.BulletItem : ParserBlockKind.OrderedItem,
                level,
                ParseInline(itemText));
        }

        return new ParsedBlock(ParserBlockKind.Paragraph, 0, ParseInline(raw));
    }

    // ─── Inline parser ──────────────────────────────────────────────────

    private static IReadOnlyList<XElement> ParseInline(string text)
    {
        var list = new List<XElement>();
        var sb = new StringBuilder();
        bool bold = false, italic = false;

        void FlushText()
        {
            if (sb.Length == 0) return;
            list.Add(TextRun(sb.ToString(), bold, italic));
            sb.Clear();
        }

        for (int i = 0; i < text.Length; i++)
        {
            var c = text[i];

            if (c == '\\' && i + 1 < text.Length)
            {
                sb.Append(text[++i]);
                continue;
            }

            // Hard line break: a newline WITHIN a paragraph becomes a real w:br (Word's
            // intra-paragraph line break), not a literal '\n' in w:t (which Word renders
            // as a space). GFM's trailing-two-spaces hard-break marker is consumed. This
            // is the write-side mirror of WmlToHtmlConverter's read-side "w:br -> '  \n'".
            if (c == '\n')
            {
                while (sb.Length > 0 && sb[sb.Length - 1] == ' ') sb.Length--;
                FlushText();
                // w:br must live inside a w:r to be valid OOXML and to be seen by the
                // projector's run-scoped AppendRunText.
                list.Add(new XElement(W.r, new XElement(W.br)));
                continue;
            }

            // Code span
            if (c == '`')
            {
                FlushText();
                int end = text.IndexOf('`', i + 1);
                if (end < 0) { sb.Append(c); continue; }
                list.Add(TextRun(text.Substring(i + 1, end - i - 1), bold: false, italic: false, code: true));
                i = end;
                continue;
            }

            // Strike (~~text~~)
            if (c == '~' && i + 1 < text.Length && text[i + 1] == '~')
            {
                FlushText();
                int end = text.IndexOf("~~", i + 2, StringComparison.Ordinal);
                if (end < 0) { sb.Append("~~"); i++; continue; }
                list.Add(TextRun(text.Substring(i + 2, end - i - 2), bold: false, italic: false, strike: true));
                i = end + 1;
                continue;
            }

            // Bold (**text**)
            if (c == '*' && i + 1 < text.Length && text[i + 1] == '*')
            {
                FlushText();
                bold = !bold;
                i++;
                continue;
            }

            // Italic (*text*)
            if (c == '*')
            {
                FlushText();
                italic = !italic;
                continue;
            }

            // Image: ![alt](url) — rejected
            if (c == '!' && i + 1 < text.Length && text[i + 1] == '[')
            {
                throw new MarkdownPayloadException(
                    EditErrorCode.ImageInsertNotSupported,
                    "Image insertion requires a binary upload. AddImage(anchor, bytes, alt) is planned for v2.");
            }

            // Footnote ref or link
            if (c == '[')
            {
                if (i + 1 < text.Length && text[i + 1] == '^')
                {
                    int close = text.IndexOf(']', i + 2);
                    if (close > 0)
                    {
                        throw new MarkdownPayloadException(
                            EditErrorCode.FootnoteRefNotSupported,
                            "Footnote/endnote references are output-only in v1. AddFootnote(anchor, md) is planned for v2.");
                    }
                }
                int rb = text.IndexOf(']', i + 1);
                if (rb > 0 && rb + 1 < text.Length && text[rb + 1] == '(')
                {
                    int rp = text.IndexOf(')', rb + 2);
                    if (rp > 0)
                    {
                        FlushText();
                        var linkText = text.Substring(i + 1, rb - i - 1);
                        var url = text.Substring(rb + 2, rp - rb - 2);
                        var hyperlink = new XElement(W.hyperlink,
                            new XAttribute(HrefAttr, url),
                            TextRun(linkText, bold, italic));
                        list.Add(hyperlink);
                        i = rp;
                        continue;
                    }
                }
                sb.Append(c);
                continue;
            }

            // Anchor token: {#cmt:...} → CommentMarker; {#kind:scope:unid} → AnchorToken
            if (c == '{' && i + 1 < text.Length && text[i + 1] == '#')
            {
                int close = text.IndexOf('}', i + 2);
                if (close > 0)
                {
                    var token = text.Substring(i + 2, close - i - 2);
                    if (token.StartsWith("cmt:", StringComparison.Ordinal))
                    {
                        throw new MarkdownPayloadException(
                            EditErrorCode.CommentMarkerNotSupported,
                            "Comment markers are output-only in v1. AddComment(anchor, author, md) is planned for v2.");
                    }
                    throw new MarkdownPayloadException(
                        EditErrorCode.AnchorTokenInPayload,
                        "Anchor tokens like {#kind:scope:unid} are projection output, not input. Remove them from the payload.");
                }
            }

            sb.Append(c);
        }

        FlushText();
        return list;
    }

    // ─── Run builder ────────────────────────────────────────────────────

    internal static XElement TextRun(string text, bool bold, bool italic,
        bool code = false, bool strike = false, bool underline = false,
        string? color = null, string? runStyle = null)
    {
        var rPr = new XElement(W.rPr);
        if (bold) rPr.Add(new XElement(W.b));
        if (italic) rPr.Add(new XElement(W.i));
        if (strike) rPr.Add(new XElement(W.strike));
        if (underline) rPr.Add(new XElement(W.u, new XAttribute(W.val, "single")));
        if (code || runStyle is not null)
            rPr.Add(new XElement(W.rStyle, new XAttribute(W.val, runStyle ?? "Code")));
        if (!string.IsNullOrEmpty(color))
            rPr.Add(new XElement(W.color, new XAttribute(W.val, color)));

        var run = new XElement(W.r);
        if (rPr.HasElements) run.Add(rPr);
        run.Add(new XElement(W.t,
            new XAttribute(XNamespace.Xml + "space", "preserve"),
            text));
        return run;
    }

    internal sealed class MarkdownPayloadException : Exception
    {
        public EditErrorCode Code { get; }
        public MarkdownPayloadException(EditErrorCode code, string msg) : base(msg) => Code = code;
    }
}
