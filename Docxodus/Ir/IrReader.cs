#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Ir;

/// <summary>
/// Reads an OOXML word-processing document into the Document IR for the body scope (spec §5,
/// M1.1 subset). The reader is <em>total</em>: any body child it does not model is preserved as an
/// <see cref="IrOpaqueBlock"/> (or <see cref="IrOpaqueInline"/> at run level), so it never throws
/// on weird-but-valid OOXML. It never mutates the caller's document — it works over a private copy.
/// </summary>
/// <remarks>
/// Pipeline: copy the caller's bytes → normalize tracked revisions per
/// <see cref="IrReaderOptions.RevisionView"/> → open the copy → assign deterministic Unids
/// (same call <c>WmlToMarkdownConverter</c> makes) → walk <c>w:body</c> children in document order.
/// In M1.1 only <see cref="IrScopes.Body"/> is honored; headers/footers, notes, and comments are
/// emitted as empty stores.
/// </remarks>
internal static class IrReader
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // The empty-unmodeled-container digest: CanonicalHash of <unmodeled/> with no children.
    // Cached because it is the fingerprint of every format record that carries no leftover props.
    private static readonly IrHash EmptyUnmodeledDigest =
        IrHasher.CanonicalHash(new XElement("unmodeled"));

    // Constant consumed-name sets for the unmodeled-digest computation, hoisted so each
    // paragraph/run/section read does not reallocate them.
    private static readonly HashSet<XName> PPrConsumed = new()
    {
        W + "pStyle", W + "jc", W + "ind", W + "spacing", W + "outlineLvl",
        W + "keepNext", W + "keepLines", W + "pageBreakBefore",
    };

    // The always-consumed rPr children. w:vertAlign is consumed conditionally (only when it maps
    // to a modeled sub/superscript); MapRunFormat handles that case without per-run allocation.
    private static readonly HashSet<XName> RPrConsumed = new()
    {
        W + "rStyle", W + "b", W + "i", W + "strike", W + "dstrike", W + "caps",
        W + "smallCaps", W + "vanish", W + "u", W + "sz", W + "color", W + "highlight",
    };

    private static readonly HashSet<XName> SectPrConsumed = new()
    {
        W + "pgSz", W + "pgMar", W + "type",
    };

    /// <summary>
    /// Read <paramref name="doc"/> into an <see cref="IrDocument"/>. The caller's
    /// <see cref="WmlDocument.DocumentByteArray"/> is left byte-for-byte unchanged.
    /// </summary>
    public static IrDocument Read(WmlDocument doc, IrReaderOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(doc);
        options ??= new IrReaderOptions();

        // 1. Work over a private copy so the caller's bytes are never mutated.
        var working = new WmlDocument(doc);

        // 2. Normalize tracked revisions (rule N13).
        working = ApplyRevisionView(working, options.RevisionView);

        // 3. Open the copy, assign deterministic Unids, and walk the body.
        using var stream = new OpenXmlMemoryStreamDocument(working);
        using var wdoc = stream.GetWordprocessingDocument();

        var main = wdoc.MainDocumentPart
            ?? throw new DocxodusException("Document has no MainDocumentPart.");
        var mainXDoc = main.GetXDocument();
        var root = mainXDoc.Root
            ?? throw new DocxodusException("MainDocumentPart has no root element.");
        UnidHelper.AssignToAllElementsDeterministic(root);

        // Stash the owning part on the root so WmlToMarkdownConverter.KindFor → IsListItem can
        // reach the StyleDefinitionsPart and walk the pStyle → basedOn chain. Without it, a
        // paragraph that is a list item only via style inheritance (no inline w:numPr) classifies
        // as `p` instead of `li`, breaking anchor-kind parity with the markdown projection (which
        // stashes the same annotation in BuildAnchorIndex).
        if (root.Annotation<OpenXmlPart>() == null)
            root.AddAnnotation(main);

        var partUri = main.Uri;
        var ctx = new ReadContext(partUri);

        var body = root.Element(W + "body")
            ?? throw new DocxodusException("Document has no w:body element.");

        var blocks = new List<IrBlock>();
        foreach (var child in body.Elements())
            blocks.Add(BuildBlock(child, ctx));

        // 5. Anchor index over blocks only (rows/cells are positional, not blocks).
        var anchorIndex = new Dictionary<string, IrBlock>(StringComparer.Ordinal);
        foreach (var b in blocks)
            IndexBlock(b, anchorIndex);

        var sources = new Dictionary<Uri, XDocument> { [partUri] = mainXDoc };

        return new IrDocument
        {
            Body = new IrScope("body", IrNodeList.From(blocks)),
            Footnotes = IrNoteStore.Empty,
            Endnotes = IrNoteStore.Empty,
            Comments = IrCommentStore.Empty,
            Styles = IrStyleRegistry.Empty,
            Numbering = IrNumberingRegistry.Empty,
            ThemeFonts = IrThemeFonts.Empty,
            AnchorIndex = anchorIndex,
            Sources = sources,
        };
    }

    /// <summary>Carries the part URI through the recursive walk for provenance.</summary>
    private sealed class ReadContext
    {
        public ReadContext(Uri partUri) => PartUri = partUri;

        public Uri PartUri { get; }

        public IrProvenance Provenance(XElement element) =>
            new() { Element = element, PartUri = PartUri };
    }

    // --- revisions --------------------------------------------------------

    private static WmlDocument ApplyRevisionView(WmlDocument working, RevisionView view)
    {
        switch (view)
        {
            case RevisionView.FailIfPresent:
                if (HasRevisionMarkup(working))
                    throw new DocxodusException(
                        "Document contains tracked revisions and RevisionView is FailIfPresent.");
                return working;
            case RevisionView.Accept:
                return RevisionProcessor.AcceptRevisions(working);
            case RevisionView.Reject:
                return RevisionProcessor.RejectRevisions(working);
            default:
                return working;
        }
    }

    private static readonly XName[] RevisionElementNames =
    {
        W + "ins", W + "del", W + "moveFrom", W + "moveTo", W + "rPrChange", W + "pPrChange",
    };

    private static bool HasRevisionMarkup(WmlDocument working)
    {
        using var stream = new OpenXmlMemoryStreamDocument(working);
        using var wdoc = stream.GetWordprocessingDocument();
        var root = wdoc.MainDocumentPart?.GetXDocument().Root;
        if (root is null)
            return false;
        var names = new HashSet<XName>(RevisionElementNames);
        return root.DescendantsAndSelf().Any(e => names.Contains(e.Name));
    }

    // --- block dispatch ---------------------------------------------------

    private static IrBlock BuildBlock(XElement el, ReadContext ctx)
    {
        if (el.Name == W + "p")
            return BuildParagraph(el, ctx);
        if (el.Name == W + "tbl")
            return BuildTable(el, ctx);
        if (el.Name == W + "sectPr")
            return BuildSectionBreak(el, ctx);
        return BuildOpaqueBlock(el, ctx);
    }

    private static string Unid(XElement el) => (string?)el.Attribute(PtOpenXml.Unid) ?? "";

    private static IrAnchor AnchorFor(IrAnchorKind kind, XElement el) =>
        new(kind, "body", Unid(el));

    // --- paragraph --------------------------------------------------------

    private static IrParagraph BuildParagraph(XElement p, ReadContext ctx)
    {
        var kindToken = WmlToMarkdownConverter.KindFor(p);
        var kind = kindToken is null ? IrAnchorKind.P : IrAnchor.KindFromToken(kindToken);

        var pPr = p.Element(W + "pPr");
        var (paraFormat, listInfo) = MapParaFormat(pPr);

        var inlines = new List<IrInline>();
        foreach (var child in p.Elements())
        {
            if (child.Name == W + "pPr")
                continue;
            if (child.Name == W + "r")
                AppendRun(child, inlines);
            else if (child.Name == W + "proofErr")
                continue; // rule N2: pure noise, never emit.
            else
                inlines.Add(new IrOpaqueInline(child.Name, IrHasher.CanonicalHash(child)));
        }

        // Post-process: drop empty text runs (N10), then coalesce adjacent equal-format runs (N5).
        var processed = CoalesceRuns(DropEmptyTextRuns(inlines));

        var contentHash = ComputeParagraphContentHash(processed);
        var formatFingerprint = IrHasher.FingerprintBlock(paraFormat, RunFormatsInOrder(processed));

        return new IrParagraph
        {
            Anchor = AnchorFor(kind, p),
            Format = paraFormat,
            List = listInfo,
            Inlines = IrNodeList.From(processed),
            ContentHash = contentHash,
            FormatFingerprint = formatFingerprint,
            Source = ctx.Provenance(p),
        };
    }

    private static void AppendRun(XElement r, List<IrInline> inlines)
    {
        var rPr = r.Element(W + "rPr");
        var runFormat = MapRunFormat(rPr);

        foreach (var child in r.Elements())
        {
            if (child.Name == W + "rPr")
                continue;
            if (child.Name == W + "t")
                inlines.Add(new IrTextRun(child.Value, runFormat));
            else if (child.Name == W + "tab")
                inlines.Add(new IrTab(runFormat));
            else if (child.Name == W + "br")
                inlines.Add(new IrBreak(BreakKind(child)));
            else
                inlines.Add(new IrOpaqueInline(child.Name, IrHasher.CanonicalHash(child)));
        }
    }

    private static IrBreakKind BreakKind(XElement br)
    {
        var type = (string?)br.Attribute(W + "type");
        return type switch
        {
            "page" => IrBreakKind.Page,
            "column" => IrBreakKind.Column,
            _ => IrBreakKind.Line, // null, "textWrapping", or anything else → line.
        };
    }

    private static List<IrInline> DropEmptyTextRuns(List<IrInline> inlines) =>
        inlines.Where(i => i is not IrTextRun { Text: "" }).ToList();

    private static List<IrInline> CoalesceRuns(List<IrInline> inlines)
    {
        var result = new List<IrInline>(inlines.Count);
        foreach (var inline in inlines)
        {
            if (inline is IrTextRun run
                && result.Count > 0
                && result[^1] is IrTextRun prev
                && prev.Format.Equals(run.Format))
            {
                result[^1] = prev with { Text = prev.Text + run.Text };
            }
            else
            {
                result.Add(inline);
            }
        }
        return result;
    }

    /// <summary>The run format carried by each inline that has one, in inline order.</summary>
    private static IEnumerable<IrRunFormat> RunFormatsInOrder(IEnumerable<IrInline> inlines)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrTextRun r: yield return r.Format; break;
                case IrTab t: yield return t.Format; break;
            }
        }
    }

    private static IrHash ComputeParagraphContentHash(IEnumerable<IrInline> inlines)
    {
        var builder = new IrContentHashBuilder();
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrTextRun r:
                    builder.AppendText(r.Text);
                    break;
                case IrTab:
                    builder.AppendSentinel(IrContentHashBuilder.SentinelTab);
                    break;
                case IrBreak b:
                    builder.AppendSentinel(b.Kind switch
                    {
                        IrBreakKind.Page => IrContentHashBuilder.SentinelPageBreak,
                        IrBreakKind.Column => IrContentHashBuilder.SentinelColumnBreak,
                        _ => IrContentHashBuilder.SentinelLineBreak,
                    });
                    break;
                case IrOpaqueInline o:
                    builder.AppendSentinel(IrContentHashBuilder.SentinelOpaque);
                    builder.AppendHash(o.CanonicalHash);
                    break;
            }
        }
        return builder.Build();
    }

    // --- paragraph format -------------------------------------------------

    private static (IrParaFormat Format, IrListInfo? List) MapParaFormat(XElement? pPr)
    {
        if (pPr is null)
            return (new IrParaFormat { UnmodeledDigest = EmptyUnmodeledDigest }, null);

        string? styleId = AttrVal(pPr.Element(W + "pStyle"));

        IrJustification? justification = null;
        var jcVal = AttrVal(pPr.Element(W + "jc"));
        if (jcVal is not null)
            justification = jcVal switch
            {
                "left" or "start" => IrJustification.Left,
                "center" => IrJustification.Center,
                "right" or "end" => IrJustification.Right,
                "both" => IrJustification.Both,
                "distribute" => IrJustification.Distribute,
                _ => IrJustification.Other,
            };

        var ind = pPr.Element(W + "ind");
        int? indentLeft = IntAttr(ind, W + "left") ?? IntAttr(ind, W + "start");
        int? indentRight = IntAttr(ind, W + "right") ?? IntAttr(ind, W + "end");
        int? indentFirst = IntAttr(ind, W + "firstLine");
        var hanging = IntAttr(ind, W + "hanging");
        if (hanging is not null)
            indentFirst = -hanging.Value;

        var spacing = pPr.Element(W + "spacing");
        int? spacingBefore = IntAttr(spacing, W + "before");
        int? spacingAfter = IntAttr(spacing, W + "after");

        IrLineSpacing? lineSpacing = null;
        var lineVal = IntAttr(spacing, W + "line");
        if (lineVal is not null)
        {
            var rule = (string?)spacing?.Attribute(W + "lineRule") switch
            {
                "atLeast" => IrLineSpacingRule.AtLeast,
                "exact" => IrLineSpacingRule.Exact,
                _ => IrLineSpacingRule.Auto,
            };
            lineSpacing = new IrLineSpacing(lineVal.Value, rule);
        }

        int? outlineLevel = IntAttr(pPr.Element(W + "outlineLvl"), W + "val");
        bool? keepNext = Toggle(pPr.Element(W + "keepNext"));
        bool? keepLines = Toggle(pPr.Element(W + "keepLines"));
        bool? pageBreakBefore = Toggle(pPr.Element(W + "pageBreakBefore"));

        IrListInfo? listInfo = null;
        var numPr = pPr.Element(W + "numPr");
        if (numPr is not null)
        {
            var numId = IntAttr(numPr.Element(W + "numId"), W + "val");
            var ilvl = IntAttr(numPr.Element(W + "ilvl"), W + "val");
            if (numId is not null)
                listInfo = new IrListInfo(numId.Value, null, ilvl ?? 0, "", null, false);
        }

        // Unmodeled leftovers: every pPr child not consumed by a modeled field above.
        // numPr is consumed for list facts but ALSO kept here so numbering still affects the
        // fingerprint until M1.3 resolution. w:rPr (mark props) and mid-doc w:sectPr stay too.
        var digest = UnmodeledDigest(pPr, PPrConsumed);

        var format = new IrParaFormat
        {
            StyleId = styleId,
            Justification = justification,
            IndentLeftTwips = indentLeft,
            IndentRightTwips = indentRight,
            IndentFirstLineTwips = indentFirst,
            SpacingBeforeTwips = spacingBefore,
            SpacingAfterTwips = spacingAfter,
            LineSpacing = lineSpacing,
            OutlineLevel = outlineLevel,
            KeepNext = keepNext,
            KeepLines = keepLines,
            PageBreakBefore = pageBreakBefore,
            UnmodeledDigest = digest,
        };
        return (format, listInfo);
    }

    // --- run format -------------------------------------------------------

    private static IrRunFormat MapRunFormat(XElement? rPr)
    {
        if (rPr is null)
            return new IrRunFormat { UnmodeledDigest = EmptyUnmodeledDigest };

        string? styleId = AttrVal(rPr.Element(W + "rStyle"));
        bool? bold = Toggle(rPr.Element(W + "b"));
        bool? italic = Toggle(rPr.Element(W + "i"));
        bool? strike = Toggle(rPr.Element(W + "strike"));
        bool? doubleStrike = Toggle(rPr.Element(W + "dstrike"));
        bool? caps = Toggle(rPr.Element(W + "caps"));
        bool? smallCaps = Toggle(rPr.Element(W + "smallCaps"));
        bool? vanish = Toggle(rPr.Element(W + "vanish"));

        IrUnderline? underline = MapUnderline(rPr.Element(W + "u"));

        // baseline vertAlign is left null and folded into the unmodeled digest below.
        IrVertAlign? vertAlign = null;
        var vertVal = AttrVal(rPr.Element(W + "vertAlign"));
        if (vertVal == "subscript")
            vertAlign = IrVertAlign.Subscript;
        else if (vertVal == "superscript")
            vertAlign = IrVertAlign.Superscript;

        string? fontAscii = (string?)rPr.Element(W + "rFonts")?.Attribute(W + "ascii");
        int? size = IntAttr(rPr.Element(W + "sz"), W + "val");
        string? colorHex = AttrVal(rPr.Element(W + "color"));
        string? highlight = AttrVal(rPr.Element(W + "highlight"));

        // Consumed rPr children come from the static RPrConsumed set. w:rFonts is only partially
        // consumed (ascii); keep it in the unmodeled digest so its other faces (hAnsi/cs/eastAsia)
        // still affect the fingerprint. w:vertAlign is consumed only when it maps to a modeled
        // sub/superscript; vertAlign="baseline" stays unmodeled. Pass it as a conditional extra so
        // no per-run set is allocated.
        var digest = UnmodeledDigest(rPr, RPrConsumed, vertAlign is not null ? W + "vertAlign" : null);

        return new IrRunFormat
        {
            StyleId = styleId,
            Bold = bold,
            Italic = italic,
            Underline = underline,
            Strike = strike,
            DoubleStrike = doubleStrike,
            VertAlign = vertAlign,
            FontAscii = fontAscii,
            SizeHalfPoints = size,
            ColorHex = colorHex,
            Highlight = highlight,
            Caps = caps,
            SmallCaps = smallCaps,
            Vanish = vanish,
            UnmodeledDigest = digest,
        };
    }

    private static IrUnderline? MapUnderline(XElement? u)
    {
        if (u is null)
            return null;
        var val = (string?)u.Attribute(W + "val");
        var kind = val switch
        {
            "single" => IrUnderlineKind.Single,
            "double" => IrUnderlineKind.Double,
            "thick" => IrUnderlineKind.Thick,
            "dotted" => IrUnderlineKind.Dotted,
            "dash" or "dashed" => IrUnderlineKind.Dashed,
            "wave" => IrUnderlineKind.Wave,
            "words" => IrUnderlineKind.Words,
            "none" => IrUnderlineKind.None,
            _ => IrUnderlineKind.Other,
        };
        var color = (string?)u.Attribute(W + "color");
        return new IrUnderline(kind, color);
    }

    // --- table ------------------------------------------------------------

    private static IrTable BuildTable(XElement tbl, ReadContext ctx)
    {
        var rows = new List<IrRow>();
        var rowHashes = new List<IrHash>();
        var cellFingerprints = new List<IrHash>();

        foreach (var tr in tbl.Elements(W + "tr"))
        {
            var (row, cellFingerprintsForRow) = BuildRow(tr, ctx);
            rows.Add(row);
            rowHashes.Add(row.ContentHash);
            cellFingerprints.AddRange(cellFingerprintsForRow);
        }

        // Non-tr children of the table (tblPr, tblGrid) + non-tc children of each tr (trPr)
        // fold into one unmodeled container so any table-level prop change flips the fingerprint.
        var unmodeledContainer = new XElement("unmodeled");
        foreach (var child in tbl.Elements().Where(e => e.Name != W + "tr"))
            unmodeledContainer.Add(new XElement(child));
        foreach (var tr in tbl.Elements(W + "tr"))
            foreach (var child in tr.Elements().Where(e => e.Name != W + "tc"))
                unmodeledContainer.Add(new XElement(child));
        var tablePropsDigest = IrHasher.CanonicalHash(unmodeledContainer);

        var contentBuilder = new IrContentHashBuilder();
        foreach (var h in rowHashes)
            contentBuilder.AppendHash(h);
        var contentHash = contentBuilder.Build();

        var fpBuilder = new IrContentHashBuilder();
        fpBuilder.AppendHash(tablePropsDigest);
        foreach (var fp in cellFingerprints)
            fpBuilder.AppendHash(fp);
        var formatFingerprint = fpBuilder.Build();

        return new IrTable
        {
            Anchor = AnchorFor(IrAnchorKind.Tbl, tbl),
            Rows = IrNodeList.From(rows),
            UnmodeledTablePropsDigest = tablePropsDigest,
            ContentHash = contentHash,
            FormatFingerprint = formatFingerprint,
            Source = ctx.Provenance(tbl),
        };
    }

    private static (IrRow Row, List<IrHash> CellFingerprints) BuildRow(XElement tr, ReadContext ctx)
    {
        var cells = new List<IrCell>();
        var cellFingerprints = new List<IrHash>();
        var rowBuilder = new IrContentHashBuilder();
        rowBuilder.AppendStructure(IrContentHashBuilder.StructureRow);

        foreach (var tc in tr.Elements(W + "tc"))
        {
            var (cell, fingerprints) = BuildCell(tc, ctx);
            cells.Add(cell);
            cellFingerprints.AddRange(fingerprints);
            rowBuilder.AppendHash(cell.ContentHash);
        }

        var row = new IrRow(AnchorFor(IrAnchorKind.Tr, tr), IrNodeList.From(cells), rowBuilder.Build())
        {
            Source = ctx.Provenance(tr),
        };
        return (row, cellFingerprints);
    }

    private static (IrCell Cell, List<IrHash> Fingerprints) BuildCell(XElement tc, ReadContext ctx)
    {
        var tcPr = tc.Element(W + "tcPr");
        int gridSpan = IntAttr(tcPr?.Element(W + "gridSpan"), W + "val") ?? 1;
        var vMerge = MapVMerge(tcPr?.Element(W + "vMerge"));

        var blocks = new List<IrBlock>();
        var fingerprints = new List<IrHash>();
        var cellBuilder = new IrContentHashBuilder();
        cellBuilder.AppendStructure(IrContentHashBuilder.StructureCell);

        foreach (var child in tc.Elements())
        {
            if (child.Name == W + "tcPr")
                continue;
            var block = BuildBlock(child, ctx);
            blocks.Add(block);
            cellBuilder.AppendHash(block.ContentHash);
            fingerprints.Add(block.FormatFingerprint);
        }

        var cell = new IrCell(
            AnchorFor(IrAnchorKind.Tc, tc),
            IrNodeList.From(blocks),
            gridSpan,
            vMerge,
            cellBuilder.Build())
        {
            Source = ctx.Provenance(tc),
        };
        return (cell, fingerprints);
    }

    private static IrVMerge MapVMerge(XElement? vMerge)
    {
        if (vMerge is null)
            return IrVMerge.None;
        var val = (string?)vMerge.Attribute(W + "val");
        return val == "restart" ? IrVMerge.Restart : IrVMerge.Continue;
    }

    // --- section break ----------------------------------------------------

    private static IrSectionBreak BuildSectionBreak(XElement sectPr, ReadContext ctx)
    {
        var pgSz = sectPr.Element(W + "pgSz");
        int? pageWidth = IntAttr(pgSz, W + "w");
        int? pageHeight = IntAttr(pgSz, W + "h");
        bool? landscape = (string?)pgSz?.Attribute(W + "orient") switch
        {
            "landscape" => true,
            null => null,
            _ => false,
        };

        var pgMar = sectPr.Element(W + "pgMar");
        int? marginTop = IntAttr(pgMar, W + "top");
        int? marginBottom = IntAttr(pgMar, W + "bottom");
        int? marginLeft = IntAttr(pgMar, W + "left");
        int? marginRight = IntAttr(pgMar, W + "right");

        string? sectionType = AttrVal(sectPr.Element(W + "type"));

        var digest = UnmodeledDigest(sectPr, SectPrConsumed);

        var format = new IrSectionFormat
        {
            PageWidthTwips = pageWidth,
            PageHeightTwips = pageHeight,
            Landscape = landscape,
            MarginTopTwips = marginTop,
            MarginBottomTwips = marginBottom,
            MarginLeftTwips = marginLeft,
            MarginRightTwips = marginRight,
            SectionType = sectionType,
            UnmodeledDigest = digest,
        };

        // ContentHash: a single opaque hash of the whole sectPr — deterministic and simple.
        var contentBuilder = new IrContentHashBuilder();
        contentBuilder.AppendHash(IrHasher.CanonicalHash(sectPr));

        return new IrSectionBreak
        {
            Anchor = AnchorFor(IrAnchorKind.Sec, sectPr),
            Format = format,
            ContentHash = contentBuilder.Build(),
            FormatFingerprint = IrHasher.FingerprintSectionFormat(format),
            Source = ctx.Provenance(sectPr),
        };
    }

    // --- opaque block -----------------------------------------------------

    private static IrOpaqueBlock BuildOpaqueBlock(XElement el, ReadContext ctx) =>
        new()
        {
            Anchor = AnchorFor(IrAnchorKind.Unk, el),
            ElementName = el.Name,
            ContentHash = IrHasher.CanonicalHash(el),
            FormatFingerprint = EmptyUnmodeledDigest,
            Source = ctx.Provenance(el),
        };

    // --- anchor index -----------------------------------------------------

    private static void IndexBlock(IrBlock block, Dictionary<string, IrBlock> index)
    {
        var key = block.Anchor.ToString();
        if (!index.TryAdd(key, block))
            throw new DocxodusException($"Duplicate IR anchor '{key}' (invariant violation).");

        if (block is IrTable table)
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                    foreach (var child in cell.Blocks)
                        IndexBlock(child, index);
    }

    // --- helpers ----------------------------------------------------------

    /// <summary>
    /// Canonical-hash a synthetic <c>&lt;unmodeled&gt;</c> container holding clones of every child
    /// of <paramref name="props"/> whose name is NOT in <paramref name="consumed"/> and is not the
    /// optional <paramref name="alsoConsumed"/> name (used for conditionally-consumed children so
    /// callers need not allocate a fresh set). When there are no leftovers the result is the cached
    /// empty-container digest (§6.4).
    /// </summary>
    private static IrHash UnmodeledDigest(XElement props, HashSet<XName> consumed, XName? alsoConsumed = null)
    {
        var leftovers = props.Elements()
            .Where(e => !consumed.Contains(e.Name) && e.Name != alsoConsumed)
            .ToList();
        if (leftovers.Count == 0)
            return EmptyUnmodeledDigest;

        var container = new XElement("unmodeled");
        foreach (var e in leftovers)
            container.Add(new XElement(e));
        return IrHasher.CanonicalHash(container);
    }

    private static string? AttrVal(XElement? el) => (string?)el?.Attribute(W + "val");

    private static int? IntAttr(XElement? el, XName name)
    {
        var raw = (string?)el?.Attribute(name);
        if (raw is null)
            return null;
        return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v)
            ? v
            : null;
    }

    /// <summary>
    /// OOXML toggle semantics: absent element → null; present with no <c>w:val</c> or a truthy
    /// value (1/true/on) → true; an explicit falsy value (0/false/off) → false.
    /// </summary>
    private static bool? Toggle(XElement? el)
    {
        if (el is null)
            return null;
        var val = (string?)el.Attribute(W + "val");
        if (val is null)
            return true;
        return val switch
        {
            "0" or "false" or "off" => false,
            _ => true,
        };
    }
}
