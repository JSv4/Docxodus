#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Ir;

/// <summary>
/// Reads an OOXML word-processing document into the Document IR (spec §5). The reader is
/// <em>total</em>: any child it does not model is preserved as an <see cref="IrOpaqueBlock"/> (or
/// <see cref="IrOpaqueInline"/> at run level), so it never throws on weird-but-valid OOXML. It never
/// mutates the caller's document — it works over a private copy.
/// </summary>
/// <remarks>
/// Pipeline: copy the caller's bytes → normalize tracked revisions per
/// <see cref="IrReaderOptions.RevisionView"/> → open the copy → assign deterministic Unids
/// (same call <c>WmlToMarkdownConverter</c> makes) → walk <c>w:body</c> children in document order →
/// walk the remaining requested scopes (headers/footers, footnotes/endnotes, comments) in the SAME
/// part order as <c>WmlToMarkdownConverter.BuildAnchorIndex</c>. <see cref="IrReaderOptions.Scopes"/>
/// selects which scopes are read; unselected scopes are emitted as empty stores/lists.
/// </remarks>
internal static class IrReader
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // Drawing element/attribute names for inline-image promotion (established constants in
    // PtOpenXmlUtil). Fully qualified because the local `R` field above shadows the `Docxodus.R`
    // namespace-constants class.
    private static readonly XName ABlip = Docxodus.A.blip;       // a:blip (drawingml)
    private static readonly XName REmbed = Docxodus.R.embed;     // r:embed (relationships)
    private static readonly XName WpExtent = Docxodus.WP.extent; // wp:extent (wordprocessingDrawing)
    private static readonly XName WpDocPr = Docxodus.WP.docPr;   // wp:docPr (wordprocessingDrawing)
    private static readonly XName WTxbxContent = W + "txbxContent"; // w:txbxContent (textbox body, wps + VML)

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

        // Registries are built BEFORE the body walk so paragraph list resolution can chase
        // numPr → IrNum → IrAbstractNum → level format (rule M1.3). The style registry feeds the
        // style-chain numPr lookup; the numbering registry resolves the facts.
        var styles = BuildStyleRegistry(main);
        var numbering = BuildNumberingRegistry(main);

        // Sources pins the parsed XDocuments so per-node IrProvenance.Element pointers stay alive. When
        // RetainSources is off we leave it EMPTY: the working XDocuments still exist for the duration of
        // this Read (the walk needs them), but nothing in the returned snapshot roots them, so they
        // become collectible once Read returns. Part-URI facts survive via IrScope.PartUri instead.
        var retain = options.RetainSources;
        var sources = new Dictionary<Uri, XDocument>();
        if (retain)
            sources[partUri] = mainXDoc;

        // --- body scope (always walked; the IrDocument requires a body) -------------------------
        // Comment targets (N15) are recorded only while walking the BODY: the markdown projection's
        // comment ranges live in the main document part, so cross-scope ranges are out of scope.
        var commentTracker = options.Scopes.HasFlag(IrScopes.Comments)
            ? new CommentTracker()
            : null;
        var bodyCtx = new ReadContext(partUri, main, styles, numbering, "body", commentTracker,
            retainSources: retain);

        var body = root.Element(W + "body")
            ?? throw new DocxodusException("Document has no w:body element.");

        var blocks = new List<IrBlock>();
        foreach (var child in body.Elements())
            AppendBlocks(child, bodyCtx, blocks);
        // Each paragraph finishes itself (BuildParagraph), so any range still open at end-of-body has
        // its provisional spans buffered but never committed — an orphan start, discarded (totality).

        // 5. Anchor index over ALL scopes' blocks (rows/cells are positional, not blocks). Collision
        // checking spans scopes, and IndexBlock throws on any duplicate kind:scope:unid string.
        // INVARIANT: collision-safety depends on every scope emitting a unique scope-name prefix.
        // Deterministic Unids are content-addressable per part, so identical content appearing in two
        // different parts yields identical Unids — without the per-scope name prefix those would
        // collide. Each scope's distinct name (body, hdr1, fn, …) is what keeps anchors unique.
        var anchorIndex = new Dictionary<string, IrBlock>(StringComparer.Ordinal);
        foreach (var b in blocks)
            IndexBlock(b, anchorIndex);

        // --- header / footer scopes (rule M1.3) -------------------------------------------------
        var headers = IrNodeList.Empty<IrHeaderFooter>();
        var footers = IrNodeList.Empty<IrHeaderFooter>();
        if (options.Scopes.HasFlag(IrScopes.HeadersFooters))
        {
            headers = ReadHeaderFooterScopes(main, body, styles, numbering, sources, anchorIndex,
                main.HeaderParts.Cast<OpenXmlPart>(), "hdr", W + "headerReference", retain);
            footers = ReadHeaderFooterScopes(main, body, styles, numbering, sources, anchorIndex,
                main.FooterParts.Cast<OpenXmlPart>(), "ftr", W + "footerReference", retain);
        }

        // --- footnote / endnote scopes (rule M1.3) ----------------------------------------------
        var footnotes = IrNoteStore.Empty;
        var endnotes = IrNoteStore.Empty;
        if (options.Scopes.HasFlag(IrScopes.Notes))
        {
            footnotes = ReadNoteStore(main, main.FootnotesPart, styles, numbering, sources,
                anchorIndex, "fn", W + "footnote", retain);
            endnotes = ReadNoteStore(main, main.EndnotesPart, styles, numbering, sources,
                anchorIndex, "en", W + "endnote", retain);
        }

        // --- comment scope (rule M1.3 + N15 record-half) ----------------------------------------
        var comments = IrCommentStore.Empty;
        if (options.Scopes.HasFlag(IrScopes.Comments))
            comments = ReadCommentStore(main, styles, numbering, sources, anchorIndex, commentTracker!,
                retain);

        return new IrDocument
        {
            Body = new IrScope("body", IrNodeList.From(blocks), partUri),
            Headers = headers,
            Footers = footers,
            Footnotes = footnotes,
            Endnotes = endnotes,
            Comments = comments,
            Styles = styles,
            Numbering = numbering,
            ThemeFonts = BuildThemeFonts(main),
            AnchorIndex = anchorIndex,
            Sources = sources,
        };
    }

    /// <summary>
    /// Carries the part URI (for provenance), the owning <see cref="MainDocumentPart"/> (for
    /// resolving image relationships), and a per-<see cref="Read"/> image-bytes hash cache through
    /// the recursive walk. The cache keys on embed rel id so a logo reused 50 times hashes once.
    /// </summary>
    private sealed class ReadContext
    {
        public ReadContext(Uri partUri, MainDocumentPart main,
            IrStyleRegistry styles, IrNumberingRegistry numbering,
            string scope, CommentTracker? commentTracker = null, int textboxDepth = 0,
            bool retainSources = true, OpenXmlPart? owningPart = null)
        {
            PartUri = partUri;
            Main = main;
            Styles = styles;
            Numbering = numbering;
            Scope = scope;
            CommentTracker = commentTracker;
            TextboxDepth = textboxDepth;
            RetainSources = retainSources;
            // The part whose relationships resolve drawing/image/diagram rel ids for opaque canonicalization.
            // Defaults to the main part (the body scope); header/footer/note scopes pass their own part so a
            // drawing's r:embed resolves against the relationships that actually own it.
            OwningPart = owningPart ?? main;
        }

        // Whether per-node provenance pins the source XElement (IrReaderOptions.RetainSources). When
        // false, Provenance() hands back the shared empty instance so the parsed XML can be collected
        // after Read returns. The part URI is still carried (PartUri above) because it is promoted to
        // the scope level regardless of this flag.
        public bool RetainSources { get; }

        // How many w:txbxContent bodies deep this context is. Bounds textbox nesting recursion
        // (textboxes can legally nest) the same way MaxSdtDepth bounds content-control nesting.
        // A textbox walked at the cap is preserved opaquely instead of recursing further.
        public int TextboxDepth { get; }

        /// <summary>A child context one textbox level deeper. The comment tracker is intentionally
        /// NOT carried inside a textbox: textbox inner paragraphs are their own blocks with their own
        /// anchors, and the body-scope comment-range offset bookkeeping (which counts visible text in
        /// document order against the CURRENT block) has no defined meaning across a textbox boundary —
        /// matching the reader's existing "no offsets inside a field" stance.</summary>
        public ReadContext IntoTextbox() =>
            new(PartUri, Main, Styles, Numbering, Scope, commentTracker: null, TextboxDepth + 1,
                RetainSources, OwningPart)
            { RelResolver = RelResolver };

        public Uri PartUri { get; }

        public MainDocumentPart Main { get; }

        // The part whose relationships own this scope's drawing/image/diagram rel ids.
        public OpenXmlPart OwningPart { get; }

        // Opaque-canonicalization relationship resolver (rel id → stable content-identity token), lazily
        // built over OwningPart and shared down through IntoTextbox so the per-part byte-hash cache is hit
        // across nested textbox contexts. Backing field assigned via the initializer or the lazy getter.
        private IrRelResolver? _relResolver;
        public IrRelResolver RelResolver
        {
            get => _relResolver ??= new IrRelResolver(OwningPart);
            init => _relResolver = value;
        }

        // The IR scope name carried into every anchor produced under this context ("body", "hdr1",
        // "ftr1", "fn", "en", "cmt"). Threaded so header/note/comment blocks get scope-tagged anchors
        // matching the markdown projection's BuildAnchorIndex naming.
        public string Scope { get; }

        // N15 comment-range tracker, non-null only while walking the BODY scope with comments
        // requested. Records (block anchor, char span) targets as commentRange markers are seen.
        public CommentTracker? CommentTracker { get; }

        // Registries for list resolution (numPr → IrNum → IrAbstractNum → level format) and the
        // pStyle → basedOn walk that finds style-inherited numbering.
        public IrStyleRegistry Styles { get; }

        public IrNumberingRegistry Numbering { get; }

        // embed rel id → (image part URI, hash of its bytes). Negative results are not cached because
        // a missing/wrong-typed rel is rare and cheap to re-check.
        public Dictionary<string, (Uri PartUri, IrHash BytesHash)> ImageHashCache { get; } =
            new(StringComparer.Ordinal);

        // Retained mode pins the source element + part URI on a fresh provenance; retention-off mode
        // hands back the shared empty instance (no element, no part URI — those facts live at the scope
        // level now) so the parsed XDocument becomes collectible after Read.
        public IrProvenance Provenance(XElement element) =>
            RetainSources
                ? new IrProvenance { Element = element, PartUri = PartUri }
                : IrProvenance.Empty;
    }

    // --- comment range tracking (N15 record-half) -------------------------

    /// <summary>
    /// Records <c>w:commentRangeStart</c>/<c>End</c>/<c>w:commentReference</c> spans seen during the
    /// BODY walk into per-comment-id <see cref="IrCommentTarget"/> lists (rule N15). The comment
    /// plumbing elements are still DROPPED from the inline stream (so they never touch ContentHash —
    /// M1.2 behavior is unchanged); this tracker only *observes* their positions.
    /// <para/>
    /// <b>Char-offset rule.</b> An offset is the count of <em>visible text characters</em> emitted so
    /// far within the current block — the summed lengths of the block's <see cref="IrTextRun"/>s. Tabs,
    /// breaks, images, note refs, field/opaque inlines, and all non-text inlines count as 0. This is
    /// the simplest rule that is stable under the N5 run-coalescing pass (which never changes the total
    /// text length of a block) and is documented on <see cref="IrCommentTarget"/>.
    /// <para/>
    /// <b>Cross-block ranges.</b> When a paragraph ends with ranges still open, each open range is
    /// closed at the paragraph's end offset (one <see cref="IrCommentTarget"/> for the touched block)
    /// and re-opened at offset 0 of the next block, so a range spanning N blocks yields N targets — one
    /// per touched block (spec §12 open-question #2, resolved this way).
    /// <para/>
    /// Totality: an orphan range-start (no matching end, or no matching comment) is discarded silently;
    /// a <c>commentReference</c> for a comment with no ranges records a zero-length target at the
    /// reference's offset.
    /// </summary>
    private sealed class CommentTracker
    {
        // An open range's per-block accumulation: the targets touched so far (provisional until the
        // matching commentRangeEnd is seen) plus the start offset within the CURRENT block.
        private sealed class OpenRangeState
        {
            public readonly List<IrCommentTarget> Provisional = new();
            public int StartInCurrentBlock;
        }

        // commentId → completed (end-seen) targets, in document order.
        private readonly Dictionary<string, List<IrCommentTarget>> _targets = new(StringComparer.Ordinal);
        // commentId → range currently open (spanning zero or more blocks until its end).
        private readonly Dictionary<string, OpenRangeState> _open = new(StringComparer.Ordinal);
        // commentIds that have at least one commentReference (for the zero-length fallback when the
        // comment has no committed range). id → first reference (block, offset).
        private readonly Dictionary<string, (IrAnchor Block, int Offset)> _danglingRefs =
            new(StringComparer.Ordinal);

        private IrAnchor _currentBlock;
        private int _charOffset;

        /// <summary>Begin a new block: reset the char counter and re-open carried-over ranges at 0.</summary>
        public void BeginParagraph(IrAnchor blockAnchor)
        {
            _currentBlock = blockAnchor;
            _charOffset = 0;
            foreach (var range in _open.Values)
                range.StartInCurrentBlock = 0;
        }

        /// <summary>Advance the char counter by an emitted text run's length.</summary>
        public void Advance(int textLength) => _charOffset += textLength;

        /// <summary>Open a comment range at the current offset within the current block.</summary>
        public void OpenRange(string id)
        {
            if (id.Length == 0)
                return;
            // A duplicate start for an already-open id is malformed; keep the first (totality).
            if (!_open.ContainsKey(id))
                _open[id] = new OpenRangeState { StartInCurrentBlock = _charOffset };
        }

        /// <summary>Close a comment range at the current offset, committing its provisional per-block
        /// targets (including the final block's) into the completed set.</summary>
        public void CloseRange(string id)
        {
            if (id.Length == 0)
                return;
            if (_open.TryGetValue(id, out var range))
            {
                range.Provisional.Add(new IrCommentTarget(_currentBlock, range.StartInCurrentBlock, _charOffset));
                foreach (var t in range.Provisional)
                    AddTarget(id, t);
                _open.Remove(id);
            }
            // A stray end with no matching start is discarded silently (totality).
        }

        /// <summary>Record a comment reference offset; resolved to a zero-length target at end-of-read
        /// only if the comment had no committed range.</summary>
        public void Reference(string id)
        {
            if (id.Length == 0)
                return;
            if (!_danglingRefs.ContainsKey(id))
                _danglingRefs[id] = (_currentBlock, _charOffset);
        }

        /// <summary>End the current block: for each still-open range, buffer the touched span for this
        /// block as PROVISIONAL (committed only if/when the range's end is later seen) and re-open it
        /// at offset 0 of the next block. A range whose end never arrives leaves its provisional spans
        /// uncommitted — i.e. an orphan range-start is discarded silently (totality).</summary>
        public void FinishParagraph()
        {
            foreach (var range in _open.Values)
                range.Provisional.Add(new IrCommentTarget(_currentBlock, range.StartInCurrentBlock, _charOffset));
        }

        /// <summary>
        /// The targets for a given comment id. A comment with no committed range but a recorded
        /// reference yields a single zero-length target at the reference offset. Targets for comment
        /// ids that no <c>w:comment</c> claims are simply never queried (orphan ranges dropped).
        /// </summary>
        public IReadOnlyList<IrCommentTarget> TargetsFor(string commentId)
        {
            if (_targets.TryGetValue(commentId, out var list) && list.Count > 0)
                return list;
            if (_danglingRefs.TryGetValue(commentId, out var r))
                return new[] { new IrCommentTarget(r.Block, r.Offset, r.Offset) };
            return Array.Empty<IrCommentTarget>();
        }

        private void AddTarget(string id, IrCommentTarget target)
        {
            if (!_targets.TryGetValue(id, out var list))
            {
                list = new List<IrCommentTarget>();
                _targets[id] = list;
            }
            list.Add(target);
        }
    }

    // --- revisions --------------------------------------------------------

    private static WmlDocument ApplyRevisionView(WmlDocument working, RevisionView view)
    {
        switch (view)
        {
            case RevisionView.FailIfPresent:
                // Original (M1.1) throw contract: the narrow ins/del/move/*PrChange set. NOT widened —
                // IrReaderTests.Read_UnknownElement_BecomesOpaque deliberately feeds a
                // w:customXmlInsRangeStart under FailIfPresent expecting it to survive as an opaque
                // block, i.e. it is intentionally NOT treated as "present revision markup" here.
                if (HasRevisionMarkup(working, FailIfPresentNameSet))
                    throw new DocxodusException(
                        "Document contains tracked revisions and RevisionView is FailIfPresent.");
                return working;

            case RevisionView.Accept:
            case RevisionView.Reject:
                // Accepting/rejecting revisions on a document with NO revision markup is a pure no-op
                // round-trip (RevisionProcessor opens, clones, walks, and re-serializes the whole
                // package only to change nothing). A cheap in-memory descendant scan lets the common
                // revision-free document skip that round-trip — the single largest per-Read cost.
                // The scan set (ProcessorActsOnNameSet) is a strict SUPERSET of every element
                // Accept/RejectRevisions acts on (run/paragraph ins/del/move, the *PrChange and
                // tblPrExChange property-revision markers, table cell/grid revisions, deleted text
                // markers, and the customXml*RangeStart range markers), and the scan covers EVERY part
                // the reader consumes (main + headers + footers + footnotes + endnotes + comments) —
                // exactly the parts RevisionProcessor itself transforms. So "no markup found" provably
                // implies the processor would not have changed a byte — output stays identical. See the
                // masking analysis on ProcessorActsOnNameSet for why w:instrText/w:t are deliberately
                // omitted (covered by their w:ins ancestor).
                if (!HasRevisionMarkup(working, ProcessorActsOnNameSet))
                    return working;
                return view == RevisionView.Accept
                    ? RevisionProcessor.AcceptRevisions(working)
                    : RevisionProcessor.RejectRevisions(working);

            default:
                return working;
        }
    }

    // The narrow M1.1 FailIfPresent set (unchanged — preserves the documented throw contract).
    private static readonly HashSet<XName> FailIfPresentNameSet = new()
    {
        W + "ins", W + "del", W + "moveFrom", W + "moveTo", W + "rPrChange", W + "pPrChange",
    };

    // Every element name RevisionProcessor.Accept/RejectRevisions reacts to. A strict superset so a
    // "no markup" scan result guarantees the processor is a no-op (see the Accept/Reject skip above).
    //
    // SET-DRIFT CONTRACT: this set MUST list every element name RevisionProcessor's transforms
    // dispatch on by name. If RevisionProcessor.cs gains/loses a revision element, update BOTH this
    // set AND the hardcoded list in IrRevisionSkipTests.ProcessorActsOnNameSet_MatchesProcessor.
    // Source of truth: the `element.Name == W.*` dispatch in RevisionProcessor.cs
    // (ReverseRevisionsTransform ~ln 96-450 / 604-1252; AcceptAllOtherRevisionsTransform ~ln 1680-1731).
    //
    // Masking analysis for elements RevisionProcessor transforms but that are NOT listed here because a
    // listed ancestor provably covers them:
    //   - w:instrText: only transformed by the Reject path (RevisionProcessor.cs:1223) and ONLY when
    //     rri.InInsert is set, which is reached exclusively inside a w:ins subtree (the InInsert flag is
    //     set when recursing through w:ins; see ReverseRevisionsTransform). No w:ins ⇒ no transform, and
    //     w:ins IS scanned. Genuinely masked, so omitted.
    //   - w:t: same — only rewritten to w:delText under rri.InInsert (RevisionProcessor.cs:1232), masked
    //     by w:ins. (Scanning w:t would defeat the optimization entirely; every document has w:t.)
    // By contrast w:delText (RevisionProcessor.cs:1241 / :1729) and w:delInstrText
    // (RevisionProcessor.cs:1214 / :1728) are transformed UNCONDITIONALLY by name — the code does NOT
    // gate them on a w:del/w:ins ancestor. Schema-valid documents only place them under w:del, but the
    // skip's soundness guarantee must not rely on producer validity, so both are listed here.
    private static readonly HashSet<XName> ProcessorActsOnNameSet = new()
    {
        W + "ins", W + "del", W + "moveFrom", W + "moveTo",
        W + "moveFromRangeStart", W + "moveFromRangeEnd", W + "moveToRangeStart", W + "moveToRangeEnd",
        W + "rPrChange", W + "pPrChange", W + "sectPrChange", W + "tblPrChange", W + "tblGridChange",
        W + "trPrChange", W + "tcPrChange", W + "tblPrExChange", W + "numberingChange",
        W + "cellIns", W + "cellDel", W + "cellMerge",
        W + "delText", W + "delInstrText",
        W + "customXmlInsRangeStart", W + "customXmlDelRangeStart",
        W + "customXmlMoveFromRangeStart", W + "customXmlMoveToRangeStart",
        // RangeEnd markers are removed unconditionally by Accept (RevisionProcessor.cs:1693-1698).
        // Schema-valid documents always pair them with a scanned RangeStart, but — same standard as
        // delText above — soundness must not rely on producer validity, so they are listed too.
        W + "customXmlInsRangeEnd", W + "customXmlDelRangeEnd",
        W + "customXmlMoveFromRangeEnd", W + "customXmlMoveToRangeEnd",
    };

    // The hardcoded revision-element list the set-drift guard test asserts against. Kept internal so
    // IrRevisionSkipTests can compare it to ProcessorActsOnNameSet without reflecting over the private
    // field — both must change together when RevisionProcessor's dispatch changes.
    internal static IReadOnlyCollection<XName> ProcessorActsOnNamesForTest => ProcessorActsOnNameSet;

    // A w:ins child of w:numPr (inserted numbering, RevisionProcessor.cs:96) is itself a w:ins element,
    // so the local-name "ins" scan above already catches it — no separate numPr probe is needed. The
    // scan matches by full XName via DescendantsAndSelf, which visits that nested w:ins like any other.
    private static bool HasRevisionMarkup(WmlDocument working, HashSet<XName> names)
    {
        using var stream = new OpenXmlMemoryStreamDocument(working);
        using var wdoc = stream.GetWordprocessingDocument();
        var main = wdoc.MainDocumentPart;
        if (main is null)
            return false;
        // The reader consumes the main part AND headers/footers/footnotes/endnotes/comments, and
        // RevisionProcessor processes every one of those parts too. Scanning only the main part would
        // let a header-only (or footnote-only, …) revision slip past the skip, so the processor's
        // transform would be silently bypassed. Scan exactly the parts the reader walks.
        foreach (var part in RevisionScannablePartsOf(main))
        {
            var root = part.GetXDocument().Root;
            if (root is null)
                continue;
            foreach (var e in root.DescendantsAndSelf())
                if (names.Contains(e.Name))
                    return true;
        }
        return false;
    }

    // The parts whose revision markup the skip must account for: the main document plus every part the
    // reader pulls block content from (headers, footers, footnotes, endnotes, comments). Mirrors the
    // parts walked in Read (main.HeaderParts/FooterParts/FootnotesPart/EndnotesPart/CommentsPart).
    private static IEnumerable<OpenXmlPart> RevisionScannablePartsOf(MainDocumentPart main)
    {
        yield return main;
        foreach (var h in main.HeaderParts)
            yield return h;
        foreach (var f in main.FooterParts)
            yield return f;
        if (main.FootnotesPart is not null)
            yield return main.FootnotesPart;
        if (main.EndnotesPart is not null)
            yield return main.EndnotesPart;
        if (main.WordprocessingCommentsPart is not null)
            yield return main.WordprocessingCommentsPart;
    }

    // --- block dispatch ---------------------------------------------------

    /// <summary>
    /// Append the IR block(s) produced by a single body/cell child element to <paramref name="sink"/>.
    /// Almost every element yields exactly one block; the exception is N12: a block-level
    /// <c>w:sdt</c> (content control) contributes nothing itself — its <c>w:sdtContent</c> children
    /// are walked through the normal block walker so each inner <c>w:p</c>/<c>w:tbl</c> gets its own
    /// anchor from its own <c>pt:Unid</c>. The SDT wrapper is recoverable via the inner blocks'
    /// provenance ancestors. Multiple or zero content children are handled naturally (we walk
    /// whatever is there). SDTs can nest, so this recurses.
    /// </summary>
    private static void AppendBlocks(XElement el, ReadContext ctx, List<IrBlock> sink, int depth = 0)
    {
        // N3 at BLOCK level: a bookmark marker (bookmarkStart/bookmarkEnd) — or comment-range plumbing — that
        // appears as a DIRECT body/cell child (a sibling of w:p, legal OOXML) is dropped, exactly as the inline
        // walker drops it inside a paragraph. WmlComparer strips ALL bookmarks in PreProcessMarkup
        // (MarkupSimplifier RemoveBookmarks=true), so a body-level bookmark is invisible to the comparison; the
        // IR previously modeled it as an IrOpaqueBlock, which made an inserted/orphaned body-level bookmarkEnd
        // a spurious content block that the markup round-trip could not toggle (it is not run content, so it is
        // emitted verbatim and survives both accept and reject) — WC022's stray w:bookmarkEnd and the
        // WC-BodyBookmarks section bookmarks. Dropping it mirrors the oracle and the inline N3 rule. Bookmarks
        // never surface in the markdown projection either, so this does not perturb M1.4 equivalence.
        if (IsDroppedParagraphChild(el.Name))
            return;

        if (el.Name == W + "sdt")
        {
            // Pathologically deep SDT nesting would recurse without bound; cap it and preserve the
            // whole subtree opaquely so totality holds without stack risk (the cap is far beyond any
            // legitimate content control nesting).
            if (depth >= MaxSdtDepth)
            {
                sink.Add(BuildOpaqueBlock(el, ctx));
                return;
            }
            var content = el.Element(W + "sdtContent");
            if (content is not null)
            {
                var before = sink.Count;
                foreach (var inner in content.Elements())
                    AppendBlocks(inner, ctx, sink, depth + 1);
                // Mark every block this SDT delivered (transitively) so the markdown emitter can mirror
                // the oracle's EmitBlocks, which skips SDT wrappers and thus never renders these blocks
                // (they remain present + indexed, matching the oracle's Descendants-based index).
                for (int i = before; i < sink.Count; i++)
                    sink[i] = MarkFromBlockSdt(sink[i]);
            }
            return;
        }
        sink.Add(BuildBlock(el, ctx));
    }

    /// <summary>Return <paramref name="block"/> with its provenance's <c>FromBlockSdt</c> set, preserving
    /// the original source element/part. Equality-neutral (provenance is excluded from record equality),
    /// so this never perturbs the block's value/hash.</summary>
    private static IrBlock MarkFromBlockSdt(IrBlock block)
    {
        var src = block.Source;
        var marked = new IrProvenance { Element = src.Element, PartUri = src.PartUri, FromBlockSdt = true };
        return block with { Source = marked };
    }

    /// <summary>
    /// Maximum content-control (<c>w:sdt</c>/<c>w:smartTag</c>) nesting the reader unwraps before
    /// falling back to an opaque node. Real documents nest only a handful deep; the cap exists
    /// purely to bound recursion against adversarial/corrupt input.
    /// </summary>
    private const int MaxSdtDepth = 64;

    /// <summary>
    /// Maximum <c>w:txbxContent</c> nesting the reader walks before falling back to an opaque inline.
    /// Textboxes can legally nest (a shape inside a textbox can itself carry a textbox); the cap
    /// bounds recursion against adversarial/corrupt input, mirroring <see cref="MaxSdtDepth"/>.
    /// </summary>
    private const int MaxTextboxDepth = 16;

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

    private static IrAnchor AnchorFor(IrAnchorKind kind, XElement el, ReadContext ctx) =>
        new(kind, ctx.Scope, Unid(el));

    // --- paragraph --------------------------------------------------------

    private static IrParagraph BuildParagraph(XElement p, ReadContext ctx)
    {
        var kindToken = WmlToMarkdownConverter.KindFor(p);
        var kind = kindToken is null ? IrAnchorKind.P : IrAnchor.KindFromToken(kindToken);

        var pPr = p.Element(W + "pPr");
        var paraFormat = MapParaFormat(pPr);
        var listInfo = pPr is null ? null : ResolveListInfo(pPr, ctx);

        // N15: announce the block to the comment tracker so range offsets accumulate against this
        // paragraph's anchor, then close any range still open at paragraph end (carrying it into the
        // next block at offset 0 for cross-block ranges).
        var anchor = AnchorFor(kind, p, ctx);
        ctx.CommentTracker?.BeginParagraph(anchor);

        // Walk the paragraph's children (skipping w:pPr) through the shared inline walker, which
        // handles run content, hyperlinks (N14), and the field state machine (N9), then applies
        // the N10 empty-drop + N5 coalescing post-process to the top-level inline list.
        var processed = WalkInlines(p.Elements().Where(c => c.Name != W + "pPr"), ctx);
        ctx.CommentTracker?.FinishParagraph();

        var contentHash = ComputeParagraphContentHash(processed);
        var formatFingerprint = IrHasher.FingerprintBlock(paraFormat, RunFormatsInOrder(processed));

        // An in-pPr w:sectPr marks an in-document section transition (the markdown projection emits a
        // {#sec:…} + thematic break after the paragraph and indexes the sectPr). Capture its anchor so
        // the emitter and anchor index can reproduce that without re-walking the skipped pPr. The
        // trailing top-level body sectPr is a standalone IrSectionBreak block, handled elsewhere.
        var inlineSectPr = pPr?.Element(W + "sectPr");
        IrAnchor? inlineSectAnchor = inlineSectPr is null
            ? null
            : AnchorFor(IrAnchorKind.Sec, inlineSectPr, ctx);

        // Resolve the auto-number marker against the LIVE package while we have it (see
        // IrParagraph.ResolvedListMarker for the rationale). RetrieveListItem returns null for
        // non-list-items; it is the exact string the markdown projection consumes. Tolerance-wrapped
        // so a malformed numbering setup degrades to null (no marker) rather than aborting the read.
        string? resolvedMarker = ResolveListMarkerText(p, ctx);

        return new IrParagraph
        {
            Anchor = anchor,
            Format = paraFormat,
            List = listInfo,
            Inlines = IrNodeList.From(processed),
            InlineSectionBreakAnchor = inlineSectAnchor,
            ResolvedListMarker = resolvedMarker,
            // The oracle's structural IsListItem verdict (numPr present inline or via the style chain,
            // numId-agnostic) — drives the emitter's trailing-blank rule for heading/Subtitle styles
            // whose chain carries a bare numPr (no numId), where List stays null. See
            // IrParagraph.IsListItemForLayout.
            IsListItemForLayout = WmlToMarkdownConverter.IsListItem(p),
            ContentHash = contentHash,
            FormatFingerprint = formatFingerprint,
            Source = ctx.Provenance(p),
        };
    }

    /// <summary>
    /// Resolve the raw auto-number marker for <paramref name="p"/> via
    /// <see cref="ListItemRetriever.RetrieveListItem(WordprocessingDocument, XElement, ListItemRetrieverSettings)"/>
    /// against the live package (<see cref="ReadContext.Main"/>'s owning <see cref="WordprocessingDocument"/>).
    /// Returns null when the document has no numbering part, the paragraph is not a list item, or any
    /// tolerable fault occurs — matching the projection's null-on-failure contract.
    /// </summary>
    private static string? ResolveListMarkerText(XElement p, ReadContext ctx)
    {
        if (ctx.Main.OpenXmlPackage is not WordprocessingDocument wdoc)
            return null;
        try
        {
            var resolved = ListItemRetriever.RetrieveListItem(wdoc, p, new ListItemRetrieverSettings());
            return string.IsNullOrEmpty(resolved) ? null : resolved;
        }
        catch
        {
            // The oracle's ResolveListMarker / ListNumberResolver.Resolve both wrap RetrieveListItem in
            // a broad catch (it throws DocxodusException on malformed numbering setups, e.g. an ilvl set
            // twice). We mirror that EXACT broad catch here — scoped to this single resolver call — so a
            // pathological numbering definition degrades to "no marker", matching the projection and
            // preserving reader totality. This is the resolver's own contract, not a reader-wide policy.
            return null;
        }
    }

    // --- inline walk (runs, hyperlinks N14, fields N9) --------------------

    /// <summary>
    /// Walk a flat sequence of inline-level OOXML elements (a paragraph's or a
    /// <c>w:hyperlink</c>'s children) into the typed inline list, applying the N9 field state
    /// machine and N14 hyperlink promotion inline, then the N10 empty-drop and N5 coalescing
    /// post-process. The same logic serves both the paragraph top level and hyperlink interiors,
    /// so empty-drop/coalescing happen within each inline list independently (a hyperlink's runs
    /// coalesce among themselves, not across the link boundary).
    /// </summary>
    private static List<IrInline> WalkInlines(IEnumerable<XElement> children, ReadContext ctx)
    {
        var walker = new InlineWalker(ctx);
        foreach (var child in children)
            walker.Feed(child);
        var emitted = walker.Finish();
        return CoalesceRuns(DropEmptyTextRuns(emitted));
    }

    /// <summary>
    /// Stateful walker driving the field (N9) state machine across a paragraph's / hyperlink's
    /// child sequence. Non-field inlines are emitted directly; between a <c>w:fldChar
    /// fldCharType="begin"</c> and its matching <c>end</c>, run content is diverted into a
    /// captured field (instruction text while in the pre-separate phase, result inlines after a
    /// <c>separate</c>). Fields can nest: an inner <c>begin</c> seen while already capturing is
    /// depth-counted and flattened into the outermost field (its instr text appends to the outer
    /// instruction, its result inlines append to the outer result) — the simplest behavior that
    /// loses no content. A <c>begin</c> with no matching <c>end</c> by the end of the sequence
    /// falls back to emitting every captured element as an opaque inline so nothing is lost.
    /// </summary>
    private sealed class InlineWalker
    {
        private readonly ReadContext _ctx;
        private readonly List<IrInline> _output = new();

        // Field capture state. _fieldDepth > 0 means we are inside one or more nested fields.
        private int _fieldDepth;
        private bool _inResult;                 // true once the (outermost) field hit "separate".
        private readonly StringBuilder _instruction = new();
        private readonly List<IrInline> _result = new();
        // Raw captured elements, kept so an unterminated field can fall back to opaque losslessly.
        private readonly List<XElement> _captured = new();

        public InlineWalker(ReadContext ctx) => _ctx = ctx;

        public void Feed(XElement child) => Feed(child, 0);

        private void Feed(XElement child, int depth)
        {
            if (child.Name == W + "r")
            {
                FeedRun(child);
            }
            else if (child.Name == W + "hyperlink")
            {
                EmitInline(BuildHyperlink(child, _ctx), child);
            }
            else if (child.Name == W + "fldSimple")
            {
                EmitInline(BuildFldSimple(child, _ctx), child);
            }
            else if (depth >= MaxSdtDepth && (child.Name == W + "sdt" || child.Name == W + "smartTag"))
            {
                // Pathologically deep inline content-control nesting: bound the recursion and
                // preserve the whole subtree opaquely rather than risk a stack overflow.
                EmitInline(new IrOpaqueInline(child.Name, IrHasher.CanonicalHash(child, _ctx.RelResolver)), child);
            }
            else if (child.Name == W + "sdt")
            {
                // N12: inline content control — splice w:sdtContent's children into the inline
                // stream so the runs inside join the paragraph and coalesce normally. The spliced
                // text runs are flagged FromInlineSdt (equality-neutral) so the markdown emitter can
                // mirror the oracle, which drops inline-SDT content from the rendered markdown.
                var content = child.Element(W + "sdtContent");
                if (content is not null)
                    SpliceWithInlineSdtMark(() =>
                    {
                        foreach (var inner in content.Elements())
                            Feed(inner, depth + 1);
                    });
            }
            else if (child.Name == W + "smartTag")
            {
                // N12: w:smartTag wraps runs (and can nest other smartTags) — splice its children
                // into the inline stream the same way; recursion handles nesting. Its own
                // w:smartTagPr metadata child is not content and is skipped. Spliced runs flagged as
                // for w:sdt above.
                SpliceWithInlineSdtMark(() =>
                {
                    foreach (var inner in child.Elements().Where(e => e.Name != W + "smartTagPr"))
                        Feed(inner, depth + 1);
                });
            }
            else if (child.Name == W + "proofErr")
            {
                // N2: pure noise, never emit.
            }
            else if (child.Name == W + "commentRangeStart" || child.Name == W + "commentRangeEnd")
            {
                // N15: comment-range plumbing is dropped from the inline stream (no ContentHash
                // impact) but its position is recorded into the comment-target tracker, only while
                // walking the body and never inside a field (offset semantics there are undefined).
                HandleCommentRangeMarker(child);
            }
            else if (IsDroppedParagraphChild(child.Name))
            {
                // N3 (bookmarks).
            }
            else
            {
                EmitInline(new IrOpaqueInline(child.Name, IrHasher.CanonicalHash(child, _ctx.RelResolver)), child);
            }
        }

        /// <summary>
        /// Record a paragraph-level <c>w:commentRangeStart</c>/<c>End</c> into the tracker (rule N15).
        /// Only fires while walking the body scope (tracker non-null) and outside a field capture
        /// (where char offsets are undefined). The element is otherwise dropped — it never reaches the
        /// inline stream or any hash.
        /// </summary>
        private void HandleCommentRangeMarker(XElement marker)
        {
            if (_ctx.CommentTracker is null || _fieldDepth > 0)
                return;
            var id = (string?)marker.Attribute(W + "id") ?? "";
            if (marker.Name == W + "commentRangeStart")
                _ctx.CommentTracker.OpenRange(id);
            else
                _ctx.CommentTracker.CloseRange(id);
        }

        private void FeedRun(XElement r)
        {
            var rPr = r.Element(W + "rPr");
            var runFormat = MapRunFormat(rPr);

            foreach (var child in r.Elements())
            {
                if (child.Name == W + "rPr")
                    continue;
                if (child.Name == W + "fldChar")
                {
                    HandleFldChar(child);
                    continue;
                }
                // N15: run-level comment plumbing. commentRangeStart/End/Reference are dropped from
                // the inline stream (no ContentHash impact) but their positions are recorded into the
                // tracker — outside fields only (offset semantics inside a field are undefined).
                if (child.Name == W + "commentRangeStart" || child.Name == W + "commentRangeEnd")
                {
                    HandleCommentRangeMarker(child);
                    continue;
                }
                if (child.Name == W + "commentReference")
                {
                    if (_ctx.CommentTracker is not null && _fieldDepth == 0)
                        _ctx.CommentTracker.Reference((string?)child.Attribute(W + "id") ?? "");
                    continue; // never emitted; never hashed.
                }
                if (_fieldDepth > 0)
                {
                    _captured.Add(child);
                    if (!_inResult)
                    {
                        // Pre-separate: accumulate instruction text (w:instrText / w:delInstrText).
                        if (child.Name == W + "instrText" || child.Name == W + "delInstrText")
                            _instruction.Append(child.Value);
                        // Other pre-separate content is field plumbing; ignore for the instruction
                        // string but keep captured for the unterminated-field fallback.
                        continue;
                    }
                    // Post-separate: divert run content into the field result.
                    EmitAndTrack(child, rPr, runFormat, _result);
                    continue;
                }

                EmitAndTrack(child, rPr, runFormat, _output);
            }
        }

        /// <summary>
        /// Emit a run child into <paramref name="sink"/> and advance the comment-target char counter
        /// (rule N15) by the visible-text length the child contributed — measured as the growth in
        /// <see cref="IrTextRun"/> characters appended to <paramref name="sink"/>. Non-text inlines
        /// (tab, break, image, note ref, opaque) contribute 0, matching the documented offset rule.
        /// </summary>
        private void EmitAndTrack(XElement child, XElement? rPr, IrRunFormat runFormat, List<IrInline> sink)
        {
            var tracker = _ctx.CommentTracker;
            if (tracker is null)
            {
                EmitRunChild(child, rPr, runFormat, _ctx, sink);
                return;
            }
            int before = sink.Count;
            EmitRunChild(child, rPr, runFormat, _ctx, sink);
            int added = 0;
            for (int i = before; i < sink.Count; i++)
                if (sink[i] is IrTextRun tr)
                    added += tr.Text.Length;
            if (added > 0)
                tracker.Advance(added);
        }

        private void HandleFldChar(XElement fldChar)
        {
            var type = (string?)fldChar.Attribute(W + "fldCharType");
            switch (type)
            {
                case "begin":
                    _fieldDepth++;
                    // Inner begins flatten into the outer field (depth-counted), so only the
                    // outermost begin resets the capture buffers.
                    if (_fieldDepth == 1)
                    {
                        _inResult = false;
                        _instruction.Clear();
                        _result.Clear();
                        _captured.Clear();
                    }
                    break;
                case "separate":
                    // Only the outermost separate flips us into result-capture. Inner separates
                    // are swallowed (their result content flattens into the outer result).
                    if (_fieldDepth == 1)
                        _inResult = true;
                    break;
                case "end":
                    if (_fieldDepth == 0)
                        break; // stray end with no begin: ignore (totality).
                    _fieldDepth--;
                    if (_fieldDepth == 0)
                    {
                        // Outermost field closed: emit one IrFieldRun. CachedResult is empty for
                        // instruction-only fields (no separate seen).
                        EmitInline(new IrFieldRun(
                            _instruction.ToString(),
                            IrNodeList.From(new List<IrInline>(_result))));
                        _inResult = false;
                        _result.Clear();
                        _captured.Clear();
                    }
                    break;
                default:
                    break; // unknown fldCharType: ignore.
            }
        }

        /// <summary>
        /// Flush the walker's accumulated inlines. When the sequence ended mid-field (a
        /// <c>begin</c> with no matching <c>end</c>) there are two cases:
        /// <list type="bullet">
        /// <item><b>The field reached its <c>separate</c></b> (instruction + result, just no closing
        /// <c>end</c> before the paragraph ends — e.g. a TOC field whose <c>end</c> is implied at
        /// paragraph close). Word still displays the last-computed result, and the oracle's
        /// field-unaware <c>GroupInlineRuns</c>/<c>Descendants(w:t)</c> both see that result text. We
        /// therefore emit a normal run-based <see cref="IrFieldRun"/> (instruction + result), exactly
        /// as the <c>end</c> handler would — so the result flows into the rendered markdown AND the
        /// TextPreview, never dropped. The instruction-phase plumbing is flattened away just like a
        /// terminated field.</item>
        /// <item><b>The field never reached <c>separate</c></b> (instruction-only, unterminated):
        /// nothing committed to <c>_output</c> — the instruction-phase run text, hyperlinks,
        /// fldSimples, and opaque elements were all diverted into <c>_captured</c>. To lose no content
        /// we re-emit every captured element as an <see cref="IrOpaqueInline"/>.</item>
        /// </list>
        /// </summary>
        public List<IrInline> Finish()
        {
            if (_fieldDepth > 0)
            {
                if (_inResult)
                {
                    // Implied end-at-paragraph-close: commit the field with its result, mirroring the
                    // "end" handler. EmitInline at depth 0 (we force the depth to 0 first) appends it
                    // to _output. The result text now reaches both the rendered markdown and the
                    // TextPreview, matching the oracle's raw Descendants(w:t) view.
                    _fieldDepth = 0;
                    EmitInline(new IrFieldRun(
                        _instruction.ToString(),
                        IrNodeList.From(new List<IrInline>(_result))));
                }
                else
                {
                    // Instruction-only unterminated field: re-emit captured plumbing opaquely so no
                    // content is lost (each captured element is canonical-hashed into an opaque inline).
                    foreach (var el in _captured)
                        _output.Add(new IrOpaqueInline(el.Name, IrHasher.CanonicalHash(el, _ctx.RelResolver)));
                    _fieldDepth = 0;
                }
                _inResult = false;
                _result.Clear();
                _captured.Clear();
            }
            return _output;
        }

        /// <summary>
        /// Emit a typed inline at the current nesting level. Three mutually-exclusive cases, made
        /// visually exhaustive:
        /// <list type="bullet">
        /// <item>Not in a field (<c>_fieldDepth == 0</c>): commit to the top-level output.</item>
        /// <item>In a field's result phase (after <c>separate</c>): divert into the field result.</item>
        /// <item>In a field's instruction phase (before <c>separate</c>): field plumbing — flatten
        /// it from the modeled stream, but ALSO capture its <paramref name="source"/> element so the
        /// unterminated-field fallback in <see cref="Finish"/> can re-emit it opaquely rather than
        /// silently dropping it. A null source (e.g. a synthesized field run) cannot be captured;
        /// such inlines only ever emit at depth 0, so this never loses content.</item>
        /// </list>
        /// </summary>
        private void EmitInline(IrInline inline, XElement? source = null)
        {
            if (_fieldDepth == 0)
                _output.Add(inline);
            else if (_inResult)
                _result.Add(inline);
            else if (source is not null)
                _captured.Add(source);
        }

        /// <summary>
        /// Run an inline-SDT/smartTag splice, then flag every <see cref="IrTextRun"/> it appended to the
        /// active sink (<see cref="_output"/> outside a field, <see cref="_result"/> inside one's result)
        /// with <see cref="IrTextRun.FromInlineSdt"/>. The flag is equality-neutral so this preserves the
        /// content-transparency of content controls while letting the markdown emitter drop the spliced
        /// text (the oracle's GroupInlineRuns never visits a w:sdt). Nested splices compound naturally:
        /// an inner splice marks its runs, and the outer pass re-marking an already-true flag is a no-op.
        /// </summary>
        private void SpliceWithInlineSdtMark(Action splice)
        {
            int outBefore = _output.Count;
            int resBefore = _result.Count;
            splice();
            MarkInlineSdt(_output, outBefore);
            MarkInlineSdt(_result, resBefore);
        }

        private static void MarkInlineSdt(List<IrInline> sink, int from)
        {
            for (int i = from; i < sink.Count; i++)
                if (sink[i] is IrTextRun tr && !tr.FromInlineSdt)
                    sink[i] = tr with { FromInlineSdt = true };
        }
    }

    /// <summary>
    /// Map a single run child (text, tab, break, special hyphens, sym, comment plumbing) into
    /// <paramref name="sink"/>. Shared by the top-level walk and the field-result diversion so a
    /// field's cached result is read with identical run semantics.
    /// </summary>
    private static void EmitRunChild(XElement child, XElement? rPr, IrRunFormat runFormat, ReadContext ctx, List<IrInline> sink)
    {
        if (child.Name == W + "t")
            sink.Add(new IrTextRun(child.Value, runFormat));
        else if (child.Name == W + "tab")
            sink.Add(new IrTab(runFormat));
        else if (child.Name == W + "br")
            sink.Add(new IrBreak(BreakKind(child)));
        else if (child.Name == W + "noBreakHyphen")
            // N7: non-breaking hyphen → text U+2011, carrying the run format so it coalesces
            // with adjacent same-format text in the post-process pass.
            sink.Add(new IrTextRun("‑", runFormat));
        else if (child.Name == W + "softHyphen")
            // N7: soft hyphen → text U+00AD, same coalescing semantics as above.
            sink.Add(new IrTextRun("­", runFormat));
        else if (child.Name == W + "sym")
            AppendSym(child, rPr, runFormat, sink); // N8.
        else if (child.Name == W + "footnoteReference")
            sink.Add(new IrNoteRef(IrNoteKind.Footnote, NoteId(child)));
        else if (child.Name == W + "endnoteReference")
            sink.Add(new IrNoteRef(IrNoteKind.Endnote, NoteId(child)));
        else if (child.Name == W + "drawing")
            AppendDrawing(child, ctx, sink); // inline image (a:blip @r:embed) or opaque fallback.
        else if (child.Name == W + "lastRenderedPageBreak")
            return; // N4: layout cache, not content.
        else if (child.Name == W + "commentReference")
            // N15 (strip half): comment plumbing never affects ContentHash. Comment ids/targets are
            // recorded upstream by CommentTracker (via FeedRun); here the reference is only dropped.
            return;
        else if (IsDroppedParagraphChild(child.Name))
            return; // N3: bookmarks can legally appear inside a run too.
        else if (child.DescendantsAndSelf(WTxbxContent).Any())
            // A w:pict (VML v:textbox), an mc:AlternateContent wrapping a DrawingML w:drawing AND a VML
            // w:pict, or any other carrier of textbox bodies: model each w:txbxContent body as an
            // IrTextbox rather than dropping it opaquely (which is what closes the ContentHash blind
            // spot). A w:drawing carrying ONLY a textbox (no resolvable blip) reaches here too — it is
            // NOT promoted to an image by AppendDrawing (that branch is above), so it lands here.
            AppendTextboxes(child, ctx, sink);
        else
            sink.Add(new IrOpaqueInline(child.Name, IrHasher.CanonicalHash(child, ctx.RelResolver)));
    }

    // --- textbox bodies (w:txbxContent inside w:drawing/wps:txbx or w:pict/v:textbox) -------------

    /// <summary>
    /// Append one <see cref="IrTextbox"/> per OUTERMOST <c>w:txbxContent</c> reachable from
    /// <paramref name="carrier"/> (a run child — a <c>w:drawing</c>, a <c>w:pict</c>, or an
    /// <c>mc:AlternateContent</c> wrapping both), in document order. Each body's block children are
    /// walked by the normal block walker under the current scope, so every inner <c>w:p</c>/<c>w:tbl</c>
    /// gets its own anchor and hashes — matching the ORACLE, whose <c>Descendants(w:t)</c> /
    /// <c>DescendantsAndSelf</c> walks see textbox text and inner-paragraph anchors regardless of the
    /// DrawingML/VML wrapper. Word emits the same logical textbox twice (an <c>mc:Choice</c> DrawingML
    /// copy and an <c>mc:Fallback</c> VML copy); this yields one node per copy on purpose, so flat text
    /// and the anchor set match the oracle's both-copies traversal. Only the OUTERMOST bodies are taken
    /// here — a nested textbox (a <c>w:txbxContent</c> inside another) is reached when the block walker
    /// descends into the inner paragraph's runs, so it is modeled once at its true depth. Depth-capped:
    /// at <see cref="MaxTextboxDepth"/> the body is preserved as a single opaque inline instead of
    /// recursing further (totality against adversarial nesting).
    /// </summary>
    private static void AppendTextboxes(XElement carrier, ReadContext ctx, List<IrInline> sink)
    {
        foreach (var txbx in OutermostTextboxBodies(carrier))
        {
            if (ctx.TextboxDepth >= MaxTextboxDepth)
            {
                sink.Add(new IrOpaqueInline(txbx.Name, IrHasher.CanonicalHash(txbx, ctx.RelResolver)));
                continue;
            }

            var innerCtx = ctx.IntoTextbox();
            var blocks = new List<IrBlock>();
            foreach (var child in txbx.Elements())
                AppendBlocks(child, innerCtx, blocks);
            sink.Add(new IrTextbox(IrNodeList.From(blocks)));
        }
    }

    /// <summary>
    /// The <c>w:txbxContent</c> elements reachable from <paramref name="root"/> that are NOT themselves
    /// nested inside another <c>w:txbxContent</c> — i.e. the outermost bodies. Nested bodies are
    /// deliberately excluded here because the recursive block walker re-discovers them when it descends
    /// into an outer body's inner paragraphs (so each is modeled exactly once at its real nesting depth).
    /// </summary>
    private static IEnumerable<XElement> OutermostTextboxBodies(XElement root)
    {
        // Take a w:txbxContent only when no ANCESTOR (walking up but stopping at the carrier root) is
        // itself a w:txbxContent — that is the outermost-within-carrier test. DescendantsAndSelf so a
        // carrier that IS a w:txbxContent (it never is today, but stays correct) is handled too.
        foreach (var txbx in root.DescendantsAndSelf(WTxbxContent))
        {
            var nestedUnderAnotherBody = false;
            foreach (var anc in txbx.Ancestors())
            {
                if (ReferenceEquals(anc, root)) break;
                if (anc.Name == WTxbxContent) { nestedUnderAnotherBody = true; break; }
            }
            if (!nestedUnderAnotherBody)
                yield return txbx;
        }
    }

    // --- hyperlinks (N14) -------------------------------------------------

    /// <summary>
    /// N14: promote a <c>w:hyperlink</c> to an <see cref="IrHyperlink"/>. Child <c>w:r</c> content
    /// is walked through the SAME inline walker as direct paragraph runs (so empty-drop + N5
    /// coalescing apply within the link's own inline list). Target resolution: an <c>@r:id</c>
    /// resolves against the main part's hyperlink relationships to the external URI (a missing
    /// relationship tolerates to <c>Target = null</c>); an <c>@w:anchor</c> internal link uses the
    /// convention <c>Target = "#" + anchor</c>. <see cref="IrHyperlink.InternalTarget"/> stays null
    /// in M1.2 — resolving the bookmark→anchor mapping is future work (TODO(M1.3+)).
    /// </summary>
    private static IrHyperlink BuildHyperlink(XElement hyperlink, ReadContext ctx)
    {
        var inlines = WalkInlines(hyperlink.Elements(), ctx);

        string? target = null;
        var relId = (string?)hyperlink.Attribute(R + "id");
        if (relId is not null)
        {
            target = ResolveHyperlinkRel(hyperlink, relId); // null when the relationship is missing.
        }
        else
        {
            var anchor = (string?)hyperlink.Attribute(W + "anchor");
            if (anchor is not null)
                target = "#" + anchor;
        }

        // TODO(M1.3+): resolve @w:anchor to the target block's IrAnchor and set InternalTarget.
        return new IrHyperlink(target, InternalTarget: null, IrNodeList.From(inlines));
    }

    /// <summary>
    /// Resolve a hyperlink <c>@r:id</c> to its external URI via the owning part's hyperlink
    /// relationships, mirroring <c>WmlToMarkdownConverter.LookupRelationshipUrl</c>: the part is
    /// stashed on the document root as an <see cref="OpenXmlPart"/> annotation by
    /// <see cref="Read"/>. Returns null when the part or relationship is absent (missing-rel
    /// tolerance — a dangling r:id must not throw).
    /// </summary>
    private static string? ResolveHyperlinkRel(XElement el, string relId)
    {
        var root = el.AncestorsAndSelf().Last();
        var part = root.Annotation<OpenXmlPart>();
        if (part is null)
            return null;
        foreach (var rel in part.HyperlinkRelationships)
            if (rel.Id == relId)
                return rel.Uri.ToString();
        return null;
    }

    // --- note refs --------------------------------------------------------

    /// <summary>
    /// The <c>@w:id</c> of a note reference, or <c>""</c> when absent. The id is intentionally NOT
    /// fed into the content hash (spec §6.1: only the kind sentinel is hashed) — it is positional
    /// bookkeeping, and renumbering notes must not flip body hashes. Separator/continuationSeparator
    /// reference variants carry no id (or carry only <c>@w:customMarkFollows</c> etc.), which yields
    /// <c>""</c> here without crashing.
    /// </summary>
    private static string NoteId(XElement noteRef) => (string?)noteRef.Attribute(W + "id") ?? "";

    // --- inline images ----------------------------------------------------

    /// <summary>
    /// Promote a <c>w:drawing</c> whose descendant <c>a:blip</c> has an <c>@r:embed</c> resolving to
    /// an image part on the main document part into an <see cref="IrInlineImage"/>. Extent comes from
    /// the first descendant <c>wp:extent</c> (<c>@cx</c>/<c>@cy</c>, 0 when absent); alt text from
    /// <c>wp:docPr/@descr</c> falling back to <c>@name</c>, else null. The image part's bytes are
    /// read fully and SHA-256'd (cached per embed rel id within one <see cref="Read"/> so a reused
    /// logo hashes once). Anything that doesn't resolve to a real image part — a missing/wrong-typed
    /// relationship, a <c>w:pict</c> (VML) with no blip, or a drawing without <c>a:blip@embed</c> —
    /// falls back to <see cref="IrOpaqueInline"/>; this method never throws on malformed drawings.
    /// </summary>
    private static void AppendDrawing(XElement drawing, ReadContext ctx, List<IrInline> sink)
    {
        // Image promotion and textbox modeling are INDEPENDENT, matching the oracle: a w:drawing can
        // carry BOTH a resolvable a:blip image AND a wps:txbx textbox body, and the oracle's
        // Descendants(w:t)/DescendantsAndSelf walks see the textbox text/anchors regardless of the
        // blip. So we promote the image if one resolves, AND (always) append a textbox per
        // w:txbxContent body. A drawing with only a textbox (no resolvable blip) yields just the
        // textbox; a drawing with neither falls back to a single opaque inline.
        var blip = drawing.Descendants(ABlip).FirstOrDefault();
        var embedId = (string?)blip?.Attribute(REmbed);
        var promotedImage = false;
        if (embedId is not null)
        {
            var resolved = ResolveImagePart(ctx, embedId);
            if (resolved is { } image)
            {
                var extent = drawing.Descendants(WpExtent).FirstOrDefault();
                long cx = LongAttr(extent, "cx");
                long cy = LongAttr(extent, "cy");

                var docPr = drawing.Descendants(WpDocPr).FirstOrDefault();
                string? altText = (string?)docPr?.Attribute("descr")
                    ?? (string?)docPr?.Attribute("name");

                // The w:drawing's pt:Unid is the IR's img-anchor identity (M1.4-T2). Deterministic
                // Unids were assigned over the whole part before the walk, so this is present and
                // stable for any drawing in the document. Equality-neutral on IrInlineImage.
                var drawingUnid = (string?)drawing.Attribute(PtOpenXml.Unid);

                sink.Add(new IrInlineImage(image.PartUri, image.BytesHash, cx, cy, altText)
                {
                    Unid = drawingUnid,
                });
                promotedImage = true;
            }
        }

        // Textbox bodies (if any) are modeled regardless of whether an image was promoted.
        var hasTextbox = drawing.DescendantsAndSelf(WTxbxContent).Any();
        if (hasTextbox)
        {
            AppendTextboxes(drawing, ctx, sink);
            return;
        }

        if (promotedImage)
            return;

        // No resolvable embedded image and no textbox (missing rel, unmodeled shape, etc.): preserve
        // the whole drawing opaquely so nothing is lost (totality).
        sink.Add(new IrOpaqueInline(drawing.Name, IrHasher.CanonicalHash(drawing, ctx.RelResolver)));
    }

    /// <summary>
    /// Resolve an embed rel id to its image part's URI and the SHA-256 of its bytes, caching per rel
    /// id within the current read (so a logo reused N times hashes once). Returns null when the
    /// relationship is missing or does not point at an image part, or when reading the part's stream
    /// fails — every such case is tolerated to an opaque fallback by the caller so a malformed
    /// package never aborts the read.
    /// </summary>
    private static (Uri PartUri, IrHash BytesHash)? ResolveImagePart(ReadContext ctx, string embedId)
    {
        if (ctx.ImageHashCache.TryGetValue(embedId, out var cached))
            return cached;

        try
        {
            var part = ctx.Main.GetPartById(embedId);
            if (part is not ImagePart imagePart)
                return null;

            using var stream = imagePart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read);
            using var ms = new System.IO.MemoryStream();
            stream.CopyTo(ms);
            var hash = IrHash.Compute(ms.GetBuffer().AsSpan(0, (int)ms.Length));
            var resolved = (imagePart.Uri, hash);
            ctx.ImageHashCache[embedId] = resolved;
            return resolved;
        }
        catch (Exception e) when (e is KeyNotFoundException or ArgumentException or System.IO.IOException
            or InvalidOperationException or NotSupportedException or System.Xml.XmlException)
        {
            // Missing rel id (KeyNotFoundException / ArgumentOutOfRangeException), wrong part type,
            // or unreadable stream: opaque fallback. Totality over the corpus depends on this.
            // OOM and other systemic exceptions are allowed to escape rather than masquerade as a
            // benign missing image.
            return null;
        }
    }

    private static long LongAttr(XElement? el, string localName)
    {
        var raw = (string?)el?.Attribute(localName);
        return raw is not null
            && long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v)
            ? v
            : 0L;
    }

    // --- fields (N9) ------------------------------------------------------

    /// <summary>
    /// N9: promote a <c>w:fldSimple</c> to an <see cref="IrFieldRun"/>. The <c>@w:instr</c> is the
    /// instruction string; the child <c>w:r</c> content is walked normally into the cached result.
    /// </summary>
    private static IrFieldRun BuildFldSimple(XElement fldSimple, ReadContext ctx)
    {
        var instruction = (string?)fldSimple.Attribute(W + "instr") ?? "";
        var result = WalkInlines(fldSimple.Elements(), ctx);
        return new IrFieldRun(instruction, IrNodeList.From(result)) { IsSimpleField = true };
    }

    /// <summary>
    /// N8: map <c>w:sym</c> to text. When <c>@w:char</c> parses as a BMP hex code point, emit the
    /// single character as an <see cref="IrTextRun"/>. The symbol's <c>@w:font</c> is glyph-bearing
    /// formatting, so the whole <c>w:sym</c> element is folded into a per-character run format
    /// (cloned into the unmodeled-digest container) — it must influence the FORMAT fingerprint, not
    /// just vanish. If <c>@w:char</c> is missing or unparseable, fall back to opaque. C0 control
    /// code points (&lt; U+0020, including U+0000) are also rejected to opaque: they cannot appear
    /// as legal XML text content and would collide with the content-hash sentinel lead bytes.
    /// </summary>
    private static void AppendSym(XElement sym, XElement? rPr, IrRunFormat baseFormat, List<IrInline> inlines)
    {
        var charAttr = (string?)sym.Attribute(W + "char");
        if (charAttr is not null
            && int.TryParse(charAttr, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var code)
            && code is >= 0x20 and <= 0xFFFF)
        {
            // BMP code point (Word's symbol convention uses the F000-F0FF private-use range): emit
            // the codepoint as a single char verbatim — no surrogate handling needed.
            var text = ((char)code).ToString();
            var symFormat = MapRunFormat(rPr, extraUnmodeled: sym);
            inlines.Add(new IrTextRun(text, symFormat));
            return;
        }

        inlines.Add(new IrOpaqueInline(sym.Name, IrHasher.CanonicalHash(sym)));
    }

    /// <summary>
    /// Paragraph/run children that rules N3 (bookmarks) and N15 (comment range plumbing) drop from
    /// the inline stream entirely. Recoverable via block provenance; never affect any hash.
    /// </summary>
    private static bool IsDroppedParagraphChild(XName name) =>
        name == W + "bookmarkStart"
        || name == W + "bookmarkEnd"
        || name == W + "commentRangeStart"   // N15: target spans are recorded by CommentTracker upstream; here only dropped.
        || name == W + "commentRangeEnd";

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
                && prev.Format.Equals(run.Format)
                // Don't coalesce across the inline-SDT boundary: the flag is the emitter's drop signal,
                // so a spliced run must not absorb (or be absorbed into) an unflagged neighbor.
                && prev.FromInlineSdt == run.FromInlineSdt)
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

    /// <summary>
    /// The run format carried by each inline that has one, in inline order. Recurses into
    /// <see cref="IrHyperlink.Inlines"/> and <see cref="IrFieldRun.CachedResult"/> so a hyperlink's
    /// or field-result's run formats participate in the paragraph's run-format sequence in place
    /// (a bolded link word flips the block fingerprint exactly as a bolded plain word does).
    /// </summary>
    private static IEnumerable<IrRunFormat> RunFormatsInOrder(IEnumerable<IrInline> inlines)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrTextRun r: yield return r.Format; break;
                case IrTab t: yield return t.Format; break;
                case IrHyperlink h:
                    foreach (var f in RunFormatsInOrder(h.Inlines)) yield return f;
                    break;
                case IrFieldRun fld:
                    foreach (var f in RunFormatsInOrder(fld.CachedResult)) yield return f;
                    break;
            }
        }
    }

    private static IrHash ComputeParagraphContentHash(IEnumerable<IrInline> inlines)
    {
        var builder = new IrContentHashBuilder();
        AppendInlinesToContentHash(inlines, builder);
        return builder.Build();
    }

    /// <summary>
    /// Append the canonical content-hash byte stream of <paramref name="inlines"/> into
    /// <paramref name="builder"/> (spec §6.1). Recursive so nested inlines (hyperlink children,
    /// field cached results) stream through the SAME per-inline dispatch as the top level.
    /// Semantics worth noting:
    /// <list type="bullet">
    /// <item><see cref="IrFieldRun"/> contributes ONLY its cached-result inlines' bytes —
    /// transparently, with no sentinels and no instruction bytes — so a PAGE field showing "5" is
    /// content-equal to a literal "5" (the hash captures what a reader sees; the instruction is
    /// consumer-visible but unhashed).</item>
    /// <item><see cref="IrHyperlink"/> is bracketed: sentinel <c>0x08</c>, the target bytes,
    /// sentinel <c>0x09</c>, then its child inlines' bytes — so a target change is a content change
    /// and linked text is never content-equal to identical plain text.</item>
    /// </list>
    /// </summary>
    private static void AppendInlinesToContentHash(IEnumerable<IrInline> inlines, IrContentHashBuilder builder)
    {
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
                case IrHyperlink h:
                    builder.AppendSentinel(IrContentHashBuilder.SentinelHyperlink);
                    builder.AppendText(h.Target ?? "");
                    builder.AppendSentinel(IrContentHashBuilder.SentinelHyperlinkTargetEnd);
                    AppendInlinesToContentHash(h.Inlines, builder);
                    break;
                case IrFieldRun fld:
                    // Transparent: cached-result bytes only, no sentinels, no instruction.
                    AppendInlinesToContentHash(fld.CachedResult, builder);
                    break;
                case IrNoteRef note:
                    // Kind sentinel ONLY (0x05/0x06) — never the note id (spec §6.1). Footnote vs
                    // endnote stays distinguishable; renumbering does not change the body hash.
                    builder.AppendSentinel(note.Kind == IrNoteKind.Footnote
                        ? IrContentHashBuilder.SentinelFootnoteRef
                        : IrContentHashBuilder.SentinelEndnoteRef);
                    break;
                case IrInlineImage img:
                    // Sentinel + the image part's bytes hash (spec §6.1). Extent and alt text do NOT
                    // affect ContentHash (they are not text); they also do not currently affect the
                    // FormatFingerprint — the image carries no run format, so a resize is invisible to
                    // both hashes today.
                    // TODO(M2): the diff engine may want extent/alt changes surfaced (e.g. as a
                    // format-grade change) so a resized-but-same-bytes image reads as "changed".
                    builder.AppendSentinel(IrContentHashBuilder.SentinelImage);
                    builder.AppendHash(img.ImageBytesHash);
                    break;
                case IrTextbox tb:
                    // Sentinel + each inner block's ContentHash in order (spec §6.1 M1.5 addendum).
                    // This is what makes textbox text participate in the containing paragraph's
                    // ContentHash — a textbox text edit now flips the paragraph hash — while the
                    // sentinel framing keeps textbox text distinct from identical inline text. Inner
                    // blocks' formatting stays in THEIR OWN fingerprints; it does NOT enter the
                    // containing paragraph's FormatFingerprint (the diff engine sees inner blocks as
                    // separately indexed blocks, so folding their formats here would double-count).
                    builder.AppendSentinel(IrContentHashBuilder.SentinelTextbox);
                    foreach (var inner in tb.Blocks)
                        builder.AppendHash(inner.ContentHash);
                    break;
                case IrOpaqueInline o:
                    builder.AppendSentinel(IrContentHashBuilder.SentinelOpaque);
                    builder.AppendHash(o.CanonicalHash);
                    break;
            }
        }
    }

    // --- paragraph format -------------------------------------------------

    /// <summary>
    /// Map a <c>w:pPr</c> element into the modeled <see cref="IrParaFormat"/> subset (§5.1). This is
    /// ctx-free and list-resolution-free so the effective-format resolver
    /// (<see cref="IrEffectiveFormats"/>) can reuse it on a style's or docDefaults' cloned
    /// <c>w:pPr</c> exactly as the body walk does on a paragraph's direct <c>w:pPr</c> — there is one
    /// pPr→IrParaFormat mapping, never a duplicate. List membership (which needs the numbering
    /// registry) is resolved separately by <see cref="ResolveListInfo"/>. Returns the empty-digest
    /// format when <paramref name="pPr"/> is null.
    /// </summary>
    internal static IrParaFormat MapParaFormat(XElement? pPr)
    {
        if (pPr is null)
            return new IrParaFormat { UnmodeledDigest = EmptyUnmodeledDigest };

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

        // Unmodeled leftovers: every pPr child not consumed by a modeled field above.
        // numPr is read for the IrListInfo facts but is intentionally NOT in PPrConsumed, so it
        // stays in this digest: IrListInfo is not hashed, and numPr has always ridden in the
        // unmodeled digest — keeping it here makes FormatFingerprint byte-identical across M1.3
        // resolution. w:rPr (mark props) and mid-doc w:sectPr stay too.
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
        return format;
    }

    // --- list resolution (M1.3) -------------------------------------------

    /// <summary>
    /// Resolve a paragraph's list membership through the numbering registry, mirroring
    /// <c>BlockMetadataOps.ResolveListMembership</c> / <c>WmlToMarkdownConverter.IsListItem</c>:
    /// <list type="number">
    /// <item>A direct <c>w:pPr/w:numPr</c> wins (<c>FromStyle=false</c>).</item>
    /// <item>Otherwise walk the <c>pStyle → basedOn</c> chain (cycle-guarded, depth ≤ 16) for the
    /// first style that contributes a <c>w:numPr</c> (<c>FromStyle=true</c>).</item>
    /// </list>
    /// OOXML rule: a <c>w:numId</c> of 0 (or absent, or one that does not resolve to a
    /// <see cref="IrNum"/> in the registry) means "no list membership" → returns null. When the
    /// numId resolves, the abstract-num id, the level's <c>w:numFmt</c> string (or "" when the
    /// level is missing), and the per-level <c>w:startOverride</c> (or null) come from the registry.
    /// </summary>
    private static IrListInfo? ResolveListInfo(XElement pPr, ReadContext ctx)
    {
        var numPr = pPr.Element(W + "numPr");
        bool fromStyle = false;

        if (numPr is null)
        {
            var styleId = AttrVal(pPr.Element(W + "pStyle"));
            numPr = ResolveStyleNumPr(styleId, ctx.Styles);
            fromStyle = numPr is not null;
        }

        if (numPr is null)
            return null;

        var numId = IntAttr(numPr.Element(W + "numId"), W + "val");
        // numId absent or 0 → no list membership (OOXML: numId="0" cancels numbering).
        if (numId is null || numId.Value == 0)
            return null;

        var ilvl = IntAttr(numPr.Element(W + "ilvl"), W + "val") ?? 0;

        // numId must resolve to a w:num in the registry; an unresolvable numId is no membership.
        if (!ctx.Numbering.Nums.TryGetValue(numId.Value, out var num))
            return null;

        int abstractNumId = num.AbstractNumId;

        // Level format from the abstract num's level table; "" when the level is missing.
        string numberFormat = "";
        if (ctx.Numbering.AbstractNums.TryGetValue(abstractNumId, out var abs)
            && abs.Levels.TryGetValue(ilvl, out var level))
        {
            numberFormat = level.NumberFormat;
        }

        int? startOverride = num.StartOverrides.TryGetValue(ilvl, out var so) ? so : (int?)null;

        return new IrListInfo(numId.Value, abstractNumId, ilvl, numberFormat, startOverride, fromStyle);
    }

    /// <summary>
    /// Walk the <c>pStyle → basedOn</c> chain for the first style that carries a <c>w:numPr</c>,
    /// reading from the deep-cloned <see cref="IrStyle.PPr"/> in the style registry. Cycle-guarded
    /// and depth-capped at 16, matching <c>WmlToMarkdownConverter.IsListItem</c> and
    /// <c>BlockMetadataOps.ResolveStyleNumPr</c>. Returns null when no style in the chain
    /// contributes numbering (or the chain is broken/cyclic).
    /// </summary>
    private static XElement? ResolveStyleNumPr(string? styleId, IrStyleRegistry styles)
    {
        if (string.IsNullOrEmpty(styleId))
            return null;

        var visited = new HashSet<string>(StringComparer.Ordinal);
        var current = styleId;
        for (int i = 0; i < 16 && current is not null; i++)
        {
            if (!visited.Add(current))
                return null; // cycle guard.
            if (!styles.Styles.TryGetValue(current, out var style))
                return null;
            var numPr = style.PPr?.Element(W + "numPr");
            if (numPr is not null)
                return numPr;
            current = style.BasedOn;
        }
        return null;
    }

    // --- run format -------------------------------------------------------

    /// <summary>
    /// Map a <c>w:rPr</c> element into the modeled <see cref="IrRunFormat"/> subset (§5.1). Ctx-free
    /// so <see cref="IrEffectiveFormats"/> can reuse it on a style's / docDefaults' cloned
    /// <c>w:rPr</c> — one rPr→IrRunFormat mapping, no duplication. The optional
    /// <paramref name="extraUnmodeled"/> is a glyph-bearing run child (a <c>w:sym</c>) folded into
    /// the digest; effective resolution never passes it.
    /// </summary>
    internal static IrRunFormat MapRunFormat(XElement? rPr, XElement? extraUnmodeled = null)
    {
        if (rPr is null && extraUnmodeled is null)
            return new IrRunFormat { UnmodeledDigest = EmptyUnmodeledDigest };

        if (rPr is null)
            // No run props, but a glyph-bearing extra (a w:sym): the digest is just that element.
            return new IrRunFormat { UnmodeledDigest = ExtraUnmodeledDigest(extraUnmodeled!) };

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
        var digest = UnmodeledDigest(rPr, RPrConsumed, vertAlign is not null ? W + "vertAlign" : null,
            extraUnmodeled);

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

        // Rows are usually direct w:tr children, but a table-level w:sdt (e.g. a repeating-section
        // content control) can wrap a w:tr. The oracle's table markdown (tbl.Elements(w:tr)) drops
        // those rows, but its anchor index (Descendants) keeps them — and the IR must not lose the
        // row's content. Unwrap table-level SDTs to their w:tr children, flag the delivered rows
        // FromTableSdt, and let the emitter narrow back to the oracle's direct-row view for parity.
        foreach (var (tr, fromSdt) in EnumerateTableRows(tbl))
        {
            var (row, cellFingerprintsForRow) = BuildRow(tr, ctx);
            if (fromSdt)
                row = row with { FromTableSdt = true };
            rows.Add(row);
            rowHashes.Add(row.ContentHash);
            cellFingerprints.AddRange(cellFingerprintsForRow);
        }

        // Non-tr children of the table (tblPr, tblGrid) + non-tc children of each tr (trPr) fold into
        // one unmodeled container so any table-level prop change flips the fingerprint. A row-delivering
        // w:sdt wrapper is excluded — its meaningful content (the rows) is now modeled above, so folding
        // the wrapper too would double-count and re-introduce the very content the rows path captures.
        var unmodeledContainer = new XElement("unmodeled");
        foreach (var child in tbl.Elements().Where(e => e.Name != W + "tr" && e.Name != W + "sdt"))
            unmodeledContainer.Add(new XElement(child));
        foreach (var (tr, _) in EnumerateTableRows(tbl))
            foreach (var child in tr.Elements().Where(e => e.Name != W + "tc" && e.Name != W + "sdt"))
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
            Anchor = AnchorFor(IrAnchorKind.Tbl, tbl, ctx),
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

        // A row's children are usually direct w:tc elements, but a w:sdt (content control) can wrap
        // a w:tc as a row-level child — Word does this for repeating-section content controls. The
        // oracle's TABLE path (Elements(w:tr).Elements(w:tc)) is blind to those cells and drops them
        // from its view; the IR must NOT lose content (the Phase 2 diff engine needs every cell in
        // the ContentHash), so we unwrap row-level SDTs to their w:tc children here, reusing the
        // established SDT-unwrap discipline + depth cap. Each SDT-delivered cell is flagged
        // FromRowSdt so the markdown emitter can mirror the oracle's narrower view for byte parity.
        foreach (var (tc, fromSdt) in EnumerateRowCells(tr))
        {
            var (cell, fingerprints) = BuildCell(tc, ctx);
            if (fromSdt)
                cell = cell with { FromRowSdt = true };
            cells.Add(cell);
            cellFingerprints.AddRange(fingerprints);
            rowBuilder.AppendHash(cell.ContentHash);
        }

        var row = new IrRow(AnchorFor(IrAnchorKind.Tr, tr, ctx), IrNodeList.From(cells), rowBuilder.Build())
        {
            Source = ctx.Provenance(tr),
        };
        return (row, cellFingerprints);
    }

    /// <summary>
    /// Yield the rows of a table in document order, unwrapping any table-level <c>w:sdt</c> to its
    /// <c>w:sdtContent</c>'s <c>w:tr</c> children (recursively, depth-capped). A direct <c>w:tr</c>
    /// yields <c>fromSdt=false</c>; an SDT-delivered row yields <c>fromSdt=true</c>. Non-tr, non-sdt
    /// table children (<c>w:tblPr</c>/<c>w:tblGrid</c>) are skipped here (folded into the unmodeled
    /// digest by <see cref="BuildTable"/>).
    /// </summary>
    private static IEnumerable<(XElement Tr, bool FromSdt)> EnumerateTableRows(XElement tbl)
    {
        foreach (var child in tbl.Elements())
        {
            if (child.Name == W + "tr")
                yield return (child, false);
            else if (child.Name == W + "sdt")
                foreach (var tr in UnwrapSdtChildren(child, W + "tr", depth: 0))
                    yield return (tr, true);
        }
    }

    /// <summary>
    /// Yield the cells of a row in document order, unwrapping any row-level <c>w:sdt</c> to its
    /// <c>w:sdtContent</c>'s <c>w:tc</c> children (recursively, depth-capped by <see cref="MaxSdtDepth"/>).
    /// A direct <c>w:tc</c> yields <c>fromSdt=false</c>; a cell delivered by an SDT yields
    /// <c>fromSdt=true</c>. Non-tc, non-sdt row children (e.g. <c>w:trPr</c>) are skipped — they are
    /// folded into the table's unmodeled-props digest by <see cref="BuildTable"/>. A pathologically
    /// deep SDT nest stops unwrapping at the cap (any deeper <c>w:tc</c> is simply not surfaced — the
    /// same bound the block walker uses), preserving totality.
    /// </summary>
    private static IEnumerable<(XElement Tc, bool FromSdt)> EnumerateRowCells(XElement tr)
    {
        foreach (var child in tr.Elements())
        {
            if (child.Name == W + "tc")
                yield return (child, false);
            else if (child.Name == W + "sdt")
                foreach (var tc in UnwrapSdtChildren(child, W + "tc", depth: 0))
                    yield return (tc, true);
        }
    }

    /// <summary>
    /// Yield the <paramref name="targetName"/> elements (<c>w:tr</c> or <c>w:tc</c>) reachable through a
    /// <c>w:sdt</c>'s <c>w:sdtContent</c>, recursing through nested SDTs up to <see cref="MaxSdtDepth"/>.
    /// Beyond the cap, deeper targets are simply not surfaced (the same bound the block walker uses),
    /// preserving totality against adversarial nesting.
    /// </summary>
    private static IEnumerable<XElement> UnwrapSdtChildren(XElement sdt, XName targetName, int depth)
    {
        if (depth >= MaxSdtDepth)
            yield break;
        var content = sdt.Element(W + "sdtContent");
        if (content is null)
            yield break;
        foreach (var inner in content.Elements())
        {
            if (inner.Name == targetName)
                yield return inner;
            else if (inner.Name == W + "sdt")
                foreach (var t in UnwrapSdtChildren(inner, targetName, depth + 1))
                    yield return t;
        }
    }

    private static (IrCell Cell, List<IrHash> Fingerprints) BuildCell(XElement tc, ReadContext ctx)
    {
        var tcPr = tc.Element(W + "tcPr");
        int gridSpan = IntAttr(tcPr?.Element(W + "gridSpan"), W + "val") ?? 1;
        var vMerge = MapVMerge(tcPr?.Element(W + "vMerge"));

        // Cell-SHELL digest: the canonical hash of the whole w:tcPr (width, gridSpan, vMerge, borders,
        // shading, …). default(IrHash) = no tcPr. Folded into the cell's ContentHash below so a
        // shell-only edit (cell width / merge change) is VISIBLE to the diff engine — without it such an
        // edit left every cell/row/table hash identical, the table pair classified EqualBlock, and the
        // edit silently vanished from both Compare and Consolidate with zero conflict recorded (a
        // soundness bug). Also stored on the IrCell so the N-way merger can attribute/conflict competing
        // shell edits without re-resolving source elements.
        var shellDigest = tcPr != null ? IrHasher.CanonicalHash(tcPr) : default;

        var blocks = new List<IrBlock>();
        var fingerprints = new List<IrHash>();
        var cellBuilder = new IrContentHashBuilder();
        cellBuilder.AppendStructure(IrContentHashBuilder.StructureCell);
        if (tcPr != null)
            cellBuilder.AppendHash(shellDigest);

        foreach (var child in tc.Elements())
        {
            if (child.Name == W + "tcPr")
                continue;
            // N12: a block-level w:sdt inside a cell unwraps to its content blocks, same as in body.
            var before = blocks.Count;
            AppendBlocks(child, ctx, blocks);
            for (int i = before; i < blocks.Count; i++)
            {
                cellBuilder.AppendHash(blocks[i].ContentHash);
                fingerprints.Add(blocks[i].FormatFingerprint);
            }
        }

        var cell = new IrCell(
            AnchorFor(IrAnchorKind.Tc, tc, ctx),
            IrNodeList.From(blocks),
            gridSpan,
            vMerge,
            cellBuilder.Build())
        {
            Source = ctx.Provenance(tc),
            ShellDigest = shellDigest,
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

        // ContentHash: a single opaque hash of the whole sectPr — deterministic and simple. Resolver-aware
        // so a header/footer-reference rel-id renumber (r:id pointing at content-identical header/footer
        // parts) does not perturb a section break's identity.
        var contentBuilder = new IrContentHashBuilder();
        contentBuilder.AppendHash(IrHasher.CanonicalHash(sectPr, ctx.RelResolver));

        return new IrSectionBreak
        {
            Anchor = AnchorFor(IrAnchorKind.Sec, sectPr, ctx),
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
            Anchor = AnchorFor(IrAnchorKind.Unk, el, ctx),
            ElementName = el.Name,
            ContentHash = IrHasher.CanonicalHash(el, ctx.RelResolver),
            FormatFingerprint = EmptyUnmodeledDigest,
            Source = ctx.Provenance(el),
        };

    // --- header / footer / note / comment scopes (rule M1.3) --------------

    /// <summary>
    /// Prepare a non-body part's root for the IR walk exactly as the body and the markdown projection
    /// do: assign deterministic Unids over the whole subtree and stash the owning part as an
    /// <see cref="OpenXmlPart"/> annotation on the root (so <c>KindFor → IsListItem</c> can reach the
    /// main part's <c>StyleDefinitionsPart</c> through the part's package). Idempotent on the
    /// annotation. Returns the part's root, or null when the part has no readable root (tolerated).
    /// </summary>
    private static XElement? PreparePartRoot(OpenXmlPart part)
    {
        var root = part.GetXDocument().Root;
        if (root is null)
            return null;
        UnidHelper.AssignToAllElementsDeterministic(root);
        if (root.Annotation<OpenXmlPart>() == null)
            root.AddAnnotation(part);
        return root;
    }

    /// <summary>
    /// Read the header (or footer) parts into <see cref="IrHeaderFooter"/> records, scope-named
    /// <c>hdr1</c>/<c>hdr2</c>… (or <c>ftr1</c>…) in the SAME enumeration order as
    /// <c>WmlToMarkdownConverter.BuildAnchorIndex</c> (<c>main.HeaderParts</c>/<c>FooterParts</c>).
    /// Each part's root is prepared (Unids + annotation), its element children are walked by the block
    /// walker under the scope name, its <see cref="XDocument"/> joins <paramref name="sources"/>, and
    /// its blocks are indexed into <paramref name="anchorIndex"/>. The occurrence
    /// <see cref="IrHeaderFooterKind"/> is resolved from the body section properties' references
    /// (<paramref name="referenceName"/> = <c>w:headerReference</c>/<c>w:footerReference</c>) whose
    /// <c>@r:id</c> points at the part's relationship id; an unreferenced part defaults to
    /// <see cref="IrHeaderFooterKind.Default"/>. Whole-part walks are tolerance-wrapped (a malformed
    /// part is skipped, not fatal) so corpus totality holds across all scopes.
    /// </summary>
    private static IrNodeList<IrHeaderFooter> ReadHeaderFooterScopes(
        MainDocumentPart main, XElement body, IrStyleRegistry styles, IrNumberingRegistry numbering,
        Dictionary<Uri, XDocument> sources, Dictionary<string, IrBlock> anchorIndex,
        IEnumerable<OpenXmlPart> parts, string scopePrefix, XName referenceName, bool retain)
    {
        var result = new List<IrHeaderFooter>();
        int i = 1;
        foreach (var part in parts)
        {
            var scopeName = $"{scopePrefix}{i++}";
            try
            {
                var root = PreparePartRoot(part);
                if (root is null)
                    continue;

                var ctx = new ReadContext(part.Uri, main, styles, numbering, scopeName,
                    retainSources: retain, owningPart: part);
                var blocks = new List<IrBlock>();
                foreach (var child in root.Elements())
                    AppendBlocks(child, ctx, blocks);

                foreach (var b in blocks)
                    IndexBlock(b, anchorIndex);
                if (retain)
                    sources[part.Uri] = part.GetXDocument();

                var kind = ResolveHeaderFooterKind(main, body, part, referenceName);
                result.Add(new IrHeaderFooter(scopeName, kind,
                    new IrScope(scopeName, IrNodeList.From(blocks), part.Uri)));
            }
            catch (Exception e) when (IsTolerableRegistryException(e))
            {
                // A malformed header/footer part is skipped; totality over the corpus depends on it.
            }
        }
        return IrNodeList.From(result);
    }

    /// <summary>
    /// Resolve a header/footer part's occurrence kind from the body section properties. The first
    /// <paramref name="referenceName"/> reference (across all body <c>w:sectPr</c>) whose <c>@r:id</c>
    /// resolves to <paramref name="part"/> wins: <c>@w:type</c> <c>first</c>/<c>even</c> map to the
    /// matching kind, everything else (including <c>default</c> and absent) to
    /// <see cref="IrHeaderFooterKind.Default"/>. A part no section references → Default.
    /// </summary>
    private static IrHeaderFooterKind ResolveHeaderFooterKind(
        MainDocumentPart main, XElement body, OpenXmlPart part, XName referenceName)
    {
        string? relId = null;
        try { relId = main.GetIdOfPart(part); }
        catch (Exception e) when (IsTolerableRegistryException(e)) { return IrHeaderFooterKind.Default; }

        foreach (var sectPr in body.Descendants(W + "sectPr"))
            foreach (var reference in sectPr.Elements(referenceName))
                if ((string?)reference.Attribute(R + "id") == relId)
                    return (string?)reference.Attribute(W + "type") switch
                    {
                        "first" => IrHeaderFooterKind.First,
                        "even" => IrHeaderFooterKind.Even,
                        _ => IrHeaderFooterKind.Default,
                    };
        return IrHeaderFooterKind.Default;
    }

    /// <summary>
    /// Read a footnotes (or endnotes) part into an <see cref="IrNoteStore"/> keyed by note id
    /// (<c>@w:id</c>). Word-reserved separator/continuation notes are skipped exactly as the projection
    /// skips them (<see cref="WmlToMarkdownConverter.IsBoilerplateNote"/>). Each real note's block
    /// children are walked under the scope name (<c>fn</c>/<c>en</c>), indexed, and its part's
    /// <see cref="XDocument"/> joins <paramref name="sources"/>. A missing part → empty store; the
    /// whole walk is tolerance-wrapped.
    /// </summary>
    private static IrNoteStore ReadNoteStore(
        MainDocumentPart main, OpenXmlPart? part, IrStyleRegistry styles, IrNumberingRegistry numbering,
        Dictionary<Uri, XDocument> sources, Dictionary<string, IrBlock> anchorIndex,
        string scopeName, XName noteName, bool retain)
    {
        if (part is null)
            return IrNoteStore.Empty;

        try
        {
            var root = PreparePartRoot(part);
            if (root is null)
                return IrNoteStore.Empty;

            var ctx = new ReadContext(part.Uri, main, styles, numbering, scopeName,
                retainSources: retain, owningPart: part);
            var notes = new Dictionary<string, IrScope>(StringComparer.Ordinal);
            // Insertion-ordered map id → the note element's own pt:Unid (the projection's label source).
            var noteUnids = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (var noteEl in root.Elements(noteName))
            {
                if (WmlToMarkdownConverter.IsBoilerplateNote(noteEl))
                    continue;
                var id = (string?)noteEl.Attribute(W + "id");
                if (id is null || notes.ContainsKey(id))
                    continue; // an id-less note is unreferenceable; first wins on a duplicate id.

                var blocks = new List<IrBlock>();
                foreach (var child in noteEl.Elements())
                    AppendBlocks(child, ctx, blocks);
                foreach (var b in blocks)
                    IndexBlock(b, anchorIndex);

                notes[id] = new IrScope(scopeName, IrNodeList.From(blocks), part.Uri);
                noteUnids[id] = Unid(noteEl);
            }

            if (retain)
                sources[part.Uri] = part.GetXDocument();
            return new IrNoteStore(notes, noteUnids);
        }
        catch (Exception e) when (IsTolerableRegistryException(e))
        {
            return IrNoteStore.Empty;
        }
    }

    /// <summary>
    /// Read the comments part into an <see cref="IrCommentStore"/>. Per <c>w:comment</c>: an anchor of
    /// kind <see cref="IrAnchorKind.Cmt"/> in scope <c>cmt</c> (unid from the comment element's
    /// <c>pt:Unid</c>), <c>@w:author</c> (?? ""), optional <c>@w:initials</c>/<c>@w:date</c> (verbatim),
    /// the comment's block children walked under the <c>cmt</c> scope, and the N15 targets recorded
    /// during the body walk (<see cref="CommentTracker.TargetsFor"/>). The part's
    /// <see cref="XDocument"/> joins <paramref name="sources"/>. A missing part → empty store; the walk
    /// is tolerance-wrapped.
    /// </summary>
    private static IrCommentStore ReadCommentStore(
        MainDocumentPart main, IrStyleRegistry styles, IrNumberingRegistry numbering,
        Dictionary<Uri, XDocument> sources, Dictionary<string, IrBlock> anchorIndex,
        CommentTracker tracker, bool retain)
    {
        var part = main.WordprocessingCommentsPart;
        if (part is null)
            return IrCommentStore.Empty;

        try
        {
            var root = PreparePartRoot(part);
            if (root is null)
                return IrCommentStore.Empty;

            var ctx = new ReadContext(part.Uri, main, styles, numbering, "cmt",
                retainSources: retain, owningPart: part);
            var comments = new List<IrComment>();

            foreach (var commentEl in root.Elements(W + "comment"))
            {
                var anchor = new IrAnchor(IrAnchorKind.Cmt, "cmt", Unid(commentEl));
                var author = (string?)commentEl.Attribute(W + "author") ?? "";
                var initials = (string?)commentEl.Attribute(W + "initials");
                var date = (string?)commentEl.Attribute(W + "date");

                var blocks = new List<IrBlock>();
                foreach (var child in commentEl.Elements())
                    AppendBlocks(child, ctx, blocks);
                foreach (var b in blocks)
                    IndexBlock(b, anchorIndex);

                var commentId = (string?)commentEl.Attribute(W + "id") ?? "";
                var targets = tracker.TargetsFor(commentId);

                comments.Add(new IrComment(anchor, author, initials, date,
                    IrNodeList.From(blocks), IrNodeList.From(targets)));
            }

            if (retain)
                sources[part.Uri] = part.GetXDocument();
            return new IrCommentStore(IrNodeList.From(comments), part.Uri);
        }
        catch (Exception e) when (IsTolerableRegistryException(e))
        {
            return IrCommentStore.Empty;
        }
    }

    // --- registries (styles / numbering / theme) --------------------------

    // The set of exceptions tolerated while building any registry — the same narrow band
    // ResolveImagePart catches. A malformed/missing part or a malformed individual entry is
    // skipped rather than aborting the read; systemic faults (OOM etc.) still escape. Totality
    // over the corpus depends on this: registry building now runs for every file.
    private static bool IsTolerableRegistryException(Exception e) =>
        e is KeyNotFoundException or ArgumentException or System.IO.IOException
            or InvalidOperationException or NotSupportedException or System.Xml.XmlException;

    /// <summary>
    /// Build the <see cref="IrStyleRegistry"/> from the main part's <c>StyleDefinitionsPart</c>.
    /// Each <c>w:style</c> contributes one <see cref="IrStyle"/> (first wins on a duplicate
    /// <c>@w:styleId</c>; a style with no id is skipped as malformed). The default paragraph style
    /// id is the first <c>w:style type="paragraph" w:default="1"</c>. Document defaults come from
    /// <c>w:docDefaults</c>. A missing/malformed part tolerates to <see cref="IrStyleRegistry.Empty"/>.
    /// </summary>
    private static IrStyleRegistry BuildStyleRegistry(MainDocumentPart main)
    {
        try
        {
            var stylesPart = main.StyleDefinitionsPart;
            if (stylesPart is null)
                return IrStyleRegistry.Empty;

            var root = stylesPart.GetXDocument().Root;
            if (root is null)
                return IrStyleRegistry.Empty;

            var styles = new Dictionary<string, IrStyle>(StringComparer.Ordinal);
            string? defaultParagraphStyleId = null;

            foreach (var styleEl in root.Elements(W + "style"))
            {
                IrStyle? style;
                try
                {
                    var id = (string?)styleEl.Attribute(W + "styleId");
                    if (id is null)
                        continue; // malformed: a style without an id is unreferenceable, skip it.

                    var type = (string?)styleEl.Attribute(W + "type") ?? "paragraph";
                    // @w:default is an OOXML toggle attribute (absent → false; "1"/"true"/"on" →
                    // true; "0"/"false"/"off" → false).
                    var isDefault = (string?)styleEl.Attribute(W + "default") is "1" or "true" or "on";

                    style = new IrStyle(
                        Id: id,
                        Name: AttrVal(styleEl.Element(W + "name")),
                        BasedOn: AttrVal(styleEl.Element(W + "basedOn")),
                        Type: type,
                        IsDefault: isDefault)
                    {
                        PPr = CloneOrNull(styleEl.Element(W + "pPr")),
                        RPr = CloneOrNull(styleEl.Element(W + "rPr")),
                    };
                }
                catch (Exception e) when (IsTolerableRegistryException(e))
                {
                    continue; // skip a single malformed entry, keep reading.
                }

                if (!styles.TryAdd(style.Id, style))
                    continue; // first wins on duplicate id.

                if (defaultParagraphStyleId is null && style.IsDefault && style.Type == "paragraph")
                    defaultParagraphStyleId = style.Id;
            }

            var docDefaults = root.Element(W + "docDefaults");
            var docDefaultsPPr = CloneOrNull(
                docDefaults?.Element(W + "pPrDefault")?.Element(W + "pPr"));
            var docDefaultsRPr = CloneOrNull(
                docDefaults?.Element(W + "rPrDefault")?.Element(W + "rPr"));

            return new IrStyleRegistry(styles, defaultParagraphStyleId, docDefaultsPPr, docDefaultsRPr);
        }
        catch (Exception e) when (IsTolerableRegistryException(e))
        {
            return IrStyleRegistry.Empty;
        }
    }

    /// <summary>
    /// Build the <see cref="IrNumberingRegistry"/> from the main part's
    /// <c>NumberingDefinitionsPart</c>. <c>w:abstractNum</c> → <see cref="IrAbstractNum"/> (levels
    /// keyed by ilvl); <c>w:num</c> → <see cref="IrNum"/> (abstractNumId + start overrides). First
    /// wins on duplicate abstractNumId/numId. An abstractNum carrying a <c>w:numStyleLink</c>
    /// indirection is recorded with whatever levels it has (possibly none) — see the TODO below; we
    /// do not chase the indirection at this tier. A missing/malformed part tolerates to
    /// <see cref="IrNumberingRegistry.Empty"/>.
    /// </summary>
    private static IrNumberingRegistry BuildNumberingRegistry(MainDocumentPart main)
    {
        try
        {
            var numberingPart = main.NumberingDefinitionsPart;
            if (numberingPart is null)
                return IrNumberingRegistry.Empty;

            var root = numberingPart.GetXDocument().Root;
            if (root is null)
                return IrNumberingRegistry.Empty;

            var abstractNums = new Dictionary<int, IrAbstractNum>();
            foreach (var absEl in root.Elements(W + "abstractNum"))
            {
                try
                {
                    var absId = IntAttr(absEl, W + "abstractNumId");
                    if (absId is null)
                        continue;
                    if (abstractNums.ContainsKey(absId.Value))
                        continue; // first wins.

                    // TODO(M1.4+): numStyleLink indirection — an abstractNum with a w:numStyleLink
                    // borrows its levels from another abstractNum via a paragraph style. We do NOT
                    // resolve that multi-hop chain here; such an abstractNum lands with whatever
                    // explicit levels it carries (often none).
                    var levels = new Dictionary<int, IrNumLevel>();
                    foreach (var lvlEl in absEl.Elements(W + "lvl"))
                    {
                        try
                        {
                            var ilvl = IntAttr(lvlEl, W + "ilvl");
                            if (ilvl is null)
                                continue;
                            if (levels.ContainsKey(ilvl.Value))
                                continue; // first wins.

                            levels[ilvl.Value] = new IrNumLevel(
                                Ilvl: ilvl.Value,
                                NumberFormat: AttrVal(lvlEl.Element(W + "numFmt")) ?? "decimal",
                                Start: IntAttr(lvlEl.Element(W + "start"), W + "val"),
                                LvlText: AttrVal(lvlEl.Element(W + "lvlText")))
                            {
                                PPr = CloneOrNull(lvlEl.Element(W + "pPr")),
                            };
                        }
                        catch (Exception e) when (IsTolerableRegistryException(e))
                        {
                            continue; // skip a single malformed level.
                        }
                    }

                    abstractNums[absId.Value] = new IrAbstractNum(absId.Value, levels);
                }
                catch (Exception e) when (IsTolerableRegistryException(e))
                {
                    continue; // skip a single malformed abstractNum.
                }
            }

            var nums = new Dictionary<int, IrNum>();
            foreach (var numEl in root.Elements(W + "num"))
            {
                try
                {
                    var numId = IntAttr(numEl, W + "numId");
                    if (numId is null)
                        continue;
                    if (nums.ContainsKey(numId.Value))
                        continue; // first wins.

                    var absId = IntAttr(numEl.Element(W + "abstractNumId"), W + "val");
                    if (absId is null)
                        continue; // a num with no abstractNumId reference is unusable.

                    var startOverrides = new Dictionary<int, int>();
                    foreach (var ovrEl in numEl.Elements(W + "lvlOverride"))
                    {
                        var ilvl = IntAttr(ovrEl, W + "ilvl");
                        var startOverride = IntAttr(ovrEl.Element(W + "startOverride"), W + "val");
                        if (ilvl is not null && startOverride is not null
                            && !startOverrides.ContainsKey(ilvl.Value))
                            startOverrides[ilvl.Value] = startOverride.Value; // first wins.
                    }

                    nums[numId.Value] = new IrNum(numId.Value, absId.Value, startOverrides);
                }
                catch (Exception e) when (IsTolerableRegistryException(e))
                {
                    continue; // skip a single malformed num.
                }
            }

            return new IrNumberingRegistry(nums, abstractNums);
        }
        catch (Exception e) when (IsTolerableRegistryException(e))
        {
            return IrNumberingRegistry.Empty;
        }
    }

    /// <summary>
    /// Build the <see cref="IrThemeFonts"/> from the main part's <c>ThemePart</c>'s
    /// <c>a:fontScheme</c>: major (heading) and minor (body) Latin/ASCII typefaces. A
    /// missing/malformed part tolerates to <see cref="IrThemeFonts.Empty"/>.
    /// </summary>
    private static IrThemeFonts BuildThemeFonts(MainDocumentPart main)
    {
        try
        {
            var themePart = main.ThemePart;
            if (themePart is null)
                return IrThemeFonts.Empty;

            var root = themePart.GetXDocument().Root;
            var fontScheme = root?.Descendants(Docxodus.A.fontScheme).FirstOrDefault();
            if (fontScheme is null)
                return IrThemeFonts.Empty;

            string? major = (string?)fontScheme.Element(Docxodus.A.majorFont)
                ?.Element(Docxodus.A.latin)?.Attribute("typeface");
            string? minor = (string?)fontScheme.Element(Docxodus.A.minorFont)
                ?.Element(Docxodus.A.latin)?.Attribute("typeface");

            if (major is null && minor is null)
                return IrThemeFonts.Empty;
            return new IrThemeFonts(major, minor);
        }
        catch (Exception e) when (IsTolerableRegistryException(e))
        {
            return IrThemeFonts.Empty;
        }
    }

    /// <summary>Deep clone of <paramref name="el"/> (detached from its source tree), or null.</summary>
    private static XElement? CloneOrNull(XElement? el) => el is null ? null : new XElement(el);

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
        else if (block is IrParagraph para)
            // Textbox inner blocks are addressable too (the oracle's DescendantsAndSelf index walk
            // reaches them). Register each so the AnchorIndex carries the same anchor set the oracle
            // produces; the recursion bottoms out naturally (inner blocks may themselves be paragraphs
            // bearing nested textboxes).
            foreach (var inline in para.Inlines)
                if (inline is IrTextbox tb)
                    foreach (var inner in tb.Blocks)
                        IndexBlock(inner, index);
    }

    // --- helpers ----------------------------------------------------------

    /// <summary>
    /// Canonical-hash a synthetic <c>&lt;unmodeled&gt;</c> container holding clones of every child
    /// of <paramref name="props"/> whose name is NOT in <paramref name="consumed"/> and is not the
    /// optional <paramref name="alsoConsumed"/> name (used for conditionally-consumed children so
    /// callers need not allocate a fresh set). When there are no leftovers the result is the cached
    /// empty-container digest (§6.4).
    /// </summary>
    private static IrHash UnmodeledDigest(XElement props, HashSet<XName> consumed,
        XName? alsoConsumed = null, XElement? extra = null)
    {
        var leftovers = props.Elements()
            .Where(e => !consumed.Contains(e.Name) && e.Name != alsoConsumed)
            .ToList();
        if (leftovers.Count == 0 && extra is null)
            return EmptyUnmodeledDigest;

        var container = new XElement("unmodeled");
        foreach (var e in leftovers)
            container.Add(new XElement(e));
        if (extra is not null)
            container.Add(new XElement(extra));
        return IrHasher.CanonicalHash(container);
    }

    /// <summary>Digest of an <c>&lt;unmodeled&gt;</c> container holding a single extra element
    /// (used when a glyph-bearing run child like <c>w:sym</c> must influence the fingerprint of a
    /// run that has no <c>w:rPr</c> of its own).</summary>
    private static IrHash ExtraUnmodeledDigest(XElement extra)
    {
        var container = new XElement("unmodeled", new XElement(extra));
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
