#nullable enable

// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Verification;
using DocumentFormat.OpenXml.Packaging;
using GridCell = Docxodus.Internal.TableGridCell;

namespace Docxodus;

// ─── Session ───────────────────────────────────────────────────────────────

/// <summary>
/// A stateful, anchor-addressed editing session over one DOCX package: project it, edit it by
/// anchor, undo and redo, and save it back to bytes. See <c>docs/architecture/docx_mutation_api.md</c>.
/// </summary>
/// <remarks>
/// <para><b>After <see cref="Dispose"/>.</b> A member whose result type carries an error reports
/// <see cref="EditErrorCode.SessionDisposed"/> in that result (every <see cref="EditResult"/> op,
/// and the other <c>*Result</c> types that have a failure state). A member with no error channel,
/// such as a query, a listing or <see cref="Save()"/>, throws <see cref="ObjectDisposedException"/>.
/// <see cref="Undo"/> and <see cref="Redo"/> return <c>false</c>, as they do whenever there is
/// nothing to undo or redo.</para>
/// </remarks>
public sealed partial class DocxSession : IDisposable
{
    /// <summary>Characters of surrounding text <see cref="Grep"/>, <see cref="GrepCrossBlock"/> and
    /// <see cref="FindPlaceholders"/> report on each side of a match when the caller names no
    /// width. Every transport reads it from here rather than repeating the literal.</summary>
    public const int DefaultContextChars = 80;

    private readonly DocxSessionSettings _settings;
    private readonly Internal.UndoRing<DocumentSnapshot> _history;
    private MemoryStream? _stream;
    private WordprocessingDocument? _doc;
    private MarkdownProjection? _cachedProjection;
    private MarkdownProjection? _initialProjection;
    private byte[]? _initialPackageBytes;
    private byte[]? _initialCheckpointBytes;
    private bool _disposed;
    private long _version;
    private PageMap? _registeredPageMap;
    private readonly object _mutationGate = new();
    private readonly Stack<TransactionState> _transactions = new();
    private long _nextTransactionId;
    private int _transactionPendingMutations;
    // Monotonic within a transaction: incremented on every recorded op and NEVER decremented.
    // _transactionPendingMutations tracks the live history depth and so falls back to its
    // baseline when an op self-rolls-back (RollbackFailedOp pops its own pre-op entry) — it
    // cannot answer "did this scope ever touch the package?". Ops write parts that the
    // selective per-op snapshot deliberately excludes (the numbering part), so a scope whose
    // pending count nets back to baseline can still have left the package dirty. That question
    // is what decides whether a rollback must reopen the checkpoint, so it gets its own witness.
    private long _transactionMutationEpoch;
    private int _revisionCounter = 1000;
    /// <summary>False until <see cref="NextRevisionId"/> has raised
    /// <see cref="_revisionCounter"/> past every <c>w:id</c> already live in the document.
    /// Seeding is lazy so a session that never records a tracked change never pays the
    /// story-part walk.</summary>
    private bool _revisionCounterSeeded;
    private long _lastFormatRevisionTicks;
    private RawDocxOps? _raw;

    internal bool IsDisposed => _disposed;

    // Mutable session configuration (issue #304): seeded from _settings at construction,
    // switchable mid-session via SetTrackedChanges/SetRevisionAuthor. Session config, not
    // document state — never captured in undo snapshots.
    private TrackedChangeMode _trackedChanges;
    private string? _revisionAuthor;

    /// <summary>
    /// Previews retained for a guarded commit (issue #760). Settable so tests can bound and clock
    /// the store; production sessions use the defaults.
    /// </summary>
    internal Internal.RetainedPreviews RetainedPreviews { get; set; } = new();

    /// <summary>Host-owned receipt evidence capture (issue #748); null unless the session was
    /// opened with <see cref="DocxSessionSettings.CaptureDeliveryEvidence"/>.</summary>
    private readonly Internal.DeliveryEvidenceRecorder? _deliveryEvidence;

    /// <summary>A transport-composed batch the recorder is currently attributing version steps to.</summary>
    private Internal.DeliveryEvidenceRecorder.PendingBatch? _clientEvidence;

    internal Internal.DeliveryEvidenceRecorder? DeliveryEvidence => _deliveryEvidence;

    internal bool InTransactionScope => _transactions.Count > 0;

    /// <summary>
    /// Set only on a preview shadow: the live session it was cloned from and the live state at
    /// cloning, which a retained preview is bound to. The base bytes are the clone's own source
    /// package, kept so the base hash is computed only when retention is requested.
    /// </summary>
    private ShadowOrigin? _shadowOrigin;

    private sealed record ShadowOrigin(
        DocxSession Owner,
        long BaseVersion,
        byte[] BaseBytes,
        TrackedChangeMode TrackedChanges,
        string? RevisionAuthor);

    private sealed record TransactionState(
        long Id,
        int OwnerThreadId,
        DocumentSnapshot PackageSnapshot,
        Internal.UndoRing<DocumentSnapshot>.State History,
        int PendingMutations,
        long MutationEpoch,
        TrackedChangeMode TrackedChanges,
        string? RevisionAuthor,
        Exception? LastInternalError,
        Exception? LastRollbackError);

    public DocxSession(byte[] docxBytes, DocxSessionSettings? settings = null)
        : this(docxBytes, settings, skipInitialProjectionCapture: false)
    {
    }

    private DocxSession(
        byte[] docxBytes,
        DocxSessionSettings? settings,
        bool skipInitialProjectionCapture)
    {
        ArgumentNullException.ThrowIfNull(docxBytes);
        _settings = settings ?? new DocxSessionSettings();
        _trackedChanges = _settings.TrackedChanges;
        _revisionAuthor = _settings.RevisionAuthor;
        _history = new Internal.UndoRing<DocumentSnapshot>(
            _settings.UndoDepth,
            _settings.UndoMemoryBudgetBytes,
            static snapshot => snapshot.ApproximateBytes,
            onRecordPreOp: _ => OnHistoryRecordPreOp(),
            onPopUndo: snapshot => OnHistoryPopUndo(snapshot));
        _stream = new MemoryStream();
        _stream.Write(docxBytes, 0, docxBytes.Length);
        _stream.Position = 0;
        _doc = WordprocessingDocument.Open(_stream, isEditable: true);

        if (_settings.CaptureInitialProjection && !skipInitialProjectionCapture)
        {
            _initialPackageBytes = docxBytes.ToArray();
            _initialProjection = WmlToMarkdownConverter.Convert(_doc!, _settings.ProjectionSettings);
        }

        if (_settings.CaptureDeliveryEvidence && !skipInitialProjectionCapture)
        {
            if (_initialPackageBytes is null)
                throw new ArgumentException(
                    "CaptureDeliveryEvidence requires CaptureInitialProjection: the opening package is the receipt's source document.",
                    nameof(settings));
            _deliveryEvidence = new Internal.DeliveryEvidenceRecorder(this, _initialPackageBytes, _version);
        }
    }

    public Exception? LastInternalError { get; private set; }

    /// <summary>
    /// Monotonic in-session document version. Starts at 0 and advances once after each
    /// committed mutation and each successful undo/redo. Failed calls, failed preconditions,
    /// and successful no-ops leave it unchanged.
    /// </summary>
    public long Version => _version;

    /// <summary>
    /// Validate and register an externally materialized layout map. Registration is read-only:
    /// it neither changes the document version nor participates in undo. A later committed
    /// mutation/undo/redo makes the map stale automatically because its document version no
    /// longer matches <see cref="Version"/>.
    /// </summary>
    public PageMapRegistrationResult RegisterPageMap(
        PageMap pageMap, string? expectedRendererFingerprint = null)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(pageMap);

        PageMapRegistrationResult Fail(PageMapRegistrationError error, string message) =>
            new() { Success = false, Error = error, Message = message };

        var portable = PageMapContract.ValidatePortable(
            pageMap, _version, expectedRendererFingerprint);
        if (!portable.Success)
            return portable;

        foreach (var fragment in pageMap.Fragments)
        {
            var target = FindAnchor(fragment.AnchorId);
            if (target is null || !string.Equals(target.Anchor.Id, fragment.AnchorId, StringComparison.Ordinal))
                return Fail(PageMapRegistrationError.InvalidMap,
                    $"PageMap fragment refers to unknown or non-canonical anchor: {fragment.AnchorId}");

            if (!PageMapContract.StoryMatchesScope(fragment.Story, target.Anchor.Scope))
                return Fail(PageMapRegistrationError.InvalidMap,
                    $"PageMap story does not match anchor scope: {fragment.AnchorId}");

            var element = target.Resolve(_doc!);
            var actuallyInTableCell = target.Anchor.Kind == "tc"
                || (element?.AncestorsAndSelf(W.tc).Any() ?? false);
            // A comment's canonical source lives in comments.xml, while its inline presentation
            // lives at the referenced range in the main story. A true table flag therefore must
            // be proven from a live body-side marker; false also validly describes the definition,
            // endnote-style, margin, or an out-of-table inline presentation.
            if (fragment.Story == PageMapStory.Comment
                && fragment.InTableCell
                && !CommentHasTableCellPresentation(element))
                return Fail(PageMapRegistrationError.InvalidMap,
                    $"PageMap comment has no table-cell presentation: {fragment.AnchorId}");
            if (fragment.Story != PageMapStory.Comment
                && fragment.InTableCell != actuallyInTableCell)
                return Fail(PageMapRegistrationError.InvalidMap,
                    $"PageMap inTableCell does not match anchor ownership: {fragment.AnchorId}");

        }

        _registeredPageMap = pageMap with
        {
            Pages = pageMap.Pages.ToArray(),
            Fragments = pageMap.Fragments.ToArray(),
        };
        return new PageMapRegistrationResult { Success = true };
    }

    /// <summary>Return explicit availability for the currently registered map.</summary>
    public PageMapStatus GetPageMapStatus(PageCitationRequest? request = null)
    {
        ThrowIfDisposed();
        var map = _registeredPageMap;
        if (map is null)
            return new PageMapStatus
            {
                Availability = PageMapAvailability.Unavailable,
                UnavailableReason = PageCitationUnavailableReason.NoPageMap,
                DocumentVersion = _version,
            };
        if (map.DocumentVersion != _version || (request is not null && request.DocumentVersion != _version))
            return new PageMapStatus
            {
                Availability = PageMapAvailability.Unavailable,
                UnavailableReason = PageCitationUnavailableReason.StaleDocumentVersion,
                DocumentVersion = _version,
                RendererFingerprint = map.RendererFingerprint,
                Mode = map.Mode,
            };
        if (request is not null && !string.Equals(
                request.RendererFingerprint, map.RendererFingerprint, StringComparison.Ordinal))
            return new PageMapStatus
            {
                Availability = PageMapAvailability.Unavailable,
                UnavailableReason = PageCitationUnavailableReason.RendererFingerprintMismatch,
                DocumentVersion = _version,
                RendererFingerprint = map.RendererFingerprint,
                Mode = map.Mode,
            };
        if (map.Mode == PageMapMode.Continuous || map.Availability == PageMapAvailability.Unavailable)
            return new PageMapStatus
            {
                Availability = PageMapAvailability.Unavailable,
                UnavailableReason = PageCitationUnavailableReason.ContinuousMode,
                DocumentVersion = _version,
                RendererFingerprint = map.RendererFingerprint,
                Mode = map.Mode,
            };
        return new PageMapStatus
        {
            Availability = PageMapAvailability.Available,
            DocumentVersion = _version,
            RendererFingerprint = map.RendererFingerprint,
            Mode = map.Mode,
        };
    }

    /// <summary>Resolve every rendered fragment for one canonical anchor.</summary>
    public PageCitation GetPageCitation(string anchorId, PageCitationRequest request)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        ArgumentNullException.ThrowIfNull(request);
        var status = GetPageMapStatus(request);
        if (status.Availability == PageMapAvailability.Unavailable)
            return UnavailableCitation(anchorId, request, status.UnavailableReason!.Value);

        return PageMapContract.ProjectCitation(_registeredPageMap!, anchorId, request);
    }

    private bool CommentHasTableCellPresentation(XElement? source)
    {
        var commentId = (string?)source?.AncestorsAndSelf(W.comment).FirstOrDefault()?.Attribute(W.id);
        var mainRoot = _doc?.MainDocumentPart?.GetXDocument().Root;
        if (commentId is null || mainRoot is null) return false;
        return mainRoot.Descendants()
            .Where(element => element.Name == W.commentRangeStart
                || element.Name == W.commentRangeEnd
                || element.Name == W.commentReference)
            .Any(element => (string?)element.Attribute(W.id) == commentId
                && element.Ancestors(W.tc).Any());
    }

    private PageCitation UnavailableCitation(
        string anchorId, PageCitationRequest request, PageCitationUnavailableReason reason) =>
        new()
        {
            AnchorId = anchorId,
            Availability = PageMapAvailability.Unavailable,
            UnavailableReason = reason,
            DocumentVersion = _version,
            RendererFingerprint = request.RendererFingerprint,
        };

    /// <summary>
    /// Set when a mutation threw AND the subsequent rollback to its pre-op snapshot ALSO threw —
    /// the one case in which a failed op can leave the document partially mutated. Null on a
    /// healthy session, including one that has seen ordinary <see cref="LastInternalError"/>
    /// failures (those rolled back cleanly). Treat a non-null value as "this session's document is
    /// no longer trustworthy": reopen from the last known-good bytes rather than continuing to edit.
    /// </summary>
    public Exception? LastRollbackError { get; private set; }

    /// <summary>Undo steps currently available. Bounded by both
    /// <see cref="DocxSessionSettings.UndoDepth"/> and
    /// <see cref="DocxSessionSettings.UndoMemoryBudgetBytes"/>.</summary>
    public int UndoCount => _history.UndoCount;

    /// <summary>Redo steps currently available.</summary>
    public int RedoCount => _history.RedoCount;

    /// <summary>Approximate heap retained by undo/redo snapshots, in bytes. Always 0 when the
    /// memory budget is disabled (<see cref="DocxSessionSettings.UndoMemoryBudgetBytes"/> &lt;= 0),
    /// because nothing is measured in that mode.</summary>
    public long UndoMemoryBytes => _history.RetainedBytes;

    /// <summary>
    /// True once the memory budget — rather than the depth cap — has discarded undo history.
    /// Sticky for the session's lifetime. An editor can surface this to explain why undo does not
    /// reach as far back as the configured depth; a headless caller can treat it as a signal to
    /// raise <see cref="DocxSessionSettings.UndoMemoryBudgetBytes"/> or work in smaller sessions.
    /// </summary>
    public bool UndoHistoryTrimmedForMemory => _history.EvictedForMemory;

    /// <summary>How subsequent mutations are recorded — switchable mid-session (issue #304).</summary>
    public TrackedChangeMode TrackedChanges => _trackedChanges;

    /// <summary>Author stamped on tracked-change markup; null means the "docxodus" default.</summary>
    public string? RevisionAuthor => _revisionAuthor;

    /// <summary>
    /// Switch how subsequent mutations are recorded. Session configuration, not a document
    /// mutation: takes no undo snapshot (Undo/Redo never change the mode) and never touches
    /// already-applied markup — switching to <see cref="TrackedChangeMode.Accept"/> does not
    /// accept existing revisions, and switching to <see cref="TrackedChangeMode.RenderInline"/>
    /// does not retroactively wrap prior direct edits.
    /// <para>
    /// This is the RECORDING knob. It is distinct from the RENDERING knob,
    /// <see cref="WmlToMarkdownConverterSettings.TrackedChanges"/> (reached via
    /// <see cref="DocxSessionSettings.ProjectionSettings"/>), which controls how existing
    /// markup is projected by <see cref="Project"/>/<see cref="ProjectAnchor"/>. With the
    /// projection knob at its <see cref="TrackedChangeMode.Accept"/> default, a projection
    /// taken after a tracked edit shows the clean accepted text with no inline
    /// <c>{+ins+}</c>/<c>{-del-}</c> markers — that does NOT mean the edit was recorded
    /// untracked. Set the projection knob to <see cref="TrackedChangeMode.RenderInline"/>
    /// to see revision markup in projections (issue #596).
    /// </para>
    /// </summary>
    public void SetTrackedChanges(TrackedChangeMode mode)
    {
        if (_disposed) return;
        _trackedChanges = mode;
    }

    /// <summary>
    /// Change the author stamped on subsequent tracked-change markup (null restores the
    /// "docxodus" default). Session configuration — same non-undoable semantics as
    /// <see cref="SetTrackedChanges"/>. Enables multi-author edits in one session.
    /// </summary>
    public void SetRevisionAuthor(string? author)
    {
        if (_disposed) return;
        _revisionAuthor = author;
    }

    public MarkdownProjection Project()
    {
        ThrowIfDisposed();
        return _cachedProjection ??=
            WmlToMarkdownConverter.Convert(_doc!, _settings.ProjectionSettings);
    }

    /// <summary>
    /// Project a sub-region of the document anchored at <paramref name="anchorId"/>.
    /// Returns a <see cref="MarkdownProjection"/> whose <c>Markdown</c> contains only
    /// the blocks in scope (per <paramref name="depth"/>) and whose <c>AnchorIndex</c>
    /// is filtered to those blocks plus their descendants.
    /// </summary>
    /// <param name="anchorId">The anchor to project. Must exist in the current
    /// <see cref="Project"/>'s AnchorIndex.</param>
    /// <param name="depth">How far below the target to include. Default
    /// <see cref="ProjectionDepth.SubtreeAndFollowingSiblings"/> — for headings, returns
    /// the full section bounded by the next same-or-higher heading.</param>
    /// <returns>A <see cref="MarkdownProjection"/> scoped to the requested region.</returns>
    /// <exception cref="InvalidOperationException">If <paramref name="anchorId"/> isn't in the AnchorIndex.</exception>
    public MarkdownProjection ProjectAnchor(
        string anchorId,
        ProjectionDepth depth = ProjectionDepth.SubtreeAndFollowingSiblings,
        PageCitationRequest? citationRequest = null)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);

        var fullProjection = Project();
        var target = FindAnchor(anchorId)
            ?? throw new InvalidOperationException($"anchor not found: {anchorId}");

        var startElement = target.Resolve(_doc!)
            ?? throw new InvalidOperationException($"anchor element resolved null: {anchorId}");

        // Compute the set of Unids in scope.
        var inRange = new HashSet<string>(StringComparer.Ordinal);
        CollectUnids(startElement, inRange);

        if (depth == ProjectionDepth.SubtreeAndFollowingSiblings && target.Anchor.Kind == "h")
        {
            // For headings, also include forward siblings up to next same-or-higher heading.
            int targetLevel = WmlToMarkdownConverter.HeadingLevel(startElement);
            foreach (var sibling in startElement.ElementsAfterSelf())
            {
                if (sibling.Name == W.p
                    && WmlToMarkdownConverter.IsHeading(sibling)
                    && WmlToMarkdownConverter.HeadingLevel(sibling) <= targetLevel)
                {
                    break;  // hit the section boundary
                }
                CollectUnids(sibling, inRange);
            }
        }
        else if (depth == ProjectionDepth.Subtree)
        {
            // CollectUnids already added self + descendants; nothing more to do.
        }

        // SelfOnly: descendants shouldn't be in scope — keep just the starting element's Unid.
        if (depth == ProjectionDepth.SelfOnly)
        {
            inRange.Clear();
            var selfUnid = (string?)startElement.Attribute(PtOpenXml.Unid);
            if (selfUnid is not null) inRange.Add(selfUnid);
        }

        // Filter the full markdown to blocks whose anchor token is in-range.
        // Blocks are separated by blank lines; each in-range block starts with {#kind:scope:unid}.
        var sb = new System.Text.StringBuilder();
        foreach (var block in fullProjection.Markdown.Split("\n\n"))
        {
            var match = System.Text.RegularExpressions.Regex.Match(block, @"\{#[^:]+:[^:]+:([^\s}]+)\}");
            if (!match.Success) continue;  // skip scope markers / dividers / etc.
            // The rendered id might be the abbreviated or sequential form — translate back
            // to the full Unid via the dual-keyed AnchorIndex.
            if (TryResolveToUnid(match, fullProjection, out var fullUnid)
                && inRange.Contains(fullUnid))
            {
                sb.Append(block).Append("\n\n");
            }
        }

        // Filter the AnchorIndex too — keep only entries whose Unid is in scope.
        var filteredIndex = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
        foreach (var (key, value) in fullProjection.AnchorIndex)
        {
            if (inRange.Contains(value.Unid))
                filteredIndex[key] = value;
        }

        var citations = citationRequest is null
            ? null
            : filteredIndex.Values
                .Select(t => t.Anchor.Id)
                .Distinct(StringComparer.Ordinal)
                .ToDictionary(id => id, id => GetPageCitation(id, citationRequest), StringComparer.Ordinal);

        return new MarkdownProjection
        {
            Markdown = sb.ToString().TrimEnd('\n'),
            AnchorIndex = filteredIndex,
            PageCitations = citations,
        };
    }

    private static void CollectUnids(XElement el, HashSet<string> sink)
    {
        var unid = (string?)el.Attribute(PtOpenXml.Unid);
        if (unid is not null) sink.Add(unid);
        foreach (var d in el.Descendants())
        {
            var dUnid = (string?)d.Attribute(PtOpenXml.Unid);
            if (dUnid is not null) sink.Add(dUnid);
        }
    }

    /// <summary>
    /// Resolve a rendered anchor id (full Unid, abbreviation, or sequential) back to
    /// the underlying full Unid by looking it up in the projection's AnchorIndex
    /// (which is dual-keyed when AnchorIdRendering is Abbreviated/Sequential).
    /// </summary>
    private static bool TryResolveToUnid(
        System.Text.RegularExpressions.Match match,
        MarkdownProjection projection,
        out string fullUnid)
    {
        // The full key is the content between {# and } — works for FullUnid and as an
        // alias key for Abbreviated/Sequential modes (BuildAnchorIndex dual-keys the index).
        var fullKey = match.Value.Substring(2, match.Value.Length - 3);
        if (projection.AnchorIndex.TryGetValue(fullKey, out var target))
        {
            fullUnid = target.Unid;
            return true;
        }
        fullUnid = match.Groups[1].Value;
        return false;
    }

    /// <summary>
    /// Looks up an anchor id with a fallback to Unid-only resolution. The dictionary
    /// is keyed by full <c>kind:scope:unid</c> id, so when a mutation flips the kind
    /// prefix (e.g., <c>p:body:abcd</c> → <c>h:body:abcd</c> after promoting to a
    /// heading), a cached old id would otherwise miss. This helper trails through
    /// to a Unid scan, so agents that hold cached ids keep working — matching the
    /// promise in <c>docs/architecture/docx_mutation_api.md</c>.
    /// </summary>
    /// <summary>
    /// The anchor index for LOOKUP (mutations, EditResult anchors). Reuses the full
    /// projection's index when one is cached; otherwise builds and caches the cheap
    /// index-only variant (no markdown emission, no per-entry TextPreview/AutoNumberPrefix)
    /// — see <see cref="WmlToMarkdownConverter.BuildAnchorIndexOnly"/>. Entries from this
    /// path therefore carry empty previews; consumers that need enrichment must call
    /// <see cref="Project"/> explicitly.
    /// </summary>
    internal IReadOnlyDictionary<string, AnchorTarget> AnchorIndex()
    {
        ThrowIfDisposed();
        if (_cachedProjection is not null) return _cachedProjection.AnchorIndex;
        return _cachedAnchorIndex ??=
            WmlToMarkdownConverter.BuildAnchorIndexOnly(_doc!, _settings.ProjectionSettings);
    }

    private IReadOnlyDictionary<string, AnchorTarget>? _cachedAnchorIndex;

    /// <summary>Whether an anchor index is currently cached (a lookup now would be a dictionary
    /// hit rather than a whole-document rebuild). Lets the block-render path choose a cheaper
    /// resolution strategy right after a mutation invalidated the cache.</summary>
    internal bool HasCachedAnchorIndex => _cachedProjection is not null || _cachedAnchorIndex is not null;

    /// <summary>
    /// The ordered top-level render units per scope container — see <see cref="RenderPlan"/>.
    /// Body = the main body's blocks in document order (each <c>w:p</c> under its
    /// projected kind, each <c>w:tbl</c> as ONE <c>tbl</c> unit), with a block-level
    /// <c>w:sdt</c> flattened to its <c>w:sdtContent</c> blocks — mirroring the renderer,
    /// which strips content controls; Footnotes/Endnotes = the
    /// non-boilerplate note definitions in part order. Elements the projection does not
    /// address (e.g. <c>w:sectPr</c>) are skipped. Unlike the projection itself, empty
    /// paragraphs are ALWAYS listed — the plan mirrors the rendered DOM, which contains
    /// every block regardless of <see cref="EmptyParagraphMode"/>.
    /// </summary>
    public RenderPlan ListBlocks(bool renderTrackedChanges = true)
    {
        ThrowIfDisposed();
        _ = AnchorIndex(); // guarantees Unids are assigned on every projected part

        var body = new List<RenderUnit>();
        var bodyEl = _doc!.MainDocumentPart?.GetXDocument().Root?.Element(W.body);
        if (bodyEl is not null)
        {
            // Mirrors the RENDERER's top-level block sequence, which is the only contract
            // that lets a DOM diff work. The converter strips content controls
            // (RemoveContentControls) and flows their block content inline, so a block-level
            // w:sdt (a TOC is the everyday case) contributes its w:sdtContent blocks as
            // top-level units here — without this, every document containing one diffs as
            // full churn and the incremental reconcile permanently falls back to remount.
            // The renderer wraps each section's units in its own container (see the
            // data-section-index divs), so a unit's section index is part of the plan: a
            // windowed mount needs it to place a unit rendered on its own. A section ends at
            // the block carrying its w:sectPr — a body paragraph's own properties, or a cell
            // paragraph's inside a table (issue #51) — and that block still belongs to it.
            // A unit's group ordinal names the run of units the renderer draws inside one
            // wrapper: adjacent bordered paragraphs share a border div (the converter's
            // CreateBorderDivs) per container. A windowed mount cuts windows between groups,
            // never through them.
            int section = 0;
            int group = 0;
            void AddUnits(XElement container)
            {
                string? previousKey = null;
                foreach (var el in container.Elements())
                {
                    if (el.Name == W.sdt)
                    {
                        if (el.Element(W.sdtContent) is { } content) AddUnits(content);
                        previousKey = null;
                        continue;
                    }
                    string? kind =
                        el.Name == W.tbl ? "tbl" :
                        el.Name == W.p ? WmlToMarkdownConverter.KindFor(el) : null;
                    bool closesSection = ClosesSection(el);
                    if (!renderTrackedChanges && IsRemovedInAcceptedRevisionView(el))
                    {
                        if (closesSection) section++;
                        continue;
                    }
                    // The renderer keys on assembled properties, so a border the style chain
                    // contributes must count here too.
                    var key = WmlToHtmlConverter.BorderGroupKey(el, el.Name == W.p
                        ? FormattingAssembler.ResolveEffectiveParagraphProperties(LiveDocument, el)
                        : null);
                    if (key.Length == 0 || key == "table" || !string.Equals(key, previousKey, StringComparison.Ordinal))
                        group++;
                    previousKey = key;
                    var unid = (string?)el.Attribute(PtOpenXml.Unid);
                    if (kind is not null && unid is not null)
                        body.Add(new RenderUnit($"{kind}:body:{unid}", kind, UnidHelper.ContentHash(el), section, group));
                    if (closesSection) { section++; previousKey = null; }
                }
            }
            AddUnits(bodyEl);
        }

        List<RenderUnit> Notes(XElement? root, XName noteName, bool endnotes, string kindScope)
        {
            var list = new List<RenderUnit>();
            if (root is null) return list;
            // MIRRORS THE RENDERER exactly (WmlToHtmlConverter's notes sections), which
            // is the only contract that lets a DOM diff work:
            //  - with ≥1 citation, the section renders the CITED notes in citation order
            //    (the tracker path) — an uncited note (Word's continuationNotice) does
            //    NOT render;
            //  - with zero citations, it renders every non-separator note in part order
            //    (so an uncited notice DOES render there).
            var cited = ListNotes(endnotes);
            if (cited.Count > 0)
            {
                foreach (var n in cited)
                    list.Add(new RenderUnit(n.DefAnchorId, kindScope,
                        ResolveNoteDef(root, noteName, n.Id) is { } def ? UnidHelper.ContentHash(def) : null));
                return list;
            }
            foreach (var n in root.Elements(noteName))
            {
                if ((string?)n.Attribute(W.type) is "separator" or "continuationSeparator") continue;
                var unid = (string?)n.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                list.Add(new RenderUnit($"{kindScope}:{kindScope}:{unid}", kindScope, UnidHelper.ContentHash(n)));
            }
            return list;
        }

        static XElement? ResolveNoteDef(XElement root, XName noteName, string id) =>
            root.Elements(noteName).FirstOrDefault(n => (string?)n.Attribute(W.id) == id);

        var main = _doc!.MainDocumentPart;
        return new RenderPlan(
            body,
            Notes(main?.FootnotesPart?.GetXDocument().Root, W.footnote, endnotes: false, "fn"),
            Notes(main?.EndnotesPart?.GetXDocument().Root, W.endnote, endnotes: true, "en"));
    }

    /// <summary>Whether a body block carries the section break that ends its section: a
    /// paragraph's own <c>w:pPr/w:sectPr</c>, or one on a paragraph inside a table's cells.</summary>
    internal static bool ClosesSection(XElement block) =>
        block.Name == W.p
            ? block.Element(W.pPr)?.Element(W.sectPr) is not null
            : block.Name == W.tbl && block.Descendants(W.sectPr).Any(s => s.Parent?.Name == W.pPr);

    private static bool IsRemovedInAcceptedRevisionView(XElement block)
    {
        if (block.Name == W.p)
        {
            var mark = block.Element(W.pPr)?.Element(W.rPr);
            return mark?.Element(W.del) is not null || mark?.Element(W.moveFrom) is not null;
        }
        if (block.Name == W.tbl)
        {
            var rows = WordprocessingMLUtil.TableRows(block).ToList();
            return rows.Count > 0 && rows.All(r => r.Element(W.trPr)?.Element(W.del) is not null);
        }
        return false;
    }

    /// <summary>Create a complete isolated clone for handle-based façades and abandonment tests.</summary>
    internal const string NotCapturingDeliveryEvidence =
        "Delivery evidence capture is not enabled for this session; open it with CaptureDeliveryEvidence.";

    /// <summary>
    /// Inline containers a boundary insert steps out of when the boundary is at their edge. Not
    /// <c>w:sdt</c>: the caret at either end of a content control is inside it, as Word's is.
    /// </summary>
    private static readonly HashSet<XName> BoundaryContainers = new() { W.ins, W.moveTo, W.hyperlink, W.smartTag };

    private const int UnifiedContextLines = 3;

    private const int SideBySideColumnWidth = 72;

    /// <summary>
    /// Serialize the current document state. Anchor-id bookkeeping is stripped unless the session
    /// was opened with <see cref="DocxSessionSettings.PersistAnchorIds"/>.
    /// </summary>
    public byte[] Save() => Save(_settings.PersistAnchorIds);

    /// <summary>
    /// Serialize with an explicit choice about the projector's <c>PtOpenXml:Unid</c> bookkeeping,
    /// overriding <see cref="DocxSessionSettings.PersistAnchorIds"/> for this call.
    /// </summary>
    /// <param name="persistAnchorIds">
    /// <c>true</c> keeps the Unid attributes in the output so a re-render of these bytes resolves to
    /// the SAME anchors the live session holds. That is an internal round-trip contract, not a
    /// document feature: the attributes are ~50 bytes each on every projected element (roughly 6x
    /// the file size of a real document), and while Word and LibreOffice both ignore them, bytes
    /// produced this way should not be handed to a user or written to disk as "the document".
    /// <c>false</c> — what a save-to-disk wants — strips them.
    /// </param>
    /// <remarks>
    /// The distinction exists because the two consumers genuinely differ: the browser editor's
    /// remount re-renders saved bytes and needs id stability across that hop, while its
    /// <c>save()</c> produces the file the user downloads. Making it a per-CALL choice rather than a
    /// session-wide setting is what keeps one from contaminating the other.
    /// </remarks>
    public byte[] Save(bool persistAnchorIds)
    {
        ThrowIfDisposed();

        // Serialization is READ-ONLY with respect to package relationships. The orphaned-media
        // sweep that used to run here now runs on the mutation path (see
        // InvalidateProjectionCache), because only a mutation can orphan a relationship and
        // because ConvertToHtml(session) is implemented as Save(persistAnchorIds: true) — a pure
        // render must not be able to delete anything. IM027 pins that invariant.

        if (persistAnchorIds)
        {
            // Flush every projected part's cached XDocument to its stream first.
            // Ops mutate the cached XDocument only; historically the per-op
            // projection rebuild flushed for them (scope.Part.PutXDocument in
            // BuildAnchorIndex), but that flush is now conditional on Unid
            // assignment — this path must not depend on it, or an op that changes
            // content without creating a new Unid (e.g. SetPageNumbering) could
            // serialize stale bytes.
            foreach (var part in EnumerateProjectedParts())
            {
                if (part.GetXDocument().Root is not null) part.PutXDocument();
            }
            _doc!.Save();
            _stream!.Flush();
            _stream.Position = 0;
            return ZipPackageOutputNormalizer.Normalize(_stream.ToArray());
        }

        // Strip the internal PtOpenXml:Unid attributes before serializing — they're
        // projector bookkeeping, not OOXML schema, and on a real document the bloat
        // is significant (each Unid is ~50 bytes and the projector assigns one to
        // every descendant of every projected scope). We snapshot first so the
        // session's in-memory state can keep using Unids after the save completes;
        // Project() / Resolve() rely on them.
        var snapshot = TakeSnapshot();
        try
        {
            StripProjectorBookkeeping(_doc!);
            _doc!.Save();
            _stream!.Flush();
            _stream.Position = 0;
            return ZipPackageOutputNormalizer.Normalize(_stream.ToArray());
        }
        finally
        {
            RestoreSnapshot(snapshot);
        }
    }

    /// <summary>
    /// A clean serialization of the current package — every part payload exactly as
    /// <see cref="Save(bool)"/> with <c>persistAnchorIds: false</c> writes it — produced from a
    /// package clone, so the live stream, caches, and any element an in-flight operation has
    /// already resolved are never touched. The clone re-serializes the package-level
    /// relationship and content-type files, so the ZIP is not byte-identical to a save of a
    /// document Word wrote; part payloads are. Delivery evidence records every package state
    /// through this path, and a delivery built from it hands back these exact bytes as the
    /// delivered document.
    /// </summary>
    internal byte[] SerializeCleanCheckpoint()
    {
        ThrowIfDisposed();
        using var stream = new MemoryStream();
        using (var clone = _doc!.Clone(stream, isEditable: true))
        {
            OverlayCachedParts(_doc!, clone);
            StripProjectorBookkeeping(clone);
            clone.Save();
        }
        return ZipPackageOutputNormalizer.Normalize(stream.ToArray());
    }

    /// <summary>
    /// Strip the internal PtOpenXml:Unid attributes before serializing — they're projector
    /// bookkeeping, not OOXML schema, and on a real document the bloat is significant (each
    /// Unid is ~50 bytes and the projector assigns one to every descendant of every projected
    /// scope). Every projected part is then rewritten, even one with no Unid: that makes clean
    /// output deterministic across a package checkpoint reopen and guarantees cached
    /// settings/story edits are never skipped merely because that part has no anchor.
    /// Rewriting a part is not the same as CHANGING it: PutXDocument preserves the part's
    /// byte-order-mark convention, so a part whose XML did not change comes back byte-identical
    /// (issue #668).
    /// </summary>
    private static void StripProjectorBookkeeping(WordprocessingDocument document)
    {
        foreach (var part in EnumerateProjectedParts(document))
        {
            var xdoc = part.GetXDocument();
            if (xdoc.Root is null) continue;
            // Other Custom XML parts are opaque application data (SharePoint metadata,
            // SDT bindings, ink, and future extensions). They never contain projector Unids,
            // and merely reading one must not cause a clean save to reserialize its payload.
            if (part is CustomXmlPart
                && (xdoc.Root.Name.NamespaceName != Internal.AnnotationsCustomXml.Namespace
                    || xdoc.Root.Name.LocalName != "annotations"))
                continue;
            foreach (var el in xdoc.Root.DescendantsAndSelf())
            {
                var attr = el.Attribute(PtOpenXml.Unid);
                attr?.Remove();
            }
            // A persisted-anchor checkpoint is reopened during transaction rollback/undo.
            // Its pt namespace declaration is then an explicit LINQ-to-XML attribute, unlike
            // the serializer-generated declaration on an in-memory document. Once every Unid
            // is stripped, remove that now-unused declaration too so normal Save output is
            // identical before and after a transaction boundary.
            bool ptNamespaceInUse = xdoc.Root.DescendantsAndSelf().Any(el =>
                el.Name.Namespace == PtOpenXml.pt
                || el.Attributes().Any(a => !a.IsNamespaceDeclaration
                    && a.Name.Namespace == PtOpenXml.pt));
            if (!ptNamespaceInUse)
            {
                var ignorablePrefixes = xdoc.Root.DescendantsAndSelf()
                    .Attributes(MC.Ignorable)
                    .SelectMany(a => a.Value.Split(
                        (char[]?)null, StringSplitOptions.RemoveEmptyEntries))
                    .ToHashSet(StringComparer.Ordinal);
                var declarations = xdoc.Root.DescendantsAndSelf()
                    .Attributes()
                    .Where(a => a.IsNamespaceDeclaration
                        && a.Value == PtOpenXml.pt.NamespaceName
                        // mc:Ignorable contains QNames-as-prefix-tokens. Removing a namespace
                        // declaration that one of those tokens names produces XML that is
                        // well-formed but rejected by the Open XML markup-compatibility reader.
                        && !ignorablePrefixes.Contains(a.Name.LocalName))
                    .ToList();
                if (declarations.Count > 0)
                    declarations.Remove();
            }
            part.PutXDocument();
        }
    }

    /// <summary>
    /// Enumerates every OOXML part the projector walks. Kept centralized so
    /// <see cref="Save"/> (Unid stripping) and any future part-level pass don't drift.
    /// </summary>
    /// <remarks>
    /// Includes every <see cref="CustomXmlPart"/> on the main document because
    /// callers like <see cref="ResolvePart"/> need to be able to look up any
    /// CustomXmlPart by URI. The snapshot/restore path uses
    /// <see cref="EnumerateProjectedPartsForSnapshot"/> instead, which narrows
    /// CustomXmlParts to the annotations part only — see that method for why.
    /// </remarks>
    private IEnumerable<OpenXmlPart> EnumerateProjectedParts() => EnumerateProjectedParts(_doc!);

    private static IEnumerable<OpenXmlPart> EnumerateProjectedParts(WordprocessingDocument document)
    {
        var main = document.MainDocumentPart;
        if (main is null) yield break;
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
        if (main.WordprocessingCommentsPart is not null) yield return main.WordprocessingCommentsPart;
        if (main.WordprocessingCommentsExPart is not null) yield return main.WordprocessingCommentsExPart;
        if (main.WordprocessingCommentsIdsPart is not null) yield return main.WordprocessingCommentsIdsPart;
        if (main.DocumentSettingsPart is not null) yield return main.DocumentSettingsPart;
        if (main.StyleDefinitionsPart is not null) yield return main.StyleDefinitionsPart;
        // Custom XML parts hold annotation metadata; include them so callers that
        // need to look up parts by URI (e.g. ResolvePart) can find them.
        foreach (var cx in main.CustomXmlParts) yield return cx;
    }

    /// <summary>
    /// Snapshot-scoped projected-part enumeration. Same as
    /// <see cref="EnumerateProjectedParts"/> for the structural parts, but narrows
    /// <see cref="DocumentFormat.OpenXml.Packaging.CustomXmlPart"/> enumeration to the Docxodus
    /// <em>annotations</em> CustomXmlPart only (identified by its root namespace
    /// via <see cref="Internal.AnnotationsCustomXml.Find"/>).
    /// </summary>
    /// <remarks>
    /// Why narrow here: <see cref="RestoreSnapshot"/> handles undo-time create/delete
    /// of CustomXmlParts via <c>AddCustomXmlPart(CustomXmlPartType.CustomXml)</c>,
    /// which hard-codes the content type and creates no
    /// <c>CustomXmlPropertiesPart</c> partner. That is correct for the annotations
    /// part but would silently corrupt other CustomXmlParts that Word/SharePoint
    /// rely on (SharePoint metadata, content-type-bound SDT data-binding parts,
    /// inkml, etc.) by re-creating them with the wrong content type and missing
    /// properties partner. Today no session op deletes non-annotation CustomXmlParts
    /// — narrowing here pre-empts the footgun before such an op is added.
    /// </remarks>
    private IEnumerable<OpenXmlPart> EnumerateProjectedPartsForSnapshot()
    {
        var main = _doc!.MainDocumentPart;
        if (main is null) yield break;
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
        if (main.WordprocessingCommentsPart is not null) yield return main.WordprocessingCommentsPart;
        // Comment-threading metadata parts: content is snapshot-scoped so reply/resolve writes
        // and DeleteBlock/RemoveComment pruning are undoable; create/delete reconciliation is
        // driven by DocumentSnapshot.CommentThreadingParts below.
        if (main.WordprocessingCommentsExPart is not null) yield return main.WordprocessingCommentsExPart;
        if (main.WordprocessingCommentsIdsPart is not null) yield return main.WordprocessingCommentsIdsPart;

        // Settings and styles are not walked by the PROJECTOR, but they are WRITTEN by ops, which is
        // what decides snapshot membership. InsertFootnote/InsertEndnote declare w:footnotePr /
        // w:endnotePr in settings and find-or-create the FootnoteText/FootnoteReference styles;
        // EnsureHeaderFooterVisible writes w:titlePg / w:evenAndOddHeaders; AddComment creates the
        // CommentText/CommentReference styles. Leaving them out did not merely weaken the error path
        // — it made those writes SURVIVE AN ORDINARY Undo(), so undoing the first footnote in a
        // document left w:footnotePr and two orphan styles behind permanently.
        if (main.DocumentSettingsPart is not null) yield return main.DocumentSettingsPart;
        if (main.StyleDefinitionsPart is not null) yield return main.StyleDefinitionsPart;
        // List mutations can create the numbering part and tracked rejection must restore the
        // exact pre-edit package, not merely remove the paragraph's numPr. Snapshot both its XML
        // and topology; RestoreSnapshot reconciles create/delete below.
        if (main.NumberingDefinitionsPart is not null) yield return main.NumberingDefinitionsPart;
        var annotationsPart = Internal.AnnotationsCustomXml.Find(_doc);
        if (annotationsPart is not null) yield return annotationsPart;
    }

    private static readonly (XName Start, XName End, string Label)[] CrossBlockRangePairs =
    {
        (W.commentRangeStart, W.commentRangeEnd, "comment"),
        (W.bookmarkStart, W.bookmarkEnd, "bookmark"),
        (W.permStart, W.permEnd, "permission"),
        (W.moveFromRangeStart, W.moveFromRangeEnd, "move-source"),
        (W.moveToRangeStart, W.moveToRangeEnd, "move-destination"),
    };

    private static readonly HashSet<XName> RevisionWrapperNames = new()
    {
        W.ins, W.del, W.moveFrom, W.moveTo,
    };

    private static readonly HashSet<string> AllowedXmlNamespaces = new()
    {
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",        // w:
        "http://schemas.openxmlformats.org/officeDocument/2006/math",          // m:
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", // wp:
        "http://schemas.openxmlformats.org/drawingml/2006/main",               // a:
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships", // r:
        "http://powertools.codeplex.com/2011",                                 // PtOpenXml (Unid)
    };

    // Cached formatting "shell" for session-attached single-block rendering (see
    // Internal.HtmlConversionOps.RenderBlockHtml): an OPEN throwaway .docx holding the formatting
    // parts (styles/numbering/theme/fontTable/settings) + a body that each render replaces
    // wholesale. Kept open across renders so a keystroke commit pays neither the part clone NOR
    // the package re-open + styles/numbering XML re-parse — those parses (and the style/numbering
    // resolution caches the converter annotates onto the part XDocuments) are the dominant fixed
    // cost of a single-block render. HtmlConversionOps owns these; it rebuilds the shell when
    // <see cref="RenderShellSignature"/> (a cheap content signature of the formatting parts) changes
    // — i.e. when a format op adds a style / numbering level. Text edits never touch those parts, so
    // the shell survives normal typing.
    internal WordprocessingDocument? RenderShellDoc;

    /// <summary>
    /// Monotonic counter of comment-family mutations (definition create/update/remove, reply,
    /// resolve/reopen, a body edit of a comment paragraph, and every snapshot restore). The render
    /// shell copies the comments parts, so <c>HtmlConversionOps</c> mixes this into
    /// <see cref="RenderShellSignature"/>: a shell built before the first comment would otherwise
    /// render highlights against an empty comments part forever, while rebuilding on every text
    /// edit would throw away the whole point of caching the shell.
    /// </summary>
    internal int CommentsVersion;
    internal MemoryStream? RenderShellStream;
    internal long RenderShellSignature;
    internal readonly Dictionary<(Internal.HtmlConversionOptions Options, string Xml), XElement> DenseTextRenderTemplates = new();

    // CT_PPr child schema order (subset covering what we insert). w:pPr children must
    // appear in this sequence or Word treats the file as needing repair.
    private static readonly string[] PPrChildOrder =
    {
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
        "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
        "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
        "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
        "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
    };

    // CT_PBdr child schema order. w:pBdr edges must appear in this sequence.
    private static readonly string[] PBdrEdgeOrder = { "top", "left", "bottom", "right", "between", "bar" };

    /// <summary>
    /// Sentinel id the citation run carries until its real id is known. Negative and far below any
    /// legal note id (Word reserves -1 and 0), so it can never be mistaken for a real citation.
    /// </summary>
    private const int NoteIdPlaceholder = int.MinValue;

    // ─── Table styling (issue #315 Stage A), addressed by a canonical tc anchor ───────────
    //
    // Localized w:tblPr / w:trPr / w:tcPr writes over the grid model above.

    // CT_TblPr / CT_TcPr / CT_TrPr / CT_TblBorders child schema order (local names), matching
    // WordprocessingMLUtil's ordering tables.
    private static readonly string[] TblPrChildOrder =
    {
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
        "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
        "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription",
    };

    private static readonly string[] TcPrChildOrder =
    {
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
        "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers",
    };

    /// <summary>CT_TrPr child order: the CT_TrPrBase property set first, then the row
    /// revision marks and finally <c>w:trPrChange</c>. The revision names belong here —
    /// a row that already carries a <c>w:trPrChange</c> (tracked SetTableRowOptions) and is
    /// then deleted in tracked mode must receive its <c>w:del</c> BEFORE that change, not
    /// appended after it.</summary>
    private static readonly string[] TrPrChildOrder =
    {
        "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
        "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden",
        "ins", "del", "trPrChange",
    };

    private static readonly string[] TblBordersEdgeOrder =
    {
        "top", "left", "start", "bottom", "right", "end", "insideH", "insideV",
    };

    // ─── Maintenance / cleanup ───────────────────────────────────────────

    /// <summary>
    /// Remove every <c>w:r</c> in the selected scopes whose only content is a
    /// <c>w:rPr</c> (no text, no tabs, no breaks, no field/footnote/comment
    /// references). Generally useful after any workflow that deletes inline
    /// content — accepting tracked changes, removing footnotes/comments, run-text
    /// refactors — and leaves behind formatting-only runs that the document
    /// model carries but that have no visible effect on rendering.
    /// </summary>
    /// <param name="scopes">Which package parts to compact. Defaults to
    /// <see cref="ProjectionScopes.All"/>.</param>
    /// <returns>How many runs were removed. <c>0</c> means the document was
    /// already compact within the selected scopes.</returns>
    /// <remarks>
    /// One pre-op snapshot is recorded; <see cref="Undo"/> rolls every removal
    /// back together. Block-level anchors (paragraphs / headings / list items /
    /// tables / table cells) are unaffected — runs aren't part of the
    /// <see cref="MarkdownProjection.AnchorIndex"/>.
    /// </remarks>
    public CompactResult CompactRuns(ProjectionScopes scopes = ProjectionScopes.All)
    {
        ThrowIfDisposed();
        _history.RecordPreOp(TakeSnapshot());

        int removed = 0;
        foreach (var part in EnumerateProjectedPartsForScopes(scopes))
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            // Materialize before mutating — Remove() during enumeration is unsafe.
            foreach (var r in root.Descendants(W.r).ToList())
            {
                if (IsEmptyRun(r))
                {
                    r.Remove();
                    removed++;
                }
            }
            part.PutXDocument();
        }
        if (removed > 0) InvalidateProjectionCache();
        else _ = _history.PopForUndo();
        return new CompactResult { RunsRemoved = removed };
    }

    private static bool IsEmptyRun(XElement r)
    {
        foreach (var child in r.Elements())
        {
            if (child.Name == W.rPr) continue;
            // any other child (w:t, w:tab, w:br, w:footnoteReference, …) is meaningful
            return false;
        }
        return true;
    }

    private IEnumerable<OpenXmlPart> EnumerateProjectedPartsForScopes(ProjectionScopes scopes)
    {
        var main = _doc!.MainDocumentPart;
        if (main is null) yield break;
        if (scopes.HasFlag(ProjectionScopes.Body)) yield return main;
        if (scopes.HasFlag(ProjectionScopes.Headers))
            foreach (var h in main.HeaderParts) yield return h;
        if (scopes.HasFlag(ProjectionScopes.Footers))
            foreach (var f in main.FooterParts) yield return f;
        if (scopes.HasFlag(ProjectionScopes.Footnotes) && main.FootnotesPart is not null)
            yield return main.FootnotesPart;
        if (scopes.HasFlag(ProjectionScopes.Endnotes) && main.EndnotesPart is not null)
            yield return main.EndnotesPart;
        if (scopes.HasFlag(ProjectionScopes.Comments) && main.WordprocessingCommentsPart is not null)
            yield return main.WordprocessingCommentsPart;
    }

    /// <summary>
    /// Dispose the session and abandon any active transactions. Active scopes may only be
    /// abandoned by their owner thread; otherwise this throws and leaves the session usable so
    /// that thread can complete them. Successfully disposing invalidates every scope, releases
    /// every recursive mutation-gate entry, and makes later scope disposal a no-op.
    /// </summary>
    public void Dispose()
    {
        if (_disposed) return;
        int activeTransactions = _transactions.Count;
        if (activeTransactions > 0
            && _transactions.Any(t => t.OwnerThreadId != Environment.CurrentManagedThreadId))
            throw new InvalidOperationException(
                "a session with active transactions must be disposed by their owner thread");

        _disposed = true;
        try
        {
            DisposeRenderShell();
            DiscardPackage(_doc, _stream);
            _doc = null;
            _stream = null;
            _raw = null;
            _transactions.Clear();
            _transactionPendingMutations = 0;
            _transactionMutationEpoch = 0;
            RetainedPreviews.Clear();
            _deliveryEvidence?.Clear();
        }
        finally
        {
            // BeginTransaction enters once for each nested scope. Release every recursion count
            // even if package disposal itself reports an error.
            for (int i = 0; i < activeTransactions; i++)
                System.Threading.Monitor.Exit(_mutationGate);
        }
    }

    // ─── Internal mutation helpers (used by tier methods landing in later phases) ───

    /// <summary>
    /// The single point every op reaches once its mutation has landed in the live XML. It drops
    /// the stale projection/anchor caches AND runs the package-wide orphaned-media sweep.
    /// </summary>
    /// <remarks>
    /// Orphaning is a property of MUTATION, not of serialization, so this — not
    /// <see cref="Save(bool)"/> — is the sweep's boundary. Higher-level transforms
    /// (<c>DeleteBlock</c>, <c>DeleteRange</c>, table row/column deletes, <c>ReplaceText</c>, the
    /// raw XML ops) can drop a <c>w:drawing</c> without any native image op running; most of them
    /// already sweep their own resolved owner, but an op that removes markup from a part other
    /// than the one it resolved — or a future op that simply forgets — would leak. Sweeping the
    /// whole story set here makes the invariant structural instead of a per-op checklist.
    /// <para>
    /// The cost is bounded and strictly below what each op already pays: every mutating op runs
    /// <c>_history.RecordPreOp(TakeSnapshot())</c>, which SERIALIZES every projected part, and
    /// <see cref="Internal.OwnedPartRelationships.SweepOrphanedImages"/> returns without reading
    /// XML at all for a story owning no image relationship — so an image-free document pays only
    /// a handful of in-memory relationship enumerations, and an image-bearing one pays one
    /// attribute walk of the trees it was already about to serialize.
    /// </para>
    /// <para>
    /// The undo/redo restore paths deliberately call <see cref="ResetProjectionCache"/> instead:
    /// a snapshot is authoritative over relationship topology (see
    /// <see cref="RestoreImageRelationships"/>), and a document opened with a pre-existing orphan
    /// must not have it swept out from under a restore.
    /// </para>
    /// </remarks>
    internal void InvalidateProjectionCache(bool sweepOrphanedImages = true)
    {
        if (sweepOrphanedImages)
            SweepOrphanedStoryImageRelationships();
        ResetProjectionCache();
    }

    /// <summary>Drop the projection/anchor caches without touching package relationships.</summary>
    private void ResetProjectionCache()
    {
        _cachedProjection = null;
        _cachedAnchorIndex = null;
    }

    /// <summary>
    /// A per-part XML snapshot covering every part the projector / mutation ops walk.
    /// Originally captured only <c>MainDocumentPart</c>, but any cross-part mutation
    /// (footnote definition removal + body reference cleanup, comment range marker
    /// stripping, Save's Unid-strip pass) needs to round-trip all parts — otherwise
    /// undo or the Save restore would leak structural changes into peer parts.
    /// </summary>
    /// <param name="Parts">Per-URI XML content of every snapshot-scoped part (content restore).</param>
    /// <param name="HeaderFooterParts">Relationship id + kind + URI of each header/footer part that
    /// existed at snapshot time. Drives create/delete reconciliation in <see cref="RestoreSnapshot"/>
    /// so ops that add a header/footer part (SetHeaderText/SetFooterText) undo/redo cleanly; the
    /// content is read back from <see cref="Parts"/> by URI when a part must be re-created.</param>
    /// <param name="NoteParts">The same, for the footnotes/endnotes parts, which
    /// InsertFootnote/InsertEndnote create on a document that had no notes.</param>
    /// <param name="CommentParts">The same, for the comments part (0 or 1 entries), which
    /// AddComment creates on a document that had no comments.</param>
    /// <param name="CommentThreadingParts">The same, for commentsExtended/commentsIds, which
    /// AddCommentReply/SetCommentResolved create when upgrading a flat comment.</param>
    /// <param name="StyleParts">The styles relationship/URI when present, so undo can remove a
    /// styles part synthesized by a direct-mode style mutation or recreate it on redo.</param>
    /// <param name="NumberingParts">The numbering relationship/URI when present, so undo and
    /// tracked rejection can remove a part created by a list mutation or recreate one on redo.</param>
    internal sealed record DocumentSnapshot(
        long Version,
        System.Collections.Generic.IReadOnlyList<(string PartUri, XDocument Xml)> Parts,
        System.Collections.Generic.IReadOnlyList<(string RelId, bool IsHeader, string PartUri)> HeaderFooterParts,
        System.Collections.Generic.IReadOnlyList<(string RelId, bool IsFootnote, string PartUri)> NoteParts,
        System.Collections.Generic.IReadOnlyList<(string RelId, string PartUri)> CommentParts,
        System.Collections.Generic.IReadOnlyList<(string RelId, bool IsCommentsEx, string PartUri)> CommentThreadingParts,
        System.Collections.Generic.IReadOnlyList<(string RelId, string PartUri)> StyleParts,
        System.Collections.Generic.IReadOnlyList<(string RelId, string PartUri)> NumberingParts,
        System.Collections.Generic.IReadOnlyList<(string PartUri, string RelId, string Uri, bool IsExternal)> HyperlinkRelationships,
        System.Collections.Generic.IReadOnlyList<(string PartUri, string ContentType, byte[] Bytes)> ImageParts,
        System.Collections.Generic.IReadOnlyList<(string OwnerPartUri, string RelId, string TargetPartUri)> ImageRelationships,
        System.Collections.Generic.IReadOnlyList<(string OwnerPartUri, string RelId, string TargetUri)> LinkedImageRelationships)
    {
        /// <summary>
        /// Optional exact package checkpoint used by transaction boundaries. Unlike the selective
        /// XML snapshot, this includes every part payload and relationship (external hyperlinks,
        /// media, custom XML, and future package topology) and can therefore back an atomic
        /// batch's undo/redo entry without teaching rollback about each relationship type.
        /// </summary>
        internal byte[]? PackageBytes { get; init; }

        internal int? RevisionCounter { get; init; }

        internal long? LastFormatRevisionTicks { get; init; }

        /// <summary>
        /// Approximate retained heap of this snapshot's cloned part trees, for the undo ring's
        /// memory budget. Computed lazily and cached: the ring asks for it at most once per
        /// snapshot, and a session with the budget disabled never asks at all.
        /// </summary>
        internal long ApproximateBytes =>
            _approximateBytes ??= PackageBytes?.LongLength
                ?? (Parts.Sum(p => Internal.XmlMemoryEstimator.Estimate(p.Xml))
                    + ImageParts.Sum(p => (long)p.Bytes.Length));

        private long? _approximateBytes;
    }

    internal DocumentSnapshot TakeSnapshot()
    {
        var parts = new System.Collections.Generic.List<(string, XDocument)>();
        foreach (var part in EnumerateProjectedPartsForSnapshot())
            parts.Add((part.Uri.ToString(), new XDocument(part.GetXDocument())));

        var hfParts = new System.Collections.Generic.List<(string, bool, string)>();
        var noteParts = new System.Collections.Generic.List<(string, bool, string)>();
        var commentParts = new System.Collections.Generic.List<(string, string)>();
        var commentThreadingParts = new System.Collections.Generic.List<(string, bool, string)>();
        var styleParts = new System.Collections.Generic.List<(string, string)>();
        var numberingParts = new System.Collections.Generic.List<(string, string)>();
        var hyperlinkRelationships = new System.Collections.Generic.List<(string, string, string, bool)>();
        var imageParts = new System.Collections.Generic.List<(string, string, byte[])>();
        var imageRelationships = new System.Collections.Generic.List<(string, string, string)>();
        var linkedImageRelationships = new System.Collections.Generic.List<(string, string, string)>();
        var main = _doc!.MainDocumentPart;
        if (main is not null)
        {
            foreach (var h in main.HeaderParts) hfParts.Add((main.GetIdOfPart(h), true, h.Uri.ToString()));
            foreach (var f in main.FooterParts) hfParts.Add((main.GetIdOfPart(f), false, f.Uri.ToString()));
            if (main.FootnotesPart is not null)
                noteParts.Add((main.GetIdOfPart(main.FootnotesPart), true, main.FootnotesPart.Uri.ToString()));
            if (main.EndnotesPart is not null)
                noteParts.Add((main.GetIdOfPart(main.EndnotesPart), false, main.EndnotesPart.Uri.ToString()));
            if (main.WordprocessingCommentsPart is not null)
                commentParts.Add((main.GetIdOfPart(main.WordprocessingCommentsPart), main.WordprocessingCommentsPart.Uri.ToString()));
            if (main.WordprocessingCommentsExPart is not null)
                commentThreadingParts.Add((main.GetIdOfPart(main.WordprocessingCommentsExPart), true,
                    main.WordprocessingCommentsExPart.Uri.ToString()));
            if (main.WordprocessingCommentsIdsPart is not null)
                commentThreadingParts.Add((main.GetIdOfPart(main.WordprocessingCommentsIdsPart), false,
                    main.WordprocessingCommentsIdsPart.Uri.ToString()));
            if (main.StyleDefinitionsPart is not null)
                styleParts.Add((main.GetIdOfPart(main.StyleDefinitionsPart),
                    main.StyleDefinitionsPart.Uri.ToString()));
            if (main.NumberingDefinitionsPart is not null)
                numberingParts.Add((main.GetIdOfPart(main.NumberingDefinitionsPart),
                    main.NumberingDefinitionsPart.Uri.ToString()));
        }
        foreach (var owner in Internal.OwnedPartRelationships.StoryParts(_doc!))
        {
            foreach (var relationship in owner.Part.HyperlinkRelationships)
                hyperlinkRelationships.Add((owner.PartUri, relationship.Id,
                    relationship.Uri.ToString(), relationship.IsExternal));
            foreach (var relationship in Internal.OwnedPartRelationships.ImageRelationships(owner.Part))
            {
                imageRelationships.Add((owner.PartUri, relationship.RelationshipId,
                    relationship.Target.Uri.ToString()));
                if (imageParts.All(part => part.Item1 != relationship.Target.Uri.ToString()))
                    imageParts.Add((relationship.Target.Uri.ToString(), relationship.Target.ContentType,
                        Internal.OwnedPartRelationships.ReadPartBytes(relationship.Target)));
            }
            foreach (var relationship in Internal.OwnedPartRelationships.ExternalImageRelationships(owner.Part))
                linkedImageRelationships.Add((owner.PartUri, relationship.Id, relationship.Uri.ToString()));
        }
        return new DocumentSnapshot(_version, parts, hfParts, noteParts, commentParts,
            commentThreadingParts, styleParts, numberingParts, hyperlinkRelationships, imageParts,
            imageRelationships, linkedImageRelationships);
    }

    /// <summary>
    /// Capture the complete OPC package for an atomic transaction boundary. The ordinary per-op
    /// snapshots remain selective and DOM-based for speed; only a transaction/its undo counterpart
    /// pays the package serialization cost.
    /// </summary>
    internal DocumentSnapshot TakePackageSnapshot()
    {
        var bytes = SerializePackageCheckpoint();
        return new DocumentSnapshot(
            _version,
            Array.Empty<(string PartUri, XDocument Xml)>(),
            Array.Empty<(string RelId, bool IsHeader, string PartUri)>(),
            Array.Empty<(string RelId, bool IsFootnote, string PartUri)>(),
            Array.Empty<(string RelId, string PartUri)>(),
            Array.Empty<(string RelId, bool IsCommentsEx, string PartUri)>(),
            Array.Empty<(string RelId, string PartUri)>(),
            Array.Empty<(string RelId, string PartUri)>(),
            Array.Empty<(string PartUri, string RelId, string Uri, bool IsExternal)>(),
            Array.Empty<(string PartUri, string ContentType, byte[] Bytes)>(),
            Array.Empty<(string OwnerPartUri, string RelId, string TargetPartUri)>(),
            Array.Empty<(string OwnerPartUri, string RelId, string TargetUri)>())
        {
            PackageBytes = bytes,
            RevisionCounter = _revisionCounter,
            LastFormatRevisionTicks = _lastFormatRevisionTicks,
        };
    }

    /// <summary>
    /// Serialize a complete transaction checkpoint without flushing cached XML into the live
    /// package. <see cref="Save(bool)"/> intentionally writes those caches to the owning stream;
    /// using it at transaction begin made a no-op batch observable by adding XML declarations,
    /// namespace declarations, or anchor attributes to later output. Cloning first preserves the
    /// live package stream/cache exactly while still carrying all current part and relationship
    /// topology. Every cached XDocument is then overlaid on its clone counterpart so edits that
    /// have not yet reached a part stream are represented in the checkpoint as well.
    /// </summary>
    internal byte[] SerializePackageCheckpoint() => SerializeCheckpointOf(_doc!);

    /// <summary>
    /// The opening package rendered through the same checkpoint pipeline the current side uses,
    /// so <see cref="GetSemanticChanges"/> compares like against like. Opened read-only from the
    /// retained bytes; a freshly opened package has no cached XDocuments, so the overlay below
    /// is a no-op for it.
    /// </summary>
    private static byte[] NormalizeOpeningPackage(byte[] packageBytes)
    {
        using var source = new MemoryStream(packageBytes, writable: false);
        using var document = WordprocessingDocument.Open(source, isEditable: false);
        return SerializeCheckpointOf(document);
    }

    private static byte[] SerializeCheckpointOf(WordprocessingDocument source)
    {
        using var stream = new MemoryStream();
        using (var clone = source.Clone(stream, isEditable: true))
        {
            OverlayCachedParts(source, clone);
            clone.Save();
        }
        return ZipPackageOutputNormalizer.Normalize(stream.ToArray());
    }

    /// <summary>Overlay every genuinely dirty cached XDocument of <paramref name="source"/> on its
    /// counterpart in <paramref name="clone"/>, so edits that have not reached a part stream are
    /// represented in the clone as well.</summary>
    private static void OverlayCachedParts(WordprocessingDocument source, WordprocessingDocument clone)
    {
        var cloneParts = EnumeratePackageParts(clone)
            .ToDictionary(part => part.Uri.ToString(), StringComparer.Ordinal);
        foreach (var sourcePart in EnumeratePackageParts(source))
        {
            var cached = sourcePart.Annotation<XDocument>();
            if (cached is null) continue;
            if (!cloneParts.TryGetValue(sourcePart.Uri.ToString(), out var clonePart))
                throw new InvalidOperationException(
                    $"package clone omitted part {sourcePart.Uri}");
            // Avoid reserializing an unchanged cached tree: XML declarations, BOMs, and
            // prefix placement are package payload too. A semantic comparison lets the clone
            // preserve the original part bytes when the cache is merely a read-through, while
            // still overlaying every genuinely dirty cached document.
            var clonedXml = clonePart.GetXDocument();
            if (XNode.DeepEquals(cached.Root, clonedXml.Root)) continue;
            clonePart.PutXDocument(new XDocument(cached));
        }
    }

    /// <summary>
    /// Describe the session's current logical OPC package without mutating it. The manifest is
    /// generated from an isolated checkpoint clone which overlays dirty XML caches, so unsaved
    /// edits are included while the live package bytes, caches, history, and version stay intact.
    /// </summary>
    /// <param name="options">Optional safety limits for ZIP/XML inspection.</param>
    /// <returns>A deterministic schema-v1 package manifest.</returns>
    public PackageManifest GetPackageManifest(PackageManifestOptions? options = null)
    {
        ThrowIfDisposed();
        return PackageManifestGenerator.Generate(SerializePackageCheckpoint(), options);
    }

    /// <summary>
    /// Deterministic digest of the current logical OPC package. The checkpoint clone overlays
    /// dirty XDocument caches without writing them to this session. Hashing ordered uncompressed
    /// entry payloads excludes ZIP timestamps/compression while retaining every part byte and
    /// relationship payload, including media and opaque custom XML.
    /// </summary>
    internal string GetPackageContentHash() => HashPackageBytes(SerializePackageCheckpoint());

    /// <summary>The <see cref="GetPackageContentHash"/> digest of an already serialized checkpoint.</summary>
    internal static string HashPackageBytes(byte[] packageBytes)
    {
        try
        {
            using var stream = new MemoryStream(packageBytes, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            var buffer = new byte[81920];
            Span<byte> intBuffer = stackalloc byte[sizeof(int)];
            Span<byte> longBuffer = stackalloc byte[sizeof(long)];
            foreach (var entry in archive.Entries.OrderBy(entry => entry.FullName, StringComparer.Ordinal))
            {
                var name = Encoding.UTF8.GetBytes(entry.FullName);
                BinaryPrimitives.WriteInt32LittleEndian(intBuffer, name.Length);
                hash.AppendData(intBuffer);
                hash.AppendData(name);
                BinaryPrimitives.WriteInt64LittleEndian(longBuffer, entry.Length);
                hash.AppendData(longBuffer);
                using var input = entry.Open();
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                    hash.AppendData(buffer, 0, read);
            }
            return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
        }
        catch (InvalidDataException)
        {
            return Convert.ToHexString(SHA256.HashData(packageBytes)).ToLowerInvariant();
        }
    }

    private static IEnumerable<OpenXmlPart> EnumeratePackageParts(OpenXmlPackage package)
    {
        var pending = new Stack<OpenXmlPart>(package.Parts.Select(pair => pair.OpenXmlPart));
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (pending.Count > 0)
        {
            var part = pending.Pop();
            if (!seen.Add(part.Uri.ToString())) continue;
            yield return part;
            foreach (var child in part.Parts)
                pending.Push(child.OpenXmlPart);
        }
    }

    internal void RestoreSnapshot(DocumentSnapshot snapshot)
    {
        CommentsVersion++;
        // The restored markup is different markup: re-seed the revision counter from it on
        // next use. Seeding only ever raises the counter, so this can never hand out an id
        // that is already live.
        _revisionCounterSeeded = false;
        if (snapshot.RevisionCounter is { } revisionCounter)
            _revisionCounter = revisionCounter;
        if (snapshot.LastFormatRevisionTicks is { } formatTicks)
            _lastFormatRevisionTicks = formatTicks;
        if (snapshot.PackageBytes is { } packageBytes)
        {
            RestorePackage(packageBytes);
            _version = snapshot.Version;
            return;
        }

        var byUri = snapshot.Parts.ToDictionary(p => p.PartUri, p => p.Xml);

        // Restore content for all parts that exist in both snapshot and document.
        // Scoped via EnumerateProjectedPartsForSnapshot — only the annotations
        // CustomXmlPart participates here; other CustomXmlParts (SharePoint
        // metadata, SDT data-binding parts, inkml, …) are intentionally outside
        // the snapshot scope.
        // A part that Save flushes from its cached XDocument needs only its cache restored: the
        // session itself reads through GetXDocument, and the stream is rewritten before the
        // package is serialized. Writing every part's stream here instead re-serialized the whole
        // document on every undo/redo, which on a real file is the bulk of an undo.
        //
        // The snapshot scope is deliberately WIDER than Save's flush scope — the comment-threading
        // parts are snapshot-scoped so reply/resolve/prune are undoable, but their ops persist
        // them themselves rather than relying on Save. For those the stream IS the source of
        // truth at save time, so they keep the flushing restore. Deriving the split from the two
        // enumerations keeps it true if either one changes.
        var flushedBySave = new HashSet<string>(
            EnumerateProjectedParts().Select(p => p.Uri.ToString()), StringComparer.Ordinal);
        foreach (var part in EnumerateProjectedPartsForSnapshot())
        {
            var uri = part.Uri.ToString();
            if (!byUri.TryGetValue(uri, out var xml)) continue;
            if (flushedBySave.Contains(uri)) part.SetXDocumentCache(new XDocument(xml));
            else part.PutXDocument(new XDocument(xml));
        }

        var main = _doc!.MainDocumentPart;

        // Header/footer part create/delete reconcile: SetHeaderText/SetFooterText can add a
        // HeaderPart/FooterPart, so undo/redo must delete the parts the snapshot doesn't have and
        // re-create (with the snapshot's relationship id, so the restored sectPr reference resolves)
        // the ones it does. Content restore above already handled parts present in both by URI.
        if (main is not null)
        {
            ReconcileHeaderFooterParts(main, snapshot, byUri);
            // Same reconcile for the footnotes/endnotes parts, which InsertFootnote/InsertEndnote
            // create on a document that had no notes.
            ReconcileNoteParts(main, snapshot, byUri);
            // And for the comments part, which AddComment creates on a document that had no comments.
            ReconcileCommentsPart(main, snapshot, byUri);
            // Reply/resolve can introduce commentsExtended/commentsIds; reconcile their topology
            // after restoring the base comments part.
            ReconcileCommentThreadingParts(main, snapshot, byUri);
            ReconcileStylePart(main, snapshot, byUri);
            ReconcileNumberingPart(main, snapshot, byUri);
        }

        RestoreHyperlinkRelationships(snapshot);

        // The annotations CustomXmlPart is reconciled the same way (its own factory) — see
        // EnumerateProjectedPartsForSnapshot for why AddCustomXmlPart(CustomXml) is unsafe for
        // non-annotation custom-xml parts (wrong content type, no CustomXmlPropertiesPart partner).
        if (main is not null)
        {
            var annotationsPart = Internal.AnnotationsCustomXml.Find(_doc);
            var snapshotAnnotationsUri = snapshot.Parts
                .FirstOrDefault(p => p.PartUri.StartsWith("/customXml/", StringComparison.OrdinalIgnoreCase))
                .PartUri;

            // Undo direction: snapshot has no annotations part but the live doc
            // does → forward-op created it, roll it back by deleting.
            if (annotationsPart is not null
                && !byUri.ContainsKey(annotationsPart.Uri.ToString()))
            {
                main.DeletePart(annotationsPart);
                annotationsPart = null;
            }

            // Redo direction: snapshot has an annotations part but the live doc
            // doesn't → undo previously removed it, restore by re-adding.
            if (annotationsPart is null && snapshotAnnotationsUri is not null
                && byUri.TryGetValue(snapshotAnnotationsUri, out var annXml))
            {
                var newPart = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                newPart.PutXDocument(new XDocument(annXml));
            }
        }

        // Binary media restoration can require recreating an exact OPC part URI. It is last
        // because it reopens the SDK package graph after low-level part/relationship repair.
        RestoreImageRelationships(snapshot);

        _version = snapshot.Version;
        // Pure cache reset: the snapshot is authoritative over relationship topology, so a sweep
        // here would either be redundant or would delete a relationship the snapshot preserved.
        ResetProjectionCache();
    }

    private void RestorePackage(byte[] packageBytes)
    {
        DisposeRenderShell();
        DiscardPackage(_doc, _stream);

        _stream = new MemoryStream(packageBytes.Length);
        _stream.Write(packageBytes, 0, packageBytes.Length);
        _stream.Position = 0;
        _doc = WordprocessingDocument.Open(_stream, isEditable: true);
        _raw = null;
        // As in RestoreSnapshot: restored package bytes ARE the intended topology.
        ResetProjectionCache();
    }

    /// <summary>Restore reference-relationship topology as well as XML. Without this, undoing a
    /// hyperlink create/delete restores <c>r:id</c> attributes but leaves the corresponding
    /// package relationship in the wrong state.</summary>
    private void RestoreHyperlinkRelationships(DocumentSnapshot snapshot)
    {
        var expectedByPart = snapshot.HyperlinkRelationships
            .GroupBy(r => r.PartUri, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.ToDictionary(r => r.RelId, StringComparer.Ordinal), StringComparer.Ordinal);
        foreach (var owner in Internal.OwnedPartRelationships.StoryParts(_doc!))
        {
            expectedByPart.TryGetValue(owner.PartUri, out var expected);
            expected ??= new System.Collections.Generic.Dictionary<string, (string PartUri, string RelId, string Uri, bool IsExternal)>(StringComparer.Ordinal);
            var live = owner.Part.HyperlinkRelationships.ToDictionary(r => r.Id, StringComparer.Ordinal);
            foreach (var relationship in live.Values)
                if (!expected.ContainsKey(relationship.Id))
                    owner.Part.DeleteReferenceRelationship(relationship.Id);
            foreach (var relationship in expected.Values)
            {
                if (live.TryGetValue(relationship.RelId, out var existing)
                    && existing.Uri.ToString() == relationship.Uri
                    && existing.IsExternal == relationship.IsExternal) continue;
                if (live.ContainsKey(relationship.RelId))
                    owner.Part.DeleteReferenceRelationship(relationship.RelId);
                owner.Part.AddHyperlinkRelationship(
                    new Uri(relationship.Uri, UriKind.RelativeOrAbsolute),
                    relationship.IsExternal, relationship.RelId);
            }
        }
    }

    /// <summary>
    /// Reconcile the live document's header/footer parts against <paramref name="snapshot"/>:
    /// delete parts created since the snapshot (relationship id present live, absent in snapshot)
    /// and re-create parts removed since it (present in snapshot, absent live) with their original
    /// relationship id + content, so the just-restored sectPr <c>headerReference</c>/<c>footerReference</c>
    /// resolves. Parts present in both keep their content (restored by URI in <see cref="RestoreSnapshot"/>).
    /// </summary>
    private static void ReconcileHeaderFooterParts(
        MainDocumentPart main, DocumentSnapshot snapshot,
        System.Collections.Generic.Dictionary<string, XDocument> byUri)
    {
        var snapByRel = new System.Collections.Generic.Dictionary<string, (bool IsHeader, string PartUri)>(StringComparer.Ordinal);
        foreach (var (relId, isHeader, partUri) in snapshot.HeaderFooterParts)
            snapByRel[relId] = (isHeader, partUri);

        // Live header/footer parts keyed by relationship id (materialized so we can DeletePart
        // without mutating a collection we're iterating).
        var live = new System.Collections.Generic.Dictionary<string, OpenXmlPart>(StringComparer.Ordinal);
        foreach (var h in main.HeaderParts) live[main.GetIdOfPart(h)] = h;
        foreach (var f in main.FooterParts) live[main.GetIdOfPart(f)] = f;

        // Delete parts the snapshot doesn't know about (undo of a create).
        foreach (var kv in live)
            if (!snapByRel.ContainsKey(kv.Key))
                main.DeletePart(kv.Value);

        // Re-create parts the snapshot has but the live doc lost (redo of a create / undo of a delete).
        foreach (var kv in snapByRel)
        {
            if (live.ContainsKey(kv.Key)) continue;
            if (!byUri.TryGetValue(kv.Value.PartUri, out var xml)) continue;
            OpenXmlPart np = kv.Value.IsHeader
                ? main.AddNewPart<HeaderPart>(kv.Key)
                : main.AddNewPart<FooterPart>(kv.Key);
            np.PutXDocument(new XDocument(xml));
        }
    }

    /// <summary>
    /// The <see cref="ReconcileHeaderFooterParts"/> twin for the footnotes/endnotes parts: delete a
    /// part created since <paramref name="snapshot"/> (undo of an InsertFootnote/InsertEndnote that
    /// introduced notes) and re-create one the live document has since lost (redo), keeping the
    /// original relationship id so the package relationship the restored XML expects still resolves.
    /// Parts present in both keep their content, restored by URI in <see cref="RestoreSnapshot"/>.
    /// </summary>
    private static void ReconcileNoteParts(
        MainDocumentPart main, DocumentSnapshot snapshot,
        System.Collections.Generic.Dictionary<string, XDocument> byUri)
    {
        var snapByRel = new System.Collections.Generic.Dictionary<string, (bool IsFootnote, string PartUri)>(StringComparer.Ordinal);
        foreach (var (relId, isFootnote, partUri) in snapshot.NoteParts)
            snapByRel[relId] = (isFootnote, partUri);

        var live = new System.Collections.Generic.Dictionary<string, OpenXmlPart>(StringComparer.Ordinal);
        if (main.FootnotesPart is not null) live[main.GetIdOfPart(main.FootnotesPart)] = main.FootnotesPart;
        if (main.EndnotesPart is not null) live[main.GetIdOfPart(main.EndnotesPart)] = main.EndnotesPart;

        foreach (var kv in live)
            if (!snapByRel.ContainsKey(kv.Key))
                main.DeletePart(kv.Value);

        foreach (var kv in snapByRel)
        {
            if (live.ContainsKey(kv.Key)) continue;
            if (!byUri.TryGetValue(kv.Value.PartUri, out var xml)) continue;
            OpenXmlPart np = kv.Value.IsFootnote
                ? main.AddNewPart<FootnotesPart>(kv.Key)
                : main.AddNewPart<EndnotesPart>(kv.Key);
            np.PutXDocument(new XDocument(xml));
        }
    }

    /// <summary>
    /// The <see cref="ReconcileNoteParts"/> twin for the comments part: delete a part created
    /// since <paramref name="snapshot"/> (undo of the AddComment that introduced comments) and
    /// re-create one the live document has since lost (redo), keeping the original relationship
    /// id. Content for a part present in both is restored by URI in <see cref="RestoreSnapshot"/>.
    /// </summary>
    private static void ReconcileCommentsPart(
        MainDocumentPart main, DocumentSnapshot snapshot,
        System.Collections.Generic.Dictionary<string, XDocument> byUri)
    {
        var snapByRel = new System.Collections.Generic.Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (relId, partUri) in snapshot.CommentParts)
            snapByRel[relId] = partUri;

        var live = new System.Collections.Generic.Dictionary<string, OpenXmlPart>(StringComparer.Ordinal);
        if (main.WordprocessingCommentsPart is not null)
            live[main.GetIdOfPart(main.WordprocessingCommentsPart)] = main.WordprocessingCommentsPart;

        foreach (var kv in live)
            if (!snapByRel.ContainsKey(kv.Key))
                main.DeletePart(kv.Value);

        foreach (var kv in snapByRel)
        {
            if (live.ContainsKey(kv.Key)) continue;
            if (!byUri.TryGetValue(kv.Value, out var xml)) continue;
            var np = main.AddNewPart<WordprocessingCommentsPart>(kv.Key);
            np.PutXDocument(new XDocument(xml));
        }
    }

    /// <summary>
    /// Create/delete reconciliation for <c>commentsExtended.xml</c> and
    /// <c>commentsIds.xml</c>. These used to be content-only snapshot parts because no session op
    /// authored them; AddCommentReply/SetCommentResolved can now create either/both, so undo must
    /// remove those parts and redo must restore their original relationship ids and XML.
    /// </summary>
    private static void ReconcileCommentThreadingParts(
        MainDocumentPart main, DocumentSnapshot snapshot,
        System.Collections.Generic.Dictionary<string, XDocument> byUri)
    {
        var snapByRel = new System.Collections.Generic.Dictionary<string, (bool IsCommentsEx, string PartUri)>(StringComparer.Ordinal);
        foreach (var (relId, isCommentsEx, partUri) in snapshot.CommentThreadingParts)
            snapByRel[relId] = (isCommentsEx, partUri);

        var live = new System.Collections.Generic.Dictionary<string, OpenXmlPart>(StringComparer.Ordinal);
        if (main.WordprocessingCommentsExPart is not null)
            live[main.GetIdOfPart(main.WordprocessingCommentsExPart)] = main.WordprocessingCommentsExPart;
        if (main.WordprocessingCommentsIdsPart is not null)
            live[main.GetIdOfPart(main.WordprocessingCommentsIdsPart)] = main.WordprocessingCommentsIdsPart;

        foreach (var kv in live)
            if (!snapByRel.ContainsKey(kv.Key))
                main.DeletePart(kv.Value);

        foreach (var kv in snapByRel)
        {
            if (live.ContainsKey(kv.Key)) continue;
            if (!byUri.TryGetValue(kv.Value.PartUri, out var xml)) continue;
            OpenXmlPart np = kv.Value.IsCommentsEx
                ? main.AddNewPart<WordprocessingCommentsExPart>(kv.Key)
                : main.AddNewPart<WordprocessingCommentsIdsPart>(kv.Key);
            np.PutXDocument(new XDocument(xml));
        }
    }

    /// <summary>Restore numbering-part topology as well as content. In particular, rejecting or
    /// undoing the first tracked list mutation in a document must remove the newly-created part;
    /// redo recreates it with its original relationship id.</summary>
    private static void ReconcileNumberingPart(
        MainDocumentPart main, DocumentSnapshot snapshot,
        System.Collections.Generic.Dictionary<string, XDocument> byUri)
    {
        var snapshotPart = snapshot.NumberingParts.FirstOrDefault();
        var live = main.NumberingDefinitionsPart;

        if (live is not null
            && (snapshotPart.RelId is null
                || !string.Equals(main.GetIdOfPart(live), snapshotPart.RelId, StringComparison.Ordinal)))
        {
            main.DeletePart(live);
            live = null;
        }

        if (live is null && snapshotPart.RelId is not null
            && byUri.TryGetValue(snapshotPart.PartUri, out var xml))
        {
            var restored = main.AddNewPart<NumberingDefinitionsPart>(snapshotPart.RelId);
            restored.PutXDocument(new XDocument(xml));
        }
    }

    /// <summary>Restore styles-part topology as well as content. Style synthesis can create the
    /// optional part, so an ordinary undo must remove it and redo must recreate it.</summary>
    private static void ReconcileStylePart(
        MainDocumentPart main, DocumentSnapshot snapshot,
        System.Collections.Generic.Dictionary<string, XDocument> byUri)
    {
        var snapshotPart = snapshot.StyleParts.FirstOrDefault();
        var live = main.StyleDefinitionsPart;

        if (live is not null
            && (snapshotPart.RelId is null
                || !string.Equals(main.GetIdOfPart(live), snapshotPart.RelId, StringComparison.Ordinal)))
        {
            main.DeletePart(live);
            live = null;
        }

        if (live is null && snapshotPart.RelId is not null
            && byUri.TryGetValue(snapshotPart.PartUri, out var xml))
        {
            var restored = main.AddNewPart<StyleDefinitionsPart>(snapshotPart.RelId);
            restored.PutXDocument(new XDocument(xml));
        }
    }

    /// <summary>
    /// Mint the next <c>w:id</c> for revision markup this session writes. The counter is
    /// seeded from the document's own markup on first use: a fixed starting point collides
    /// with the four-digit ids Word and DocxDiff already emit, and two live groups sharing a
    /// <c>w:id</c> in one part are permanently <see cref="RevisionResolutionStatus.Ambiguous"/>
    /// — neither individually resolvable nor bulk-resolvable, with no recovery short of
    /// hand-editing the XML.
    /// </summary>
    internal int NextRevisionId()
    {
        // Monitor is re-entrant, so this is safe whether or not the caller already holds
        // the gate. Raising the counter is always safe: ids only have to be unique.
        lock (_mutationGate)
        {
            if (!_revisionCounterSeeded)
            {
                _revisionCounterSeeded = true;
                long highest = 0;
                foreach (var story in RevisionStoryParts())
                    highest = Math.Max(highest, Internal.RevisionOps.MaxRevisionId(story.Root));
                if (highest >= int.MaxValue)
                    throw new InvalidOperationException(
                        "The document has exhausted the supported positive w:id revision range.");
                if (highest > _revisionCounter)
                    _revisionCounter = (int)highest;
            }

            if (_revisionCounter >= int.MaxValue)
                throw new InvalidOperationException(
                    "The document has exhausted the supported positive w:id revision range.");
            return ++_revisionCounter;
        }
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DocxSession));
    }

    // ─── Mutation helpers (shared across tiers) ───────────────────────────

    /// <summary>
    /// The per-op patch, or <c>null</c> when <see cref="DocxSessionSettings.EmitMarkdownPatch"/>
    /// is off — every mutation's <c>Patch =</c> site routes through here so the opt-out
    /// cannot be missed by a new op.
    /// </summary>
    private MarkdownPatch? PatchFor(AnchorTarget target) =>
        _settings.EmitMarkdownPatch ? ProjectScope(target) : null;

    internal MarkdownPatch ProjectScope(AnchorTarget target)
    {
        // Phase 3 implementation: re-project the whole document. The patch contract
        // (smallest enclosing block) is honored by ScopeAnchorId; the markdown payload
        // is the full projection until we optimize this in a later phase.
        //
        // Every Patch site runs AFTER the op's InvalidateProjectionCache, so the fresh
        // projection built here IS the post-op state — cache it. Without this, a
        // default-settings caller pays this Convert per op AND a second index build on
        // the next op's FindAnchor.
        var fresh = WmlToMarkdownConverter.Convert(_doc!, _settings.ProjectionSettings);
        _cachedProjection = fresh;
        _cachedAnchorIndex = null;
        return new MarkdownPatch(target.Anchor.Id, fresh.Markdown);
    }

    // Zero-width, semantically-significant bare paragraph children that must survive
    // ReplaceText. Discarding them silently destroys bookmark/comment/permission ranges that
    // point into the paragraph from other parts of the document. The comment reference itself
    // lives inside a run and is covered by MarkerRunContentNames below.
    private static readonly HashSet<XName> PreservedMarkerNames = new()
    {
        W.bookmarkStart, W.bookmarkEnd,
        W.commentRangeStart, W.commentRangeEnd,
        W.permStart, W.permEnd,
        W.proofErr,
    };

    // Zero-width, semantically-significant content that lives INSIDE a <w:r> rather than as
    // a bare paragraph child, so it is detected per run via IsMarkerOnlyRun: the body-side
    // note references (dropping <w:footnoteReference w:id="N"/> orphans the note definition
    // and silently loses content on a text edit — issue B3), the comment reference that
    // makes a comment range visible at all, and — inside a note or comment body — the note's
    // own number mark and separator marks and the comment's annotation mark, without which
    // Word renders the note unnumbered or the comment without its author and date.
    private static readonly HashSet<XName> MarkerRunContentNames = new()
    {
        W.footnoteReference, W.endnoteReference, W.commentReference,
        W.footnoteRef, W.endnoteRef, W.separator, W.continuationSeparator, W.annotationRef,
    };

    // True for a run whose only meaningful (non-rPr) content is a marker — i.e. it carries
    // no visible text. Such a run is preserved by a replacement; a run that mixes a marker
    // with text is ordinary content and is replaced.
    private static bool IsMarkerOnlyRun(XElement e)
    {
        if (e.Name != W.r) return false;
        bool sawMarker = false;
        foreach (var child in e.Elements())
        {
            if (child.Name == W.rPr) continue;
            if (MarkerRunContentNames.Contains(child.Name)) { sawMarker = true; continue; }
            return false; // any other content (w:t, w:tab, w:br, …) ⇒ ordinary run
        }
        return sawMarker;
    }

    // Carriers whose visible width is the sum of their children's, so one that holds nothing
    // but markers is itself a marker: a tracked-inserted note reference, a link holding only a
    // reference, an emptied shell.
    private static readonly HashSet<XName> ZeroWidthCarrierNames = new()
    {
        W.ins, W.del, W.moveTo, W.moveFrom, W.hyperlink, W.fldSimple, W.smartTag,
        W.sdtContent, W.customXml, W.dir, W.bdo,
    };

    /// <summary>A paragraph child a replacement keeps where it sits: a bare range/marker
    /// element, a marker-only run, or an envelope or carrier containing nothing else.</summary>
    private static bool IsZeroWidthMarker(XElement e) =>
        PreservedMarkerNames.Contains(e.Name)
        || (e.Name == W.r ? IsMarkerOnlyRun(e)
            : e.Name == W.sdt ? e.Element(W.sdtContent) is not { } content || IsZeroWidthMarker(content)
            : ZeroWidthCarrierNames.Contains(e.Name) && e.Elements().All(IsZeroWidthMarker));

    private sealed record PreservedMarkerPosition(XElement Element, int Offset, int Order);

    /// <summary>
    /// Capture zero-width markers before a whole-paragraph replacement. Bookmark endpoints retain
    /// their old character coordinate when the replacement is long enough and clamp to its end
    /// otherwise. This is deterministic and, unlike the old leading-marker fallback, cannot invert
    /// or collapse a range merely because its end marker originally sat between two runs.
    /// </summary>
    private static List<PreservedMarkerPosition> CapturePreservedMarkerPositions(XElement paragraph)
    {
        var candidates = paragraph.Elements()
            .Where(IsZeroWidthMarker)
            .Concat(paragraph.Descendants()
                .Where(e => e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd))
            .Distinct()
            .OrderBy(e => e, Comparer<XElement>.Create(XNode.DocumentOrderComparer.Compare))
            .ToList();
        var result = candidates.Select((element, order) =>
            new PreservedMarkerPosition(element, MarkerOffset(paragraph, element), order)).ToList();
        foreach (var marker in candidates) marker.Remove();
        return result;
    }

    private static void RestorePreservedMarkerPositions(
        XElement paragraph, IReadOnlyList<PreservedMarkerPosition> markers)
    {
        int length = ParagraphText(paragraph).Length;
        foreach (var group in markers.GroupBy(m => Math.Clamp(m.Offset, 0, length)).OrderBy(g => g.Key))
            InsertMarkersAtOffset(paragraph, group.Key,
                group.OrderBy(m => m.Order).Select(m => m.Element).ToList());
    }

    /// <summary>
    /// If <paramref name="paragraph"/> carries a resolvable <c>w:numPr</c> auto-number
    /// (e.g. <c>"1."</c>, <c>"Fourth"</c>), strip a matching leading prefix from
    /// <paramref name="payload"/> plus one optional separator character (ASCII space,
    /// tab, or NBSP — matching the projector's emission and the common variants an
    /// agent might use). Idempotent when the prefix isn't present.
    /// </summary>
    private string StripResolvedAutoNumberPrefix(XElement paragraph, string payload)
    {
        if (string.IsNullOrEmpty(payload)) return payload;
        // ListItemRetrieverSettings is internal to the projector; pass null so the
        // resolver uses defaults that match what the projector itself emits.
        var prefix = Internal.ListNumberResolver.Resolve(paragraph, _doc!);
        if (string.IsNullOrEmpty(prefix)) return payload;
        if (!payload.StartsWith(prefix, StringComparison.Ordinal)) return payload;

        var after = payload.Substring(prefix.Length);
        if (after.Length > 0 && (after[0] == ' ' || after[0] == '\t' || after[0] == ' '))
            after = after.Substring(1);
        return after;
    }

    private static void ApplyReplaceTextAccept(XElement paragraph, IReadOnlyList<Internal.ParsedBlock> blocks)
    {
        var pPr = paragraph.Element(W.pPr);
        var markers = CapturePreservedMarkerPositions(paragraph);
        paragraph.RemoveNodes();
        if (pPr is not null) paragraph.Add(pPr);
        if (blocks.Count > 0)
            foreach (var run in blocks[0].RunElements)
                paragraph.Add(new XElement(run));
        RestorePreservedMarkerPositions(paragraph, markers);
    }

    private void ApplyReplaceTextTracked(XElement paragraph, IReadOnlyList<Internal.ParsedBlock> blocks)
    {
        var stamp = NewRevisionStamp();

        // Delete the live content in place — inside hyperlinks, fields, inline controls and
        // other authors' insertions (w:ins > w:del) — so a reject puts every run back exactly
        // where it was. The paragraph mark, the control shells and the zero-width markers
        // (bookmark/comment ranges as bare children; note and comment references as
        // marker-only runs, which must survive on BOTH accept and reject — issue B3) stay
        // where they sit rather than being lifted out and re-placed.
        MarkParagraphAsTrackedDeleted(paragraph, stamp, preserveParagraphMark: true, retainInlineStructure: true);

        if (blocks.Count == 0 || blocks[0].RunElements.Count == 0) return;
        var inserted = WrapPayloadAsInserted(blocks[0].RunElements, stamp);

        // The replacement precedes the paragraph's trailing markers — the references and range
        // ends after its last text-bearing run, wherever they sit, and the carrier that holds
        // them — so a note reference that closed the old sentence closes the new one too. With
        // no trailing marker it simply follows everything it supersedes.
        var lastText = paragraph.Descendants(W.r).LastOrDefault(run => !IsMarkerOnlyRun(run));
        var firstTrailingMarker = paragraph.Descendants()
            .Where(d => PreservedMarkerNames.Contains(d.Name) || IsMarkerOnlyRun(d))
            .FirstOrDefault(marker => lastText is null || marker.IsAfter(lastText));
        if (firstTrailingMarker is null)
            paragraph.Add(inserted);
        else
            firstTrailingMarker.AncestorsAndSelf().First(e => ReferenceEquals(e.Parent, paragraph))
                .AddBeforeSelf(inserted);
    }

    /// <summary>Wrap a parsed payload's inline elements as the session author's insertion.
    /// Contiguous runs share one <c>w:ins</c>; a hyperlink keeps its shell and owns its inserted
    /// runs (<c>w:hyperlink &gt; w:ins &gt; w:r</c>), the only nesting the schema allows and the
    /// shape Word writes — <c>w:ins &gt; w:hyperlink</c> fails validation.</summary>
    private List<XElement> WrapPayloadAsInserted(IReadOnlyList<XElement> runElements, RevisionStamp stamp)
    {
        var result = new List<XElement>();
        XElement? envelope = null;
        foreach (var element in runElements)
        {
            if (element.Name == W.r)
            {
                if (envelope is null) result.Add(envelope = CreateRevisionEnvelope(W.ins, stamp));
                envelope.Add(new XElement(element));
                continue;
            }
            envelope = null;
            var carrier = new XElement(element.Name, element.Attributes());
            carrier.Add(CreateRevisionEnvelope(W.ins, stamp,
                element.Elements().Select(child => new XElement(child)).ToArray()));
            result.Add(carrier);
        }
        return result;
    }

    /// <summary>
    /// Marks a whole paragraph as a tracked deletion: wraps live descendant runs in
    /// <c>w:del</c> AND marks the paragraph mark
    /// itself by adding <c>w:del</c> inside <c>w:pPr/w:rPr</c>. The combination tells
    /// Word the entire paragraph — content plus paragraph break — is a tracked deletion,
    /// so accepting the change actually removes the paragraph (instead of leaving an
    /// empty paragraph behind). The final pilcrow of a retained story/cell is preserved.
    /// </summary>
    private void MarkParagraphAsTrackedDeleted(XElement paragraph, RevisionStamp stamp,
        bool preserveParagraphMark = false, bool retainInlineStructure = false) =>
        MarkParagraphContentAndMark(paragraph, W.del, stamp.Author, stamp.Date, preserveParagraphMark, retainInlineStructure);

    /// <summary>
    /// Marks a whole table as a tracked deletion: every row gets a <c>w:trPr/w:del</c>
    /// marker (Word's row-deletion convention — there is no table-level "delete" markup),
    /// and every paragraph inside every cell is treated like
    /// <see cref="MarkParagraphAsTrackedDeleted"/>. Nested tables recurse.
    /// </summary>
    private void MarkTableAsTrackedDeleted(XElement table, RevisionStamp stamp)
    {
        foreach (var row in WordprocessingMLUtil.TableRows(table).ToList())
        {
            MarkRowAsTrackedRevision(row, inserted: false, stamp.Author, stamp.Date, markContent: false);

            foreach (var cell in row.Descendants(W.tc).Where(cell => cell.Ancestors(W.tr).First() == row).ToList())
            {
                foreach (var child in cell.Elements().ToList())
                    MarkTrackedStructuredContentChild(child, stamp);
            }
        }
    }

    /// <summary>
    /// Tracks deletion of a block <c>w:sdt</c> or <c>w:customXml</c> wrapper without
    /// discarding its ownership metadata. Two paired custom-XML deletion ranges cross
    /// the opening and closing tags, while every payload block receives its ordinary
    /// paragraph/table deletion markup. Accept therefore removes both wrapper and
    /// payload; reject restores the original wrapper and content. A content control's
    /// payload lives in <c>w:sdtContent</c>; a custom-XML wrapper is its own container,
    /// with <c>w:customXmlPr</c> as properties rather than payload.
    /// </summary>
    private void MarkStructuredBlockAsTrackedDeleted(XElement wrapper, RevisionStamp stamp)
    {
        var contentContainer = wrapper.Name == W.sdt
            ? wrapper.Element(W.sdtContent)
                ?? throw new InvalidOperationException("block w:sdt has no w:sdtContent")
            : wrapper.Name == W.customXml
                ? wrapper
                : throw new InvalidOperationException(
                    $"unsupported structured wrapper: {wrapper.Name}");

        foreach (var child in contentContainer.Elements()
            .Where(child => child.Name != W.customXmlPr).ToList())
            MarkTrackedStructuredContentChild(child, stamp);

        var boundaries = Internal.StructuredRevisionOps.AddCrossBoundaryMarkers(
            contentContainer,
            W.customXmlDelRangeStart,
            W.customXmlDelRangeEnd,
            name => CreateRevisionEnvelope(name, stamp));
        wrapper.AddBeforeSelf(boundaries.Before);
        wrapper.AddAfterSelf(boundaries.After);
    }

    /// <summary>
    /// Recursively marks one block payload node. Nested SDT and custom-XML wrappers receive their own
    /// reversible envelope; other transparent containers are preserved while their
    /// block-bearing descendants are marked.
    /// </summary>
    private void MarkTrackedStructuredContentChild(XElement child, RevisionStamp stamp)
    {
        if (child.Name == W.p)
        {
            MarkParagraphAsTrackedDeleted(child, stamp);
            return;
        }
        if (child.Name == W.tbl)
        {
            MarkTableAsTrackedDeleted(child, stamp);
            return;
        }
        if (child.Name == W.sdt || child.Name == W.customXml)
        {
            MarkStructuredBlockAsTrackedDeleted(child, stamp);
            return;
        }

        foreach (var nested in child.Elements().ToList())
            MarkTrackedStructuredContentChild(nested, stamp);
    }

    private void PromoteHyperlinkRelationships(XElement paragraph)
    {
        var owner = Internal.OwnedPartRelationships.FindOwner(_doc!, paragraph)
            ?? throw new InvalidOperationException("hyperlink paragraph has no owning package part");
        foreach (var link in paragraph.Descendants(W.hyperlink).ToList())
        {
            var hrefAttr = link.Attribute(Internal.MarkdownPayloadParser.HrefAttr);
            if (hrefAttr is null) continue;
            var url = hrefAttr.Value;
            if (url.StartsWith("#", StringComparison.Ordinal))
            {
                // Internal jumps are relationship-free OOXML. Writing r:id to an external
                // relationship whose URI is literally "#name" is invalid Word semantics (#469).
                link.SetAttributeValue(W.anchor, url.Substring(1));
                link.Attribute(R.id)?.Remove();
            }
            else
            {
                var relationship = Internal.OwnedPartRelationships.FindOrAddExternalHyperlink(
                    owner.Part, new Uri(url, UriKind.RelativeOrAbsolute));
                link.SetAttributeValue(R.id, relationship.Id);
                link.Attribute(W.anchor)?.Remove();
            }
            hrefAttr.Remove();
        }
    }

    private static void ApplyFormatToRun(XElement run, FormatOp op)
    {
        var rPr = run.Element(W.rPr);
        if (rPr is null) { rPr = new XElement(W.rPr); run.AddFirst(rPr); }

        static void Toggle(XElement rPr, XName name, bool? set)
        {
            if (set is null) return;
            var existing = rPr.Element(name);
            if (set.Value)
            {
                // Turn the property ON. A run may already carry an explicit OFF element
                // (e.g. Google Docs stamps <w:b w:val="0"/> on every run); just adding a new
                // element when one is "missing" would leave that w:val="0" in place and the
                // toggle would silently do nothing. Normalize: drop the w:val so the bare
                // element (<w:b/>) means on; add one only when truly absent.
                if (existing is null)
                    WordprocessingMLUtil.InsertRPrChildInOrder(rPr, new XElement(name));
                else existing.Attribute(W.val)?.Remove();
            }
            else existing?.Remove();
        }

        Toggle(rPr, W.b, op.Bold);
        Toggle(rPr, W.i, op.Italic);
        Toggle(rPr, W.strike, op.Strike);

        if (op.Underline is true)
        {
            rPr.Element(W.u)?.Remove();
            WordprocessingMLUtil.InsertRPrChildInOrder(
                rPr, new XElement(W.u, new XAttribute(W.val, "single")));
        }
        else if (op.Underline is false) rPr.Element(W.u)?.Remove();

        if (op.Code is true)
        {
            rPr.Element(W.rStyle)?.Remove();
            WordprocessingMLUtil.InsertRPrChildInOrder(
                rPr, new XElement(W.rStyle, new XAttribute(W.val, "Code")));
        }
        else if (op.Code is false) rPr.Element(W.rStyle)?.Remove();

        if (op.Color is not null)
        {
            rPr.Element(W.color)?.Remove();
            if (op.Color.Length > 0)
                WordprocessingMLUtil.InsertRPrChildInOrder(
                    rPr, new XElement(W.color, new XAttribute(W.val, op.Color)));
        }

        if (op.RunStyle is not null)
        {
            rPr.Element(W.rStyle)?.Remove();
            if (op.RunStyle.Length > 0)
                WordprocessingMLUtil.InsertRPrChildInOrder(
                    rPr, new XElement(W.rStyle, new XAttribute(W.val, op.RunStyle)));
        }

        if (op.VertAlign is not null)
        {
            rPr.Element(W.vertAlign)?.Remove();
            var v = op.VertAlign switch
            {
                "super" => "superscript",
                "sub" => "subscript",
                "none" or "baseline" => "",
                _ => op.VertAlign,
            };
            if (v.Length > 0)
            {
                if (v is not ("superscript" or "subscript"))
                    throw new ArgumentException($"invalid vertAlign: {op.VertAlign}");
                WordprocessingMLUtil.InsertRPrChildInOrder(
                    rPr, new XElement(W.vertAlign, new XAttribute(W.val, v)));
            }
        }

        if (op.FontSizePts is { } pts)
        {
            // w:sz / w:szCs are half-points. Clearing (<= 0) drops the explicit size so the run
            // inherits the style/default size again.
            rPr.Element(W.sz)?.Remove();
            rPr.Element(W.szCs)?.Remove();
            if (pts > 0)
            {
                var halfPts = ((int)System.Math.Round(pts * 2, System.MidpointRounding.AwayFromZero))
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);
                WordprocessingMLUtil.InsertRPrChildInOrder(
                    rPr, new XElement(W.sz, new XAttribute(W.val, halfPts)));
                WordprocessingMLUtil.InsertRPrChildInOrder(
                    rPr, new XElement(W.szCs, new XAttribute(W.val, halfPts)));
            }
        }

        if (op.FontFamily is not null)
        {
            // "" clears the explicit font so the run inherits the style/default.
            rPr.Element(W.rFonts)?.Remove();
            if (op.FontFamily.Length > 0)
            {
                var rFonts = new XElement(W.rFonts,
                    new XAttribute(W.ascii, op.FontFamily),
                    new XAttribute(W.hAnsi, op.FontFamily),
                    new XAttribute(W.cs, op.FontFamily));
                WordprocessingMLUtil.InsertRPrChildInOrder(rPr, rFonts);
            }
        }

        if (op.Highlight is not null)
        {
            // w:highlight is ST_HighlightColor — a closed set of Word swatch names, not a hex
            // value (that would be w:shd). "" / "none" clears; anything outside the set is
            // refused rather than written, because Word ignores an unknown token silently.
            rPr.Element(W.highlight)?.Remove();
            var name = op.Highlight.Trim();
            if (name.Length > 0 && !string.Equals(name, "none", StringComparison.OrdinalIgnoreCase))
            {
                var canonical = CanonicalHighlightName(name)
                    ?? throw new ArgumentException($"invalid highlight: {op.Highlight}");
                WordprocessingMLUtil.InsertRPrChildInOrder(
                    rPr, new XElement(W.highlight, new XAttribute(W.val, canonical)));
            }
        }

        // Caps and small caps are one either/or slot in Word's UI: turning one on removes the
        // other, exactly as the Font dialog does. An op that sets BOTH true is resolved in
        // declaration order, so SmallCaps wins.
        if (op.Caps is true) rPr.Element(W.smallCaps)?.Remove();
        Toggle(rPr, W.caps, op.Caps);
        if (op.SmallCaps is true) rPr.Element(W.caps)?.Remove();
        Toggle(rPr, W.smallCaps, op.SmallCaps);
    }

    /// <summary>The sixteen <c>ST_HighlightColor</c> swatch names Word accepts on <c>w:highlight</c>,
    /// in their canonical casing (the attribute is case-sensitive in Word).</summary>
    private static readonly string[] HighlightColorNames =
    {
        "yellow", "green", "cyan", "magenta", "blue", "red",
        "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow",
        "darkGray", "lightGray", "black", "white",
    };

    /// <summary>Map a caller-supplied highlight token to its canonical swatch name (matching
    /// case-insensitively, so <c>darkblue</c> and <c>DarkBlue</c> both write <c>darkBlue</c>), or
    /// null when it is not a Word highlight colour.</summary>
    private static string? CanonicalHighlightName(string name)
    {
        foreach (var candidate in HighlightColorNames)
            if (string.Equals(candidate, name, StringComparison.OrdinalIgnoreCase)) return candidate;
        return null;
    }

    /// <summary>
    /// Apply a run-format mutation using Word's native tracked-format representation:
    /// the run keeps its new properties while <c>w:rPr/w:rPrChange/w:rPr</c> stores
    /// the old properties for reject. A run may carry only one direct rPrChange. When
    /// it already has one, preserve that marker (including its original baseline and
    /// attribution) and fold the new formatting into the same pending revision; replacing
    /// its baseline with the intermediate format would make reject-all stop halfway.
    /// </summary>
    private void ApplyFormatToRunTracked(
        XElement run, FormatOp op, string revisionAuthor, string revisionDate)
    {
        var originalRPr = run.Element(W.rPr);
        var originalRPrClone = originalRPr is null ? null : new XElement(originalRPr);
        var oldProperties = SnapshotRunPropertiesForRevision(originalRPr);

        // rPrChange is the final CT_RPr child. Detach it while ApplyFormatToRun edits
        // properties, then reinsert it once below. Taking only the first also prevents
        // malformed duplicate markers from becoming nested/stacked.
        var existingChanges = originalRPr?.Elements(W.rPrChange).ToList()
            ?? new List<XElement>();
        foreach (var change in existingChanges) change.Remove();
        var existingChange = existingChanges.FirstOrDefault();
        var archivedProperties = existingChange is null
            ? null
            : SnapshotRunPropertiesForRevision(existingChange.Element(W.rPr));

        try
        {
            ApplyFormatToRun(run, op);
        }
        catch
        {
            RestoreRunProperties(run, originalRPrClone);
            throw;
        }

        var currentRPr = run.Element(W.rPr)!;
        var newProperties = SnapshotRunPropertiesForRevision(currentRPr);
        bool changed = !RunPropertiesEquivalentForRevision(oldProperties, newProperties);

        if (archivedProperties is not null
            && RunPropertiesEquivalentForRevision(archivedProperties, newProperties))
        {
            // Editing a pending format change back to its archived baseline resolves
            // that portion of the change. Keeping rPrChange here would leave a phantom
            // revision whose accept and reject results are identical. Reuse the stored
            // baseline XML so lexical-only differences do not survive as document churn.
            RestoreRunProperties(run,
                archivedProperties.HasElements
                    || archivedProperties.Attributes().Any(a => !a.IsNamespaceDeclaration)
                    ? archivedProperties
                    : null);
            return;
        }

        if (!changed)
        {
            // A no-op must not manufacture a format revision OR normalize/reorder the
            // caller's existing XML as a side effect. Put the exact rPr back.
            RestoreRunProperties(run, originalRPrClone);
            return;
        }

        if (existingChange is not null)
        {
            WordprocessingMLUtil.InsertRPrChildInOrder(currentRPr, existingChange);
            return;
        }

        WordprocessingMLUtil.InsertRPrChildInOrder(currentRPr,
            CreateRevisionEnvelope(
                W.rPrChange, revisionAuthor, revisionDate, oldProperties));
    }

    /// <summary>
    /// Clone the direct run properties suitable for the inner payload of rPrChange.
    /// Existing change markup is deliberately excluded (CT_RPrOriginal cannot contain
    /// another rPrChange), as is projector-only Unid bookkeeping.
    /// </summary>
    private static XElement SnapshotRunPropertiesForRevision(XElement? rPr)
    {
        if (rPr is null) return new XElement(W.rPr);

        var snapshot = new XElement(rPr);
        snapshot.Descendants(W.rPrChange).Remove();
        foreach (var element in snapshot.DescendantsAndSelf())
            element.Attributes()
                .Where(a => a.Name.Namespace == PtOpenXml.pt)
                .Remove();
        return snapshot;
    }

    /// <summary>Compare run properties by their schema order and normalize the lexical
    /// true spellings of the on/off toggles ApplyFormat writes as bare elements, bare
    /// underline as <c>single</c>, and canonical lexical forms for color/half-point
    /// values. This keeps semantically identical writes (for example
    /// <c>w:b w:val="true"</c> → <c>w:b</c>, <c>w:u</c> →
    /// <c>w:u w:val="single"</c>, or remove/re-add of an unchanged color) out of the
    /// review pane.</summary>
    private static bool RunPropertiesEquivalentForRevision(XElement left, XElement right)
    {
        static XElement Normalize(XElement source)
        {
            var ordered = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(source);
            foreach (var name in new[] { W.b, W.i, W.strike })
            {
                foreach (var property in ordered.Elements(name).ToList())
                {
                    var value = (string?)property.Attribute(W.val);
                    if (value is null || value is "1"
                        || value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("on", StringComparison.OrdinalIgnoreCase))
                    {
                        property.Attribute(W.val)?.Remove();
                    }
                    else if (value is "0"
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("off", StringComparison.OrdinalIgnoreCase))
                    {
                        property.Remove();
                    }
                }
            }

            foreach (var underline in ordered.Elements(W.u).ToList())
            {
                var value = (string?)underline.Attribute(W.val);
                if (string.IsNullOrEmpty(value)
                    || value.Equals("single", StringComparison.OrdinalIgnoreCase))
                {
                    underline.SetAttributeValue(W.val, "single");
                }
                else if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    underline.Remove();
                }
            }

            foreach (var vertAlign in ordered.Elements(W.vertAlign).ToList())
            {
                var value = (string?)vertAlign.Attribute(W.val);
                if (value is not null
                    && value.Equals("baseline", StringComparison.OrdinalIgnoreCase))
                {
                    vertAlign.Remove();
                }
            }

            foreach (var color in ordered.Elements(W.color))
            {
                var value = (string?)color.Attribute(W.val);
                if (value is null) continue;
                if (value.Equals("auto", StringComparison.OrdinalIgnoreCase))
                    color.SetAttributeValue(W.val, "auto");
                else if (value.Length == 6 && value.All(Uri.IsHexDigit))
                    color.SetAttributeValue(W.val, value.ToUpperInvariant());
            }

            foreach (var name in new[] { W.sz, W.szCs })
            {
                foreach (var size in ordered.Elements(name))
                {
                    var value = (string?)size.Attribute(W.val);
                    if (uint.TryParse(value, System.Globalization.NumberStyles.None,
                            System.Globalization.CultureInfo.InvariantCulture, out var parsed))
                    {
                        size.SetAttributeValue(W.val, parsed.ToString(
                            System.Globalization.CultureInfo.InvariantCulture));
                    }
                }
            }
            return ordered;
        }

        return XNode.DeepEquals(Normalize(left), Normalize(right));
    }

    private static void RestoreRunProperties(XElement run, XElement? snapshot)
    {
        run.Element(W.rPr)?.Remove();
        if (snapshot is not null) run.AddFirst(snapshot);
    }

    internal XElement BuildParagraphFromParsedBlock(Internal.ParsedBlock block)
    {
        var p = new XElement(W.p);
        var pPr = new XElement(W.pPr);

        switch (block.Kind)
        {
            case Internal.ParserBlockKind.Heading1:
            case Internal.ParserBlockKind.Heading2:
            case Internal.ParserBlockKind.Heading3:
            case Internal.ParserBlockKind.Heading4:
            case Internal.ParserBlockKind.Heading5:
            case Internal.ParserBlockKind.Heading6:
                {
                    int level = (int)block.Kind - (int)Internal.ParserBlockKind.Heading1 + 1;
                    var styleId = $"Heading{level}";
                    pPr.Add(new XElement(W.pStyle, new XAttribute(W.val, styleId)));
                    if (HeadingNumberingSuppressor(styleId) is { } suppressor)
                        pPr.Add(suppressor);
                    break;
                }
            case Internal.ParserBlockKind.Quote:
                pPr.Add(new XElement(W.pStyle, new XAttribute(W.val, "Quote")));
                break;
            case Internal.ParserBlockKind.Code:
                pPr.Add(new XElement(W.pStyle, new XAttribute(W.val, "Code")));
                break;
            // List items declare numbering, not a style. The builder has no view of the
            // payload's other blocks or the insertion point's neighbours, both of which decide
            // which w:num the item joins, so the caller assigns it through
            // AssignPayloadListNumbering.
        }

        if (pPr.HasElements) p.Add(pPr);
        foreach (var run in block.RunElements)
            p.Add(new XElement(run));
        return p;
    }

    /// <summary>
    /// The `numId=0` numbering suppressor a markdown-authored heading needs — but ONLY
    /// when the document's heading style actually attaches numbering (a legal-outline
    /// template), where it stops the inherited prefix from changing the authored text.
    /// Written unconditionally it made every markdown heading diff as
    /// FormatChanged(numId, numLevel) against an identical Style-dropdown heading, whose
    /// `pPr` carries no `numPr` at all (#572). Null when the style (or its basedOn chain)
    /// does not number, so all authoring paths produce the same paragraph mark.
    /// </summary>
    private XElement? HeadingNumberingSuppressor(string styleId) =>
        _doc is not null && Internal.StyleFactory.StyleAttachesNumbering(_doc, styleId)
            ? new XElement(W.numPr,
                new XElement(W.ilvl, new XAttribute(W.val, 0)),
                new XElement(W.numId, new XAttribute(W.val, 0)))
            : null;

    internal static string ParserBlockKindToAnchorKind(Internal.ParserBlockKind kind) => kind switch
    {
        Internal.ParserBlockKind.Heading1
            or Internal.ParserBlockKind.Heading2
            or Internal.ParserBlockKind.Heading3
            or Internal.ParserBlockKind.Heading4
            or Internal.ParserBlockKind.Heading5
            or Internal.ParserBlockKind.Heading6 => "h",
        Internal.ParserBlockKind.BulletItem
            or Internal.ParserBlockKind.OrderedItem => "li",
        _ => "p",
    };

    /// <summary>
    /// Mirror the classifier used by <see cref="WmlToMarkdownConverter"/> so the kind
    /// reported in <see cref="EditResult.Created"/> matches what the projector will
    /// emit on the next <see cref="DocxSession.Project"/>. If we used the parser's
    /// kind blindly, a bullet-payload paragraph without a <c>w:numPr</c> would be
    /// reported as "li" but appear as "p" in the projection — a stale anchor id.
    /// </summary>
    internal static string ClassifyParagraphKind(XElement paragraph)
    {
        var pPr = paragraph.Element(W.pPr);
        var styleId = (string?)pPr?.Element(W.pStyle)?.Attribute(W.val);
        if (!string.IsNullOrEmpty(styleId)
            && (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
                || styleId.Equals("Title", StringComparison.OrdinalIgnoreCase)
                || styleId.Equals("Subtitle", StringComparison.OrdinalIgnoreCase)))
            return "h";
        var directNumId = (int?)pPr?.Element(W.numPr)?.Element(W.numId)?.Attribute(W.val);
        if (directNumId is not null && directNumId != 0) return "li";
        return "p";
    }

    /// <summary>
    /// Classify any block-level XElement to the kind used in anchor ids. Mirrors
    /// the kinds the projector emits — paragraphs go through
    /// <see cref="ClassifyParagraphKind"/>; tables/rows/cells map to their fixed kinds.
    /// Falls back to "p" for unknown shapes.
    /// </summary>
    internal static string ClassifyBlockKind(XElement element)
    {
        if (element.Name == W.p) return ClassifyParagraphKind(element);
        if (element.Name == W.tbl) return "tbl";
        if (element.Name == W.tr) return "tr";
        if (element.Name == W.tc) return "tc";
        return "p";
    }

    /// <summary>
    /// What one markdown payload's list blocks have decided so far. Consecutive ordered items
    /// form one list and share one <c>w:num</c> instance that starts where the first marker
    /// says; a non-list block or a top-level bullet ends that list, so a later ordered list
    /// restarts on its own instance, as Word does for separate lists. Bullets have no sequence
    /// to keep apart and share the document's Docxodus bullet definition. The payload's first
    /// list continues an adjacent, format-compatible list item instead of starting its own —
    /// the block before the insertion point, or, when the whole payload is one list, the block
    /// after it.
    /// </summary>
    private sealed class PayloadListState
    {
        public PayloadListState(IReadOnlyList<Internal.ParsedBlock> blocks)
        {
            var topLevelKinds = blocks
                .Where(block => block.ListLevel == 0)
                .Select(block => block.Kind)
                .Distinct()
                .ToList();
            SingleList = blocks.All(block => block.Kind
                    is Internal.ParserBlockKind.BulletItem or Internal.ParserBlockKind.OrderedItem)
                && topLevelKinds.Count == 1;
        }

        public XElement? PrecedingNeighbor { get; init; }
        public XElement? FollowingNeighbor { get; init; }
        public bool SingleList { get; }
        public bool FirstListPending { get; set; } = true;
        public int? OrderedNumId { get; set; }
    }

    /// <summary>
    /// Give a paragraph built from a markdown list block the native numbering it declares,
    /// through the same owner <see cref="ApplyListFormat"/> uses. Non-list blocks end the
    /// current ordered list and get nothing. Nesting maps the indent level to <c>w:ilvl</c>;
    /// a document's own list that is continued has any missing level synthesized.
    /// </summary>
    private void AssignPayloadListNumbering(
        XElement paragraph, Internal.ParsedBlock block, PayloadListState state)
    {
        bool ordered = block.Kind == Internal.ParserBlockKind.OrderedItem;
        if (!ordered && block.Kind != Internal.ParserBlockKind.BulletItem)
        {
            state.OrderedNumId = null;
            return;
        }

        int ilvl = Math.Clamp(block.ListLevel, 0, 8);
        if (!ordered && ilvl == 0) state.OrderedNumId = null;

        int? numId = null;
        if (state.FirstListPending && ilvl == 0)
        {
            state.FirstListPending = false;
            numId = CompatibleNeighborNumbering(state.PrecedingNeighbor, ordered)
                ?? (state.SingleList
                    ? CompatibleNeighborNumbering(state.FollowingNeighbor, ordered)
                    : null);
            if (numId is { } continued && ordered) state.OrderedNumId = continued;
        }

        if (numId is null && ordered)
        {
            state.OrderedNumId ??= Internal.NumberingFactory.CreateNumberingInstance(
                _doc!, ListFormat.Decimal, block.ListStart);
            numId = state.OrderedNumId;
        }
        numId ??= Internal.NumberingFactory.EnsureNumbering(_doc!, ListFormat.Bullet);
        if (ilvl > 0) Internal.NumberingFactory.EnsureLevelDefined(_doc!, numId.Value, ilvl);

        var pPr = paragraph.Element(W.pPr);
        if (pPr is null) { pPr = new XElement(W.pPr); paragraph.AddFirst(pPr); }
        pPr.Element(W.numPr)?.Remove();
        SetPPrChildInOrder(pPr, new XElement(W.numPr,
            new XElement(W.ilvl, new XAttribute(W.val, ilvl)),
            new XElement(W.numId, new XAttribute(W.val, numId.Value))));
    }

    /// <summary>The numbering instance of <paramref name="neighbor"/> when it is a list item
    /// whose level renders in the same family as the block (bullet vs numbered); else null.</summary>
    private int? CompatibleNeighborNumbering(XElement? neighbor, bool ordered)
    {
        if (neighbor is null || neighbor.Name != W.p) return null;
        var numPr = neighbor.Element(W.pPr)?.Element(W.numPr);
        var numId = (int?)numPr?.Element(W.numId)?.Attribute(W.val);
        if (numId is null or 0) return null;
        int ilvl = (int?)numPr!.Element(W.ilvl)?.Attribute(W.val) ?? 0;
        var format = Internal.NumberingFactory.ResolveNumberFormat(_doc!, numId.Value, ilvl);
        if (format is null) return null;
        return (format == NumberFormat.Bullet) == !ordered ? numId : null;
    }

    // Top-level inline children of <w:p> that participate in text flow.
    // Hyperlinks, sdts, fldSimple and smartTag are transparent containers — their
    // descendant runs contribute to the paragraph's visible text. Bookmark/comment
    // markers (zero-width) are tracked separately and not enumerated here.
    //
    // w:ins and w:moveTo are transparent for the same reason: an insertion's text is
    // present in the document as ordinary w:t and is exactly what a reader sees, so it
    // belongs in the flat text every offset-addressed op works over. Their deleting
    // counterparts w:del and w:moveFrom deliberately are NOT here — that content is
    // w:delText, which RunText does not read, so deleted text stays out of the visible
    // string and offsets keep addressing what the document actually shows.
    private static readonly HashSet<XName> InlineContainerNames = new()
    {
        W.hyperlink, W.sdt, W.fldSimple, W.smartTag, W.ins, W.moveTo,
    };

    private static bool IsInlineChild(XElement e) =>
        e.Name == W.r || InlineContainerNames.Contains(e.Name);

    /// <summary>
    /// All <c>&lt;w:r&gt;</c> elements that contribute to the paragraph's visible text,
    /// in document order — including runs nested inside hyperlinks, sdts, fldSimple,
    /// smartTags, and tracked insertions (<c>w:ins</c>/<c>w:moveTo</c>). Iterating only
    /// <c>Elements(W.r)</c> silently skips hyperlink-internal runs, which produced the
    /// bugs documented in DS080-DS090; skipping insertion-internal runs made an agent
    /// unable to re-find its own tracked edit (DS409-DS410).
    /// </summary>
    internal static IEnumerable<XElement> InlineRuns(XElement paragraph)
    {
        foreach (var child in paragraph.Elements())
        {
            if (child.Name == W.r) yield return child;
            else if (InlineContainerNames.Contains(child.Name))
                foreach (var run in child.Descendants(W.r))
                    yield return run;
        }
    }

    internal static string ParagraphText(XElement paragraph) =>
        string.Concat(InlineRuns(paragraph).Select(RunText));

    internal static string RunText(XElement run) =>
        string.Concat(run.Elements(W.t).Select(t => (string)t));

    private static int InlineChildTextLength(XElement child) =>
        string.Concat(child.DescendantsAndSelf(W.t).Select(t => (string)t)).Length;

    /// <summary>
    /// Whether inserting a new top-level paragraph child at <paramref name="offset"/> would
    /// silently move it away from the requested visible-text position. Plain runs are split by
    /// <see cref="SplitRunsAtOffset"/>. A top-level hyperlink containing direct runs is split by
    /// <see cref="SplitInlineContainersAtOffset"/>. Every other container — including
    /// <c>w:ins</c>/<c>w:moveTo</c> — is atomic because splitting it requires revision- or
    /// field-specific semantics.
    /// </summary>
    private static bool HasUnsupportedInlineInsertionBoundary(XElement paragraph, int offset)
    {
        int consumed = 0;
        foreach (var child in paragraph.Elements().Where(IsInlineChild))
        {
            int length = InlineChildTextLength(child);
            if (consumed < offset && offset < consumed + length)
            {
                if (child.Name == W.r) return false;
                if (child.Name != W.hyperlink) return true;

                int localOffset = offset - consumed;
                int hyperlinkConsumed = 0;
                foreach (var nested in child.Elements().Where(IsInlineChild))
                {
                    int nestedLength = InlineChildTextLength(nested);
                    if (hyperlinkConsumed < localOffset
                        && localOffset < hyperlinkConsumed + nestedLength)
                        return nested.Name != W.r;
                    hyperlinkConsumed += nestedLength;
                }
                return false;
            }
            consumed += length;
        }
        return false;
    }

    /// <summary>
    /// If a run straddles <paramref name="offset"/>, split it into two adjacent runs
    /// at that offset. Walks runs inside hyperlinks/sdts/etc. too, so the boundary
    /// is clean regardless of which container the run lives in. The new sibling run
    /// is inserted into the same parent as the original (preserving hyperlink/sdt
    /// membership for the keep-half).
    /// </summary>
    internal static void SplitRunsAtOffset(XElement paragraph, int offset)
    {
        int consumed = 0;
        foreach (var run in InlineRuns(paragraph).ToList())
        {
            var runText = RunText(run);
            if (consumed == offset) return;
            if (consumed + runText.Length <= offset) { consumed += runText.Length; continue; }
            int splitAt = offset - consumed;
            if (splitAt <= 0) return;

            var keep = runText.Substring(0, splitAt);
            var move = runText.Substring(splitAt);

            foreach (var t in run.Elements(W.t).ToList()) t.Remove();
            run.Add(new XElement(W.t,
                new XAttribute(XNamespace.Xml + "space", "preserve"), keep));

            var rPr = run.Element(W.rPr);
            var newRun = new XElement(W.r);
            if (rPr is not null) newRun.Add(new XElement(rPr));
            newRun.Add(new XElement(W.t,
                new XAttribute(XNamespace.Xml + "space", "preserve"), move));
            run.AddAfterSelf(newRun);
            return;
        }
    }

    /// <summary>
    /// Ensures no top-level inline child straddles <paramref name="offset"/>: if a
    /// hyperlink (or other splittable container) crosses the boundary, it's split
    /// into two sibling containers sharing the same attributes (e.g. <c>r:id</c>),
    /// each holding half the runs. After this call, <see cref="MoveInlineChildrenAfter"/>
    /// can move whole-child elements without slicing through anything.
    /// </summary>
    internal static void SplitInlineContainersAtOffset(XElement paragraph, int offset)
    {
        int consumed = 0;
        foreach (var child in paragraph.Elements().Where(IsInlineChild).ToList())
        {
            int len = InlineChildTextLength(child);
            if (consumed + len <= offset) { consumed += len; continue; }
            if (consumed == offset) return; // boundary already clean
            int local = offset - consumed;

            if (child.Name == W.hyperlink)
                SplitHyperlinkAt(child, local);
            // For <w:r>: SplitRunsAtOffset already handled it. For sdt/fldSimple/smartTag and
            // revision wrappers: treat as atomic — callers that insert a top-level child must
            // reject an interior boundary with HasUnsupportedInlineInsertionBoundary before
            // reaching this helper.
            return;
        }
    }

    private static void SplitHyperlinkAt(XElement hyperlink, int localOffset)
    {
        // Split runs inside the hyperlink at the local offset (works because SplitRunsAtOffset
        // walks descendants through container types).
        SplitRunsAtOffset(hyperlink, localOffset);

        int consumed = 0;
        var movedChildren = new List<XElement>();
        foreach (var child in hyperlink.Elements().ToList())
        {
            int len = IsInlineChild(child) ? InlineChildTextLength(child) : 0;
            if (consumed >= localOffset) movedChildren.Add(child);
            consumed += len;
        }
        if (movedChildren.Count == 0) return;

        var newLink = new XElement(W.hyperlink);
        // Every attribute EXCEPT the Unid: that one is this hyperlink's identity, and the public
        // hl:<scope>:<unid> id is derived from it. Cloning it would give the two halves the same id,
        // so FindHyperlinkElement would match only the first and Update/RemoveHyperlink would
        // silently act on half the link.
        foreach (var a in hyperlink.Attributes())
            if (a.Name != PtOpenXml.Unid) newLink.SetAttributeValue(a.Name, a.Value);
        newLink.SetAttributeValue(PtOpenXml.Unid, UnidHelper.GenerateUnid());
        foreach (var child in movedChildren) { child.Remove(); newLink.Add(child); }
        hyperlink.AddAfterSelf(newLink);
    }

    /// <summary>
    /// Move every paragraph child (inline run/container OR zero-width marker)
    /// whose position is at or past <paramref name="offset"/> from
    /// <paramref name="paragraph"/> into <paramref name="destination"/>. Inline
    /// children advance the position counter by their text length; markers
    /// (bookmarkStart/End, comment range markers, etc.) advance it by 0 and so
    /// inherit the position they're sandwiched between.
    /// </summary>
    internal static void MoveInlineChildrenAfter(XElement paragraph, int offset, XElement destination)
    {
        int consumed = 0;
        var toMove = new List<XElement>();
        foreach (var child in paragraph.Elements().ToList())
        {
            if (child.Name == W.pPr) continue;
            int len = IsInlineChild(child) ? InlineChildTextLength(child) : 0;
            if (consumed >= offset) toMove.Add(child);
            consumed += len;
        }
        foreach (var c in toMove) { c.Remove(); destination.Add(c); }
    }
}
