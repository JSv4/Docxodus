// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;

namespace Docxodus.Verification;

/// <summary>Options for the stable semantic comparison surface.</summary>
public sealed record SemanticDiffOptions
{
    /// <summary>The existing IR comparison settings. A fresh default instance is used when omitted.</summary>
    public DocxDiffSettings? DiffSettings { get; init; }

    /// <summary>
    /// Include relationship, media, revision, bookmark, annotation, and unknown-part changes that are
    /// outside the modeled IR. Enabled by default so unmodeled package edits cannot disappear silently.
    /// </summary>
    public bool IncludePackageChanges { get; init; } = true;

    /// <summary>
    /// Shared package-manifest safety policy used for the mandatory preflight and, when enabled,
    /// the package-level semantic supplement.
    /// </summary>
    public PackageManifestOptions PackageOptions { get; init; } = new()
    {
        MaxEntryCount = 10_000,
        MaxEntryUncompressedBytes = 64L * 1024 * 1024,
        MaxTotalUncompressedBytes = 256L * 1024 * 1024,
        MaxXmlPartBytes = 64L * 1024 * 1024,
        MaxCompressionRatio = 1_000,
        MaxUriLength = 2_048,
    };
}

/// <summary>
/// Public, stable semantic DOCX comparison. The existing <see cref="DocxDiff"/> IR and edit script
/// remain the alignment authority, while this type projects that internal data into a versioned audit
/// schema and supplements it with a narrow package-level detector for facts the IR does not model.
/// </summary>
public static class SemanticDiff
{
    /// <summary>
    /// Compare raw DOCX bytes after applying the bounded package preflight, before constructing an
    /// Open XML SDK document. Prefer this overload at byte-oriented trust boundaries.
    /// </summary>
    public static SemanticChangeSet Compare(
        byte[] leftBytes,
        byte[] rightBytes,
        SemanticDiffOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(leftBytes);
        ArgumentNullException.ThrowIfNull(rightBytes);
        if (leftBytes.Length == 0)
            throw new ArgumentException("No left document data provided", nameof(leftBytes));
        if (rightBytes.Length == 0)
            throw new ArgumentException("No right document data provided", nameof(rightBytes));
        return SemanticDiffEngine.Compare(leftBytes, rightBytes, options ?? new SemanticDiffOptions());
    }

    public static SemanticChangeSet Compare(
        WmlDocument left,
        WmlDocument right,
        SemanticDiffOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);
        return SemanticDiffEngine.Compare(left, right, options ?? new SemanticDiffOptions());
    }

    public static string CompareJson(
        byte[] leftBytes,
        byte[] rightBytes,
        SemanticDiffOptions? options = null,
        bool indented = true) => Compare(leftBytes, rightBytes, options).ToJson(indented);

    public static string CompareJson(
        WmlDocument left,
        WmlDocument right,
        SemanticDiffOptions? options = null,
        bool indented = true) => Compare(left, right, options).ToJson(indented);
}

internal sealed record SemanticChangeDraft(
    SemanticChangeOperation Operation,
    SemanticChangeFamily Family,
    string PartUri,
    string Path,
    string? LeftAnchor,
    string? RightAnchor,
    string? LeftScope,
    string? RightScope,
    string? MoveId,
    SemanticValue Before,
    SemanticValue After,
    // Set only on package-detector Move drafts: both sides' full location keys, so the engine
    // can ask the IR alignment whether the "relocation" is just the containing block's
    // content-derived Unid changing in place. Never serialized into the public schema.
    string? PackageLocationBefore = null,
    string? PackageLocationAfter = null);

/// <summary>
/// Keeps projection from the shared package-manifest inspection separate from the public semantic
/// schema and the IR projection.
/// </summary>
internal interface ISemanticPackageChangeDetector
{
    IReadOnlyList<SemanticChangeDraft> Compare(
        byte[] leftBytes,
        byte[] rightBytes,
        SemanticDiffOptions options);
}

internal static class SemanticDiffEngine
{
    private static readonly IrReaderOptions ReadOptions = new()
    {
        RetainSources = true,
        RevisionView = RevisionView.Accept,
    };

    private static readonly ISemanticPackageChangeDetector PackageDetector =
        new OpcSemanticPackageChangeDetector();

    public static SemanticChangeSet Compare(
        byte[] leftBytes,
        byte[] rightBytes,
        SemanticDiffOptions options)
    {
        // This overload owns the raw-byte trust boundary: do not construct WmlDocument (whose
        // constructor opens the OPC package to identify its document type) until inspection has
        // enforced all declared archive and XML limits.
        var packageChanges = PackageDetector.Compare(leftBytes, rightBytes, options);
        var left = new WmlDocument("left.docx", leftBytes);
        var right = new WmlDocument("right.docx", rightBytes);
        return CompareValidated(left, right, options, packageChanges);
    }

    public static SemanticChangeSet Compare(
        WmlDocument originalLeft,
        WmlDocument originalRight,
        SemanticDiffOptions options)
    {
        // Inspect the raw package before the Open XML SDK/IR path. This makes the declared package
        // limits an actual boundary for the default public operation instead of a late supplement
        // reached only after an untrusted archive has already been expanded.
        var packageChanges = PackageDetector.Compare(
            originalLeft.DocumentByteArray,
            originalRight.DocumentByteArray,
            options);
        return CompareValidated(originalLeft, originalRight, options, packageChanges);
    }

    private static SemanticChangeSet CompareValidated(
        WmlDocument originalLeft,
        WmlDocument originalRight,
        SemanticDiffOptions options,
        IReadOnlyList<SemanticChangeDraft> packageChanges)
    {
        var settings = options.DiffSettings ?? new DocxDiffSettings();
        var left = PreAccept(originalLeft, settings);
        var right = PreAccept(originalRight, settings);
        var diffSettings = settings.ToIrDiffSettings() with { CrossParagraphTokenDiff = false };
        var leftIr = IrReader.Read(left, ReadOptions);
        var rightIr = IrReader.Read(right, ReadOptions);
        var script = IrEditScriptBuilder.Build(leftIr, rightIr, diffSettings);

        var drafts = new List<SemanticChangeDraft>();
        var projector = new Projector(leftIr, rightIr, diffSettings, drafts);
        projector.Project(script);
        projector.CompareRegistries();
        projector.CompareComments();

        // The IR edit script is the alignment authority: a package fact whose "new location" is
        // the same aligned block (its content-derived Unid merely re-hashed under a text edit or
        // duplicate-shift) has not moved, so its Move record is dropped rather than published.
        drafts.AddRange(packageChanges.Where(draft =>
            !IsAlignedRelocation(draft, projector.AlignedBlockIdentities)));

        var changes = drafts.Select((draft, index) => new SemanticChange
        {
            Id = $"chg-{index + 1:D6}",
            Operation = draft.Operation,
            Family = draft.Family,
            PartUri = draft.PartUri,
            Path = draft.Path,
            LeftAnchor = draft.LeftAnchor,
            RightAnchor = draft.RightAnchor,
            LeftScope = draft.LeftScope,
            RightScope = draft.RightScope,
            MoveId = draft.MoveId,
            Before = draft.Before,
            After = draft.After,
        }).ToArray();
        return new SemanticChangeSet(changes);
    }

    private static WmlDocument PreAccept(WmlDocument document, DocxDiffSettings settings) =>
        settings.PreAcceptInputRevisions && !settings.PreserveInputRevisions
            ? RevisionProcessor.AcceptRevisions(document)
            : document;

    /// <summary>
    /// True when a package-detector Move draft's before-location equals its after-location once
    /// every anchor component is resolved through the IR block alignment — i.e. the fact sits in
    /// the same aligned block on both sides and only that block's content hash changed.
    /// </summary>
    private static bool IsAlignedRelocation(
        SemanticChangeDraft draft,
        IReadOnlyDictionary<string, string> alignedBlockIdentities)
    {
        if (draft.Operation != SemanticChangeOperation.Move
            || draft.PackageLocationBefore is null
            || draft.PackageLocationAfter is null)
            return false;
        var before = draft.PackageLocationBefore.Split('\u001f');
        var after = draft.PackageLocationAfter.Split('\u001f');
        if (before.Length != after.Length) return false;
        for (int index = 0; index < before.Length; index++)
        {
            var beforeIdentity = AnchorIdentity(before[index]);
            var afterIdentity = AnchorIdentity(after[index]);
            if (beforeIdentity is null != afterIdentity is null) return false;
            if (beforeIdentity is null)
            {
                if (!string.Equals(before[index], after[index], StringComparison.Ordinal))
                    return false;
                continue;
            }
            if (string.Equals(beforeIdentity, afterIdentity, StringComparison.Ordinal)) continue;
            if (!alignedBlockIdentities.TryGetValue(beforeIdentity, out var aligned)
                || !string.Equals(aligned, afterIdentity, StringComparison.Ordinal))
                return false;
        }
        return true;
    }

    /// <summary>
    /// "scope:unid" identity of an anchor-shaped component, or null for paths/names/uris. The
    /// kind token is deliberately dropped: the IR derives it from presentation (p vs h vs li)
    /// while the package detector always says "p", and identity is scope + Unid.
    /// </summary>
    private static string? AnchorIdentity(string component)
    {
        var parts = component.Split(':');
        if (parts.Length != 3 || parts[2].Length != 32) return null;
        foreach (var letter in parts[0])
            if (!char.IsAsciiLetterLower(letter)) return null;
        foreach (var digit in parts[2])
            if (!char.IsAsciiHexDigitLower(digit)) return null;
        return parts[0].Length == 0 ? null : parts[1] + ":" + parts[2];
    }

    private sealed class Projector
    {
        private readonly IrDocument _left;
        private readonly IrDocument _right;
        private readonly IrDiffSettings _settings;
        private readonly List<SemanticChangeDraft> _changes;

        public Projector(
            IrDocument left,
            IrDocument right,
            IrDiffSettings settings,
            List<SemanticChangeDraft> changes)
        {
            _left = left;
            _right = right;
            _settings = settings;
            _changes = changes;
        }

        /// <summary>
        /// Every left→right "scope:unid" block identity the edit script aligned in place
        /// (equal, modified, format-only, split/merge, and move pairs, including rows and
        /// cells), collected during projection for the package-Move suppression check.
        /// </summary>
        public Dictionary<string, string> AlignedBlockIdentities { get; } =
            new(StringComparer.Ordinal);

        private void RecordAlignment(string? leftAnchor, string? rightAnchor)
        {
            if (leftAnchor is null || rightAnchor is null) return;
            var leftIdentity = AnchorIdentity(leftAnchor);
            var rightIdentity = AnchorIdentity(rightAnchor);
            if (leftIdentity is null || rightIdentity is null) return;
            AlignedBlockIdentities.TryAdd(leftIdentity, rightIdentity);
        }

        public void Project(IrEditScript script)
        {
            var bodyPart = _right.Body.PartUri?.ToString()
                ?? _left.Body.PartUri?.ToString()
                ?? "/word/document.xml";
            ProjectOps(script.Operations, bodyPart, "body", "body");

            if (script.NoteOps is { } notes)
            {
                foreach (var note in notes)
                {
                    var leftStore = note.Kind == IrNoteKind.Footnote ? _left.Footnotes : _left.Endnotes;
                    var rightStore = note.Kind == IrNoteKind.Footnote ? _right.Footnotes : _right.Endnotes;
                    IrScope? leftScope = null;
                    IrScope? rightScope = null;
                    if (note.LeftNoteId is { } leftId)
                        leftStore.Notes.TryGetValue(leftId, out leftScope);

                    // A deleted-only note deliberately reuses its left id as NoteId. Looking that id
                    // up in the right store can accidentally attach an unrelated note when ids collide.
                    // The operation anchors are the authoritative side-presence signal.
                    bool hasRight = note.LeftNoteId is null || note.Ops.Any(HasRightSide);
                    if (hasRight)
                        rightStore.Notes.TryGetValue(note.NoteId, out rightScope);
                    var family = note.Kind == IrNoteKind.Footnote
                        ? SemanticChangeFamily.Footnote
                        : SemanticChangeFamily.Endnote;
                    var part = rightScope?.PartUri?.ToString()
                        ?? leftScope?.PartUri?.ToString()
                        ?? (note.Kind == IrNoteKind.Footnote
                            ? "/word/footnotes.xml"
                            : "/word/endnotes.xml");
                    Add(
                        NoteOperation(leftScope, rightScope), family, part,
                        $"note[{note.NoteId}]", null, null,
                        leftScope?.Name, rightScope?.Name, null,
                        ScopeValue(leftScope), ScopeValue(rightScope));
                    ProjectOps(note.Ops, part, leftScope?.Name, rightScope?.Name);
                }
            }

            if (script.HeaderFooterOps is { } stories)
            {
                foreach (var story in stories)
                {
                    var leftStory = FindStory(_left, story.IsHeader,
                        story.LeftPartUri, story.LeftScopeName);
                    var rightStory = FindStory(_right, story.IsHeader,
                        story.RightPartUri, story.ScopeName);
                    var family = story.IsHeader
                        ? SemanticChangeFamily.Header
                        : SemanticChangeFamily.Footer;
                    var part = story.RightPartUri?.ToString()
                        ?? story.LeftPartUri?.ToString()
                        ?? "/word/document.xml";
                    Add(
                        StoryOperation(story), family, part,
                        $"{(story.IsHeader ? "header" : "footer")}[section={story.SectionIndex},kind={story.Kind}]",
                        null, null, story.LeftScopeName, story.ScopeName, null,
                        StoryValue(leftStory), StoryValue(rightStory));
                    ProjectOps(story.Ops, part, story.LeftScopeName, story.ScopeName);
                }
            }
        }

        public void CompareRegistries()
        {
            CompareStyles();
            CompareNumbering();
            if (_left.ThemeFonts != _right.ThemeFonts)
            {
                var part = _right.ThemeFonts.PartUri?.ToString()
                    ?? _left.ThemeFonts.PartUri?.ToString()
                    ?? "/word/theme/theme1.xml";
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.Style,
                    part, "theme.fonts", null, null, null, null, null,
                    ThemeValue(_left.ThemeFonts), ThemeValue(_right.ThemeFonts));
            }
        }

        public void CompareComments()
        {
            var left = _left.Comments.Comments;
            var right = _right.Comments.Comments;
            var part = _right.Comments.PartUri?.ToString()
                ?? _left.Comments.PartUri?.ToString()
                ?? "/word/comments.xml";

            var unmatchedLeft = new HashSet<int>(Enumerable.Range(0, left.Count));
            var unmatchedRight = new HashSet<int>(Enumerable.Range(0, right.Count));

            // Preserve unchanged comments first. Generated comment Unids can be structural, so an
            // insertion ahead of an existing definition may reuse the old position-derived anchor.
            // Matching the complete semantic value prevents that anchor churn from cascading.
            MatchComments(
                left,
                right,
                unmatchedLeft,
                unmatchedRight,
                comment => ValueKey(CommentValue(comment)),
                part);

            // Editing a comment changes its value. Its target ranges are the next stable identity
            // axis and pair the edited definition back to the same review location.
            MatchComments(
                left,
                right,
                unmatchedLeft,
                unmatchedRight,
                CommentTargetKey,
                part,
                requireNonEmptyKey: true);

            // A comment moved to another range retains its authored definition even though its
            // target key changed.
            MatchComments(
                left,
                right,
                unmatchedLeft,
                unmatchedRight,
                CommentDefinitionKey,
                part);

            // Explicitly persisted Unids remain useful for comments whose definition and target
            // both changed, but they are deliberately weaker than semantic identity.
            MatchComments(
                left,
                right,
                unmatchedLeft,
                unmatchedRight,
                comment => comment.Anchor.ToString(),
                part);

            // Deterministic fallback for malformed/id-less or simultaneously moved-and-edited
            // comments. Pair only the still-unmatched residue in original order.
            var leftResidue = unmatchedLeft.OrderBy(index => index).ToArray();
            var rightResidue = unmatchedRight.OrderBy(index => index).ToArray();
            int pairedResidue = Math.Min(leftResidue.Length, rightResidue.Length);
            for (int index = 0; index < pairedResidue; index++)
            {
                int leftIndex = leftResidue[index];
                int rightIndex = rightResidue[index];
                EmitCommentPair(left[leftIndex], right[rightIndex], part);
                unmatchedLeft.Remove(leftIndex);
                unmatchedRight.Remove(rightIndex);
            }

            foreach (int index in unmatchedLeft.OrderBy(index => index))
            {
                var comment = left[index];
                Add(SemanticChangeOperation.Delete, SemanticChangeFamily.Comment, part,
                    $"comment[{index}]", comment.Anchor.ToString(), null,
                    comment.Anchor.Scope, null, null, CommentValue(comment), SemanticValue.Absent);
            }
            foreach (int index in unmatchedRight.OrderBy(index => index))
            {
                var comment = right[index];
                Add(SemanticChangeOperation.Insert, SemanticChangeFamily.Comment, part,
                    $"comment[{index}]", null, comment.Anchor.ToString(),
                    null, comment.Anchor.Scope, null, SemanticValue.Absent, CommentValue(comment));
            }
        }

        private void MatchComments(
            IReadOnlyList<IrComment> left,
            IReadOnlyList<IrComment> right,
            HashSet<int> unmatchedLeft,
            HashSet<int> unmatchedRight,
            Func<IrComment, string> keySelector,
            string part,
            bool requireNonEmptyKey = false)
        {
            var rightByKey = unmatchedRight
                .GroupBy(index => keySelector(right[index]), StringComparer.Ordinal)
                .ToDictionary(
                    group => group.Key,
                    group => new Queue<int>(group.OrderBy(index => index)),
                    StringComparer.Ordinal);
            foreach (int leftIndex in unmatchedLeft.OrderBy(index => index).ToArray())
            {
                var key = keySelector(left[leftIndex]);
                if (requireNonEmptyKey && string.IsNullOrEmpty(key)) continue;
                if (!rightByKey.TryGetValue(key, out var candidates)) continue;
                while (candidates.Count > 0 && !unmatchedRight.Contains(candidates.Peek()))
                    candidates.Dequeue();
                if (candidates.Count == 0) continue;
                int rightIndex = candidates.Dequeue();
                EmitCommentPair(left[leftIndex], right[rightIndex], part);
                unmatchedLeft.Remove(leftIndex);
                unmatchedRight.Remove(rightIndex);
            }
        }

        private void EmitCommentPair(IrComment left, IrComment right, string part)
        {
            var before = CommentValue(left);
            var after = CommentValue(right);
            if (ValueKey(before).Equals(ValueKey(after), StringComparison.Ordinal)) return;
            Add(SemanticChangeOperation.Modify, SemanticChangeFamily.Comment, part,
                "comment", left.Anchor.ToString(), right.Anchor.ToString(),
                left.Anchor.Scope, right.Anchor.Scope, null, before, after);
        }

        private static string CommentTargetKey(IrComment comment) => string.Join("\u001f",
            comment.Targets.Select(target => string.Join(":",
                target.BlockAnchor.ToString(),
                target.StartChar.ToString(System.Globalization.CultureInfo.InvariantCulture),
                target.EndChar.ToString(System.Globalization.CultureInfo.InvariantCulture))));

        private static string CommentDefinitionKey(IrComment comment) => ValueKey(Obj(
            ("author", SemanticValue.String(comment.Author)),
            ("initials", SemanticValue.String(comment.Initials)),
            ("date", SemanticValue.String(comment.Date)),
            ("blocks", SemanticValue.Array(comment.Blocks.Select(BlockValue)))));

        private void ProjectOps(
            IReadOnlyList<IrEditOp> ops,
            string partUri,
            string? leftScope,
            string? rightScope)
        {
            var moveSources = ops
                .Where(op => op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock
                    && op.IsMoveSource == true && op.MoveGroupId.HasValue)
                .ToDictionary(op => op.MoveGroupId!.Value);

            foreach (var op in ops)
            {
                RecordAlignment(op.LeftAnchor, op.RightAnchor);
                if (op.Kind == IrEditOpKind.MergeBlock && op.SplitMergeAnchors is { } mergedLefts)
                    foreach (var mergedLeft in mergedLefts)
                        RecordAlignment(mergedLeft, op.RightAnchor);
                if (op.Kind == IrEditOpKind.SplitBlock && op.SplitMergeAnchors is { } splitRights)
                    foreach (var splitRight in splitRights)
                        RecordAlignment(op.LeftAnchor, splitRight);

                if (op.Kind == IrEditOpKind.EqualBlock)
                    continue;

                if (op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock)
                {
                    if (op.IsMoveSource == true) continue;
                    moveSources.TryGetValue(op.MoveGroupId ?? -1, out var source);
                    var leftAnchor = source?.LeftAnchor;
                    RecordAlignment(leftAnchor, op.RightAnchor);
                    var leftBlock = Find(_left, leftAnchor);
                    var rightBlock = Find(_right, op.RightAnchor);
                    Add(SemanticChangeOperation.Move, SemanticChangeFamily.BlockStructure, partUri,
                        "block", leftAnchor, op.RightAnchor,
                        Scope(leftAnchor) ?? leftScope, Scope(op.RightAnchor) ?? rightScope,
                        op.MoveGroupId is { } group
                            ? MoveId("block", partUri, leftAnchor, op.RightAnchor, group)
                            : null,
                        BlockValue(leftBlock), BlockValue(rightBlock));
                    if (op.Kind == IrEditOpKind.MoveModifyBlock && leftBlock is not null && rightBlock is not null)
                        CompareBlocks(leftBlock, rightBlock, op, partUri, leftScope, rightScope);
                    continue;
                }

                var leftBlockForOp = Find(_left, op.LeftAnchor);
                var rightBlockForOp = Find(_right, op.RightAnchor);
                switch (op.Kind)
                {
                    case IrEditOpKind.InsertBlock:
                        Add(SemanticChangeOperation.Insert, SemanticChangeFamily.BlockStructure, partUri,
                            "block", null, op.RightAnchor, null,
                            Scope(op.RightAnchor) ?? rightScope, null,
                            SemanticValue.Absent, BlockValue(rightBlockForOp));
                        EmitWholeBlockFeatures(null, rightBlockForOp, partUri, leftScope, rightScope);
                        break;
                    case IrEditOpKind.DeleteBlock:
                        Add(SemanticChangeOperation.Delete, SemanticChangeFamily.BlockStructure, partUri,
                            "block", op.LeftAnchor, null,
                            Scope(op.LeftAnchor) ?? leftScope, null, null,
                            BlockValue(leftBlockForOp), SemanticValue.Absent);
                        EmitWholeBlockFeatures(leftBlockForOp, null, partUri, leftScope, rightScope);
                        break;
                    case IrEditOpKind.ModifyBlock:
                    case IrEditOpKind.FormatOnlyBlock:
                        if (leftBlockForOp is not null && rightBlockForOp is not null)
                            CompareBlocks(leftBlockForOp, rightBlockForOp, op, partUri, leftScope, rightScope);
                        break;
                    case IrEditOpKind.SplitBlock:
                    case IrEditOpKind.MergeBlock:
                        Add(SemanticChangeOperation.Modify, SemanticChangeFamily.BlockStructure, partUri,
                            op.Kind == IrEditOpKind.SplitBlock ? "paragraph.split" : "paragraph.merge",
                            op.LeftAnchor, op.RightAnchor,
                            Scope(op.LeftAnchor) ?? leftScope, Scope(op.RightAnchor) ?? rightScope, null,
                            AnchorArray(op.Kind == IrEditOpKind.SplitBlock
                                ? new[] { op.LeftAnchor }
                                : op.SplitMergeAnchors),
                            AnchorArray(op.Kind == IrEditOpKind.SplitBlock
                                ? op.SplitMergeAnchors
                                : new[] { op.RightAnchor }));
                        ProjectSplitMergeTokenChanges(op, partUri, leftScope, rightScope);
                        break;
                    default:
                        Add(SemanticChangeOperation.Modify, SemanticChangeFamily.BlockStructure, partUri,
                            "block", op.LeftAnchor, op.RightAnchor,
                            Scope(op.LeftAnchor) ?? leftScope, Scope(op.RightAnchor) ?? rightScope, null,
                            BlockValue(leftBlockForOp), BlockValue(rightBlockForOp));
                        break;
                }
            }
        }

        private void CompareBlocks(
            IrBlock left,
            IrBlock right,
            IrEditOp op,
            string fallbackPart,
            string? leftScope,
            string? rightScope)
        {
            var part = right.Source.PartUri?.ToString()
                ?? left.Source.PartUri?.ToString()
                ?? fallbackPart;
            if (left.GetType() != right.GetType())
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.BlockStructure, part,
                    "block.kind", left.Anchor.ToString(), right.Anchor.ToString(),
                    left.Anchor.Scope, right.Anchor.Scope, null,
                    SemanticValue.String(BlockKind(left)), SemanticValue.String(BlockKind(right)));
                EmitWholeBlockFeatures(left, null, part, leftScope, rightScope);
                EmitWholeBlockFeatures(null, right, part, leftScope, rightScope);
                return;
            }

            switch (left, right)
            {
                case (IrParagraph lp, IrParagraph rp):
                    CompareParagraph(lp, rp, op, part);
                    break;
                case (IrTable lt, IrTable rt):
                    CompareTable(lt, rt, op.TableDiff, part);
                    break;
                case (IrSectionBreak ls, IrSectionBreak rs):
                    CompareSection(ls.Format, rs.Format, part,
                        ls.Anchor.ToString(), rs.Anchor.ToString(), ls.Anchor.Scope, rs.Anchor.Scope);
                    break;
                case (IrSdtBlock lsdt, IrSdtBlock rsdt):
                    if (lsdt.EnvelopeDigest != rsdt.EnvelopeDigest)
                    {
                        Add(SemanticChangeOperation.Modify, SemanticChangeFamily.ContentControl, part,
                            "content_control.envelope", lsdt.Anchor.ToString(), rsdt.Anchor.ToString(),
                            lsdt.Anchor.Scope, rsdt.Anchor.Scope, null,
                            Digest(lsdt.EnvelopeDigest, "docxodus-ir-sdt-envelope-v1"),
                            Digest(rsdt.EnvelopeDigest, "docxodus-ir-sdt-envelope-v1"));
                    }
                    var sdtAlignment = IrBlockAligner.AlignBlocks(lsdt.Blocks, rsdt.Blocks, _settings);
                    ProjectOps(
                        IrEditScriptBuilder.ProjectAlignment(lsdt.Blocks, sdtAlignment, _settings),
                        part,
                        lsdt.Anchor.Scope,
                        rsdt.Anchor.Scope);
                    break;
                case (IrOpaqueBlock lo, IrOpaqueBlock ro):
                    if (lo.ContentHash != ro.ContentHash || lo.ElementName != ro.ElementName)
                    {
                        Add(SemanticChangeOperation.Modify, SemanticChangeFamily.OpaquePackagePart, part,
                            "block.opaque", lo.Anchor.ToString(), ro.Anchor.ToString(),
                            lo.Anchor.Scope, ro.Anchor.Scope, null, BlockValue(lo), BlockValue(ro));
                    }
                    break;
            }
        }

        private void CompareParagraph(IrParagraph left, IrParagraph right, IrEditOp op, string part)
        {
            var la = left.Anchor.ToString();
            var ra = right.Anchor.ToString();
            if (left.Format != right.Format)
            {
                if (left.Format.StyleId != right.Format.StyleId)
                {
                    Add(SemanticChangeOperation.Modify, SemanticChangeFamily.Style, part,
                        "paragraph.style", la, ra, left.Anchor.Scope, right.Anchor.Scope, null,
                        SemanticValue.String(left.Format.StyleId), SemanticValue.String(right.Format.StyleId));
                }
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.ParagraphFormatting, part,
                    "paragraph.format", la, ra, left.Anchor.Scope, right.Anchor.Scope, null,
                    ParagraphFormatValue(left.Format), ParagraphFormatValue(right.Format));
            }

            if (!Equals(left.List, right.List)
                || left.ResolvedListMarker != right.ResolvedListMarker
                || left.IsListItemForLayout != right.IsListItemForLayout)
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.List, part,
                    "paragraph.list", la, ra, left.Anchor.Scope, right.Anchor.Scope, null,
                    ListValue(left), ListValue(right));
            }

            if (left.InlineSectionFormat != right.InlineSectionFormat)
                CompareSection(left.InlineSectionFormat, right.InlineSectionFormat, part,
                    la, ra, left.Anchor.Scope, right.Anchor.Scope);

            if (left.InlineEnvelopeDigest != right.InlineEnvelopeDigest)
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.ContentControl, part,
                    "paragraph.inline_content_control", la, ra,
                    left.Anchor.Scope, right.Anchor.Scope, null,
                    Digest(left.InlineEnvelopeDigest, "docxodus-ir-inline-envelope-v1"),
                    Digest(right.InlineEnvelopeDigest, "docxodus-ir-inline-envelope-v1"));
            }
            if (left.FieldEnvelopeDigest != right.FieldEnvelopeDigest)
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.Field, part,
                    "paragraph.field_envelope", la, ra,
                    left.Anchor.Scope, right.Anchor.Scope, null,
                    Digest(left.FieldEnvelopeDigest, "docxodus-ir-field-envelope-v1"),
                    Digest(right.FieldEnvelopeDigest, "docxodus-ir-field-envelope-v1"));
            }

            EmitTokenChanges(left, right, op.TokenDiff, part);
            CompareInlineFeatures(left, right, part);
            if (op.TextboxDiffs is { } textboxDiffs)
            {
                foreach (var textboxDiff in textboxDiffs)
                    ProjectOps(textboxDiff.Ops, part, left.Anchor.Scope, right.Anchor.Scope);
            }
        }

        private void EmitTokenChanges(
            IrParagraph left,
            IrParagraph right,
            IrTokenDiff? tokenDiff,
            string part)
        {
            bool visibleTextIsEqual = string.Equals(
                PlainText(left), PlainText(right), StringComparison.Ordinal);
            if (tokenDiff is null)
            {
                if (left.ContentHash != right.ContentHash
                    && !visibleTextIsEqual)
                {
                    Add(SemanticChangeOperation.Modify, SemanticChangeFamily.Text, part,
                        "paragraph.text", left.Anchor.ToString(), right.Anchor.ToString(),
                        left.Anchor.Scope, right.Anchor.Scope, null,
                        SemanticValue.String(PlainText(left)), SemanticValue.String(PlainText(right)));
                }
                CompareAtomicTokenFormats(left, right, part);
                CompareRunFormats(left, right, part);
                return;
            }

            var leftTokens = IrDiffTokenizer.Tokenize(left, _settings);
            var rightTokens = IrDiffTokenizer.Tokenize(right, _settings);
            bool needsCharacterFormatComparison = visibleTextIsEqual
                && tokenDiff.Ops.Any(tokenOp =>
                    tokenOp.Kind is IrTokenOpKind.Insert or IrTokenOpKind.Delete
                    && (HasOrdinaryToken(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd)
                        || HasOrdinaryToken(rightTokens, tokenOp.RightStart, tokenOp.RightEnd)));
            EmitTokenDiff(
                leftTokens,
                rightTokens,
                tokenDiff,
                part,
                "paragraph.tokens",
                left.Anchor.ToString(),
                right.Anchor.ToString(),
                left.Anchor.Scope,
                right.Anchor.Scope,
                emitFormatChanges: !needsCharacterFormatComparison);
            if (needsCharacterFormatComparison)
                CompareRunFormats(left, right, part);
        }

        private void ProjectSplitMergeTokenChanges(
            IrEditOp op,
            string part,
            string? leftScope,
            string? rightScope)
        {
            if (op.SplitMergeAnchors is null || op.SegmentDiffs is null
                || op.SplitMergeAnchors.Count != op.SegmentDiffs.Count)
                return;

            if (op.Kind == IrEditOpKind.SplitBlock)
            {
                if (Find(_left, op.LeftAnchor) is not IrParagraph left) return;
                var leftTokens = IrDiffTokenizer.Tokenize(left, _settings);
                int leftOffset = 0;
                for (int index = 0; index < op.SegmentDiffs.Count; index++)
                {
                    if (Find(_right, op.SplitMergeAnchors[index]) is not IrParagraph right) return;
                    var diff = op.SegmentDiffs[index];
                    int leftLength = diff.Ops.Sum(item => item.LeftLength);
                    if (leftLength < 0 || leftOffset > leftTokens.Count - leftLength) return;
                    var leftSlice = leftTokens.Skip(leftOffset).Take(leftLength).ToArray();
                    var rightTokens = IrDiffTokenizer.Tokenize(right, _settings);
                    EmitTokenDiff(leftSlice, rightTokens, diff, part,
                        $"paragraph.split.segment[{index}].tokens",
                        left.Anchor.ToString(), right.Anchor.ToString(),
                        left.Anchor.Scope ?? leftScope, right.Anchor.Scope ?? rightScope);
                    leftOffset += leftLength;
                }
                return;
            }

            if (op.Kind != IrEditOpKind.MergeBlock
                || Find(_right, op.RightAnchor) is not IrParagraph merged)
                return;
            var mergedTokens = IrDiffTokenizer.Tokenize(merged, _settings);
            int rightOffset = 0;
            for (int index = 0; index < op.SegmentDiffs.Count; index++)
            {
                if (Find(_left, op.SplitMergeAnchors[index]) is not IrParagraph left) return;
                var diff = op.SegmentDiffs[index];
                int rightLength = diff.Ops.Sum(item => item.RightLength);
                if (rightLength < 0 || rightOffset > mergedTokens.Count - rightLength) return;
                var leftTokens = IrDiffTokenizer.Tokenize(left, _settings);
                var rightSlice = mergedTokens.Skip(rightOffset).Take(rightLength).ToArray();
                EmitTokenDiff(leftTokens, rightSlice, diff, part,
                    $"paragraph.merge.segment[{index}].tokens",
                    left.Anchor.ToString(), merged.Anchor.ToString(),
                    left.Anchor.Scope ?? leftScope, merged.Anchor.Scope ?? rightScope);
                rightOffset += rightLength;
            }
        }

        private void EmitTokenDiff(
            IReadOnlyList<IrDiffToken> leftTokens,
            IReadOnlyList<IrDiffToken> rightTokens,
            IrTokenDiff tokenDiff,
            string part,
            string path,
            string? leftAnchor,
            string? rightAnchor,
            string? leftScope,
            string? rightScope,
            bool emitFormatChanges = true)
        {
            bool textIsEqual = string.Equals(
                string.Concat(leftTokens.Select(token => token.Text)),
                string.Concat(rightTokens.Select(token => token.Text)),
                StringComparison.Ordinal);
            foreach (var tokenOp in tokenDiff.Ops)
            {
                if (tokenOp.Kind == IrTokenOpKind.Equal) continue;
                if (tokenOp.Kind == IrTokenOpKind.FormatChanged)
                {
                    if (emitFormatChanges)
                    {
                        Add(SemanticChangeOperation.Modify, SemanticChangeFamily.RunFormatting, part,
                            $"{path}[{tokenOp.LeftStart}:{tokenOp.LeftEnd}]", leftAnchor,
                            rightAnchor, leftScope, rightScope, null,
                            TokenFormatValue(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd),
                            TokenFormatValue(rightTokens, tokenOp.RightStart, tokenOp.RightEnd));
                    }
                    continue;
                }

                // Relationship or structural-carrier edits can make the token differ report a
                // delete+insert even though the user-visible text is unchanged. The corresponding
                // hyperlink/field/content-control family records the real semantic change.
                // Zero-width atomic tokens are different: break kinds, note references, images,
                // and opaque carriers have no PlainText but are semantic content in their own right.
                if (textIsEqual
                    && !HasAtomicToken(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd)
                    && !HasAtomicToken(rightTokens, tokenOp.RightStart, tokenOp.RightEnd))
                    continue;

                var operation = tokenOp.Kind == IrTokenOpKind.Insert
                    ? SemanticChangeOperation.Insert
                    : SemanticChangeOperation.Delete;
                Add(operation, SemanticChangeFamily.Text, part,
                    $"{path}[{tokenOp.LeftStart}:{tokenOp.LeftEnd}|{tokenOp.RightStart}:{tokenOp.RightEnd}]",
                    leftAnchor, rightAnchor, leftScope, rightScope, null,
                    tokenOp.Kind == IrTokenOpKind.Insert
                        ? SemanticValue.Absent
                        : TokenTextValue(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd),
                    tokenOp.Kind == IrTokenOpKind.Delete
                        ? SemanticValue.Absent
                        : TokenTextValue(rightTokens, tokenOp.RightStart, tokenOp.RightEnd));
            }
        }

        private static bool HasAtomicToken(
            IReadOnlyList<IrDiffToken> tokens,
            int start,
            int end) => tokens.Skip(start).Take(end - start)
                .Any(token => token.Kind is not (IrDiffTokenKind.Word or IrDiffTokenKind.Separator));

        private static bool HasOrdinaryToken(
            IReadOnlyList<IrDiffToken> tokens,
            int start,
            int end) => tokens.Skip(start).Take(end - start)
                .Any(token => token.Kind is IrDiffTokenKind.Word or IrDiffTokenKind.Separator);

        private void CompareAtomicTokenFormats(IrParagraph left, IrParagraph right, string part)
        {
            var leftTokens = IrDiffTokenizer.Tokenize(left, _settings);
            var rightTokens = IrDiffTokenizer.Tokenize(right, _settings);
            var diff = IrTokenDiffer.Diff(leftTokens, rightTokens, _settings);
            foreach (var op in diff.Ops.Where(item => item.Kind == IrTokenOpKind.FormatChanged))
            {
                for (int offset = 0; offset < op.LeftLength; offset++)
                {
                    int leftIndex = op.LeftStart + offset;
                    int rightIndex = op.RightStart + offset;
                    var leftToken = leftTokens[leftIndex];
                    var rightToken = rightTokens[rightIndex];
                    if (leftToken.Kind is IrDiffTokenKind.Word or IrDiffTokenKind.Separator
                        || rightToken.Kind is IrDiffTokenKind.Word or IrDiffTokenKind.Separator)
                        continue;

                    Add(SemanticChangeOperation.Modify, SemanticChangeFamily.RunFormatting, part,
                        $"paragraph.atomic_tokens[{leftIndex}:{rightIndex}].format",
                        left.Anchor.ToString(), right.Anchor.ToString(),
                        left.Anchor.Scope, right.Anchor.Scope, null,
                        RunFormatValue(leftToken.Format), RunFormatValue(rightToken.Format));
                }
            }
        }

        private void CompareRunFormats(IrParagraph left, IrParagraph right, string part)
        {
            var leftRuns = FlattenInlines(left.Inlines).OfType<IrTextRun>().ToArray();
            var rightRuns = FlattenInlines(right.Inlines).OfType<IrTextRun>().ToArray();
            if (!string.Equals(
                    string.Concat(leftRuns.Select(run => run.Text)),
                    string.Concat(rightRuns.Select(run => run.Text)),
                    StringComparison.Ordinal))
                return;

            var leftSpans = RunFormatSpans(leftRuns);
            var rightSpans = RunFormatSpans(rightRuns);
            int leftIndex = 0;
            int rightIndex = 0;
            var deltas = new List<RunFormatDelta>();
            while (leftIndex < leftSpans.Count && rightIndex < rightSpans.Count)
            {
                var leftSpan = leftSpans[leftIndex];
                var rightSpan = rightSpans[rightIndex];
                int start = Math.Max(leftSpan.Start, rightSpan.Start);
                int end = Math.Min(leftSpan.End, rightSpan.End);
                if (start < end && !IrModeledFormat.RunFormatEqual(
                        leftSpan.Format, rightSpan.Format, _settings.FormatComparison))
                {
                    var delta = new RunFormatDelta(start, end, leftSpan.Format, rightSpan.Format);
                    if (deltas.Count > 0
                        && deltas[^1].End == delta.Start
                        && deltas[^1].Before == delta.Before
                        && deltas[^1].After == delta.After)
                    {
                        deltas[^1] = deltas[^1] with { End = delta.End };
                    }
                    else
                    {
                        deltas.Add(delta);
                    }
                }

                if (leftSpan.End == end) leftIndex++;
                if (rightSpan.End == end) rightIndex++;
            }

            foreach (var delta in deltas)
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.RunFormatting, part,
                    $"paragraph.characters[{delta.Start}:{delta.End}].format",
                    left.Anchor.ToString(), right.Anchor.ToString(),
                    left.Anchor.Scope, right.Anchor.Scope, null,
                    RunFormatValue(delta.Before), RunFormatValue(delta.After));
            }
        }

        private static IReadOnlyList<RunFormatSpan> RunFormatSpans(
            IReadOnlyList<IrTextRun> runs)
        {
            var spans = new List<RunFormatSpan>();
            int offset = 0;
            foreach (var run in runs)
            {
                int end = offset + run.Text.Length;
                if (offset < end)
                {
                    if (spans.Count > 0
                        && spans[^1].End == offset
                        && spans[^1].Format == run.Format)
                        spans[^1] = spans[^1] with { End = end };
                    else
                        spans.Add(new RunFormatSpan(offset, end, run.Format));
                }
                offset = end;
            }
            return spans;
        }

        private void CompareInlineFeatures(IrParagraph left, IrParagraph right, string part)
        {
            var leftFeatures = InlineFeatures(left.Inlines).ToArray();
            var rightFeatures = InlineFeatures(right.Inlines).ToArray();
            var leftKeys = leftFeatures.Select(InlineFeatureKey).ToArray();
            var rightKeys = rightFeatures.Select(InlineFeatureKey).ToArray();
            var unmatchedLeft = new HashSet<int>(Enumerable.Range(0, leftFeatures.Length));
            var unmatchedRight = new HashSet<int>(Enumerable.Range(0, rightFeatures.Length));

            // Preserve a longest ordered exact-match spine. A greedy first match can badly
            // misattribute a rotation (A,B,C -> B,C,A) as movement of B and C instead of A.
            // The bounded LCS gives the minimal residue for ordinary paragraphs; exceptionally
            // feature-dense inputs use the deterministic linear-memory greedy fallback below.
            foreach (var (leftIndex, rightIndex) in ExactFeatureSpine(leftKeys, rightKeys))
            {
                unmatchedLeft.Remove(leftIndex);
                unmatchedRight.Remove(rightIndex);
            }

            foreach (var family in leftFeatures.Select(feature => feature.Family)
                .Concat(rightFeatures.Select(feature => feature.Family))
                .Distinct()
                .OrderBy(value => value))
            {
                var leftResidue = unmatchedLeft
                    .Where(index => leftFeatures[index].Family == family)
                    .OrderBy(index => index)
                    .ToArray();
                var rightResidue = unmatchedRight
                    .Where(index => rightFeatures[index].Family == family)
                    .OrderBy(index => index)
                    .ToArray();
                int paired = Math.Min(leftResidue.Length, rightResidue.Length);
                for (int index = 0; index < paired; index++)
                {
                    int leftIndex = leftResidue[index];
                    int rightIndex = rightResidue[index];
                    var l = leftFeatures[leftIndex];
                    var r = rightFeatures[rightIndex];
                    if (leftKeys[leftIndex] == rightKeys[rightIndex])
                    {
                        // Equal values left outside the ordered spine are a genuine reorder.
                        EmitInlineFeature(SemanticChangeOperation.Delete, l, leftIndex,
                            left, right, part);
                        EmitInlineFeature(SemanticChangeOperation.Insert, r, rightIndex,
                            left, right, part);
                    }
                    else
                    {
                        Add(SemanticChangeOperation.Modify, family, part,
                            $"paragraph.{l.Path}[{leftIndex}]", left.Anchor.ToString(),
                            right.Anchor.ToString(), left.Anchor.Scope, right.Anchor.Scope,
                            null, l.Value, r.Value);
                    }
                    unmatchedLeft.Remove(leftIndex);
                    unmatchedRight.Remove(rightIndex);
                }
            }

            foreach (int index in unmatchedLeft.OrderBy(index => index))
                EmitInlineFeature(SemanticChangeOperation.Delete, leftFeatures[index], index,
                    left, right, part);
            foreach (int index in unmatchedRight.OrderBy(index => index))
                EmitInlineFeature(SemanticChangeOperation.Insert, rightFeatures[index], index,
                    left, right, part);
        }

        private static IReadOnlyList<(int Left, int Right)> ExactFeatureSpine(
            IReadOnlyList<string> left,
            IReadOnlyList<string> right)
        {
            const long cellLimit = 1_000_000;
            if ((long)left.Count * right.Count > cellLimit)
                return GreedyFeatureSpine(left, right);

            var lengths = new int[left.Count + 1, right.Count + 1];
            for (int leftIndex = left.Count - 1; leftIndex >= 0; leftIndex--)
            {
                for (int rightIndex = right.Count - 1; rightIndex >= 0; rightIndex--)
                {
                    lengths[leftIndex, rightIndex] = string.Equals(
                        left[leftIndex], right[rightIndex], StringComparison.Ordinal)
                        ? lengths[leftIndex + 1, rightIndex + 1] + 1
                        : Math.Max(
                            lengths[leftIndex + 1, rightIndex],
                            lengths[leftIndex, rightIndex + 1]);
                }
            }

            var result = new List<(int Left, int Right)>();
            int li = 0;
            int ri = 0;
            while (li < left.Count && ri < right.Count)
            {
                if (string.Equals(left[li], right[ri], StringComparison.Ordinal)
                    && lengths[li, ri] == lengths[li + 1, ri + 1] + 1)
                {
                    result.Add((li++, ri++));
                }
                else if (lengths[li + 1, ri] >= lengths[li, ri + 1])
                {
                    li++;
                }
                else
                {
                    ri++;
                }
            }
            return result;
        }

        private static IReadOnlyList<(int Left, int Right)> GreedyFeatureSpine(
            IReadOnlyList<string> left,
            IReadOnlyList<string> right)
        {
            var rightPositions = Enumerable.Range(0, right.Count)
                .GroupBy(index => right[index], StringComparer.Ordinal)
                .ToDictionary(
                    group => group.Key,
                    group => new SortedSet<int>(group),
                    StringComparer.Ordinal);
            var result = new List<(int Left, int Right)>();
            int nextRight = 0;
            for (int leftIndex = 0; leftIndex < left.Count; leftIndex++)
            {
                if (!rightPositions.TryGetValue(left[leftIndex], out var candidates)) continue;
                var tail = candidates.GetViewBetween(nextRight, int.MaxValue);
                if (tail.Count == 0) continue;
                int rightIndex = tail.Min;
                candidates.Remove(rightIndex);
                result.Add((leftIndex, rightIndex));
                nextRight = rightIndex + 1;
            }
            return result;
        }

        private void EmitInlineFeature(
            SemanticChangeOperation operation,
            InlineFeature feature,
            int index,
            IrParagraph left,
            IrParagraph right,
            string part) => Add(
                operation,
                feature.Family,
                part,
                $"paragraph.{feature.Path}[{index}]",
                operation == SemanticChangeOperation.Insert ? null : left.Anchor.ToString(),
                operation == SemanticChangeOperation.Delete ? null : right.Anchor.ToString(),
                operation == SemanticChangeOperation.Insert ? null : left.Anchor.Scope,
                operation == SemanticChangeOperation.Delete ? null : right.Anchor.Scope,
                null,
                operation == SemanticChangeOperation.Insert ? SemanticValue.Absent : feature.Value,
                operation == SemanticChangeOperation.Delete ? SemanticValue.Absent : feature.Value);

        private void CompareTable(IrTable left, IrTable right, IrTableDiff? diff, string part)
        {
            var la = left.Anchor.ToString();
            var ra = right.Anchor.ToString();
            if (left.ContentHash != right.ContentHash || left.FormatFingerprint != right.FormatFingerprint)
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.Table, part,
                    "table", la, ra, left.Anchor.Scope, right.Anchor.Scope, null,
                    BlockValue(left), BlockValue(right));
            }
            if (left.TblPrDigest != right.TblPrDigest)
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.TableStyle, part,
                    "table.properties", la, ra, left.Anchor.Scope, right.Anchor.Scope, null,
                    Digest(left.TblPrDigest, "docxodus-ir-table-properties-v1"),
                    Digest(right.TblPrDigest, "docxodus-ir-table-properties-v1"));
                CompareTableTypedProperties(left, right, part);
            }
            if (left.TblGridDigest != right.TblGridDigest)
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.TableWidth, part,
                    "table.grid", la, ra, left.Anchor.Scope, right.Anchor.Scope, null,
                    TableGridValue(left), TableGridValue(right));
            }
            if (diff is null)
            {
                CompareRowsPositionally(left, right, part);
                return;
            }

            var moveSources = diff.RowOps
                .Where(row => row.Kind == IrRowOpKind.MovedRow
                    && row.IsMoveSource == true && row.MoveGroupId.HasValue)
                .ToDictionary(row => row.MoveGroupId!.Value);
            foreach (var rowOp in diff.RowOps)
            {
                var leftRow = FindRow(left, rowOp.LeftRowAnchor);
                var rightRow = FindRow(right, rowOp.RightRowAnchor);
                switch (rowOp.Kind)
                {
                    case IrRowOpKind.EqualRow:
                        CompareRow(leftRow, rightRow, null, part);
                        break;
                    case IrRowOpKind.InsertRow:
                        EmitWholeRowFeatures(null, rightRow, part);
                        break;
                    case IrRowOpKind.DeleteRow:
                        EmitWholeRowFeatures(leftRow, null, part);
                        break;
                    case IrRowOpKind.MovedRow:
                        if (rowOp.IsMoveSource == true) break;
                        moveSources.TryGetValue(rowOp.MoveGroupId ?? -1, out var source);
                        leftRow = FindRow(left, source?.LeftRowAnchor);
                        Add(SemanticChangeOperation.Move, SemanticChangeFamily.TableRow, part,
                            "table.row", source?.LeftRowAnchor, rowOp.RightRowAnchor,
                            Scope(source?.LeftRowAnchor), Scope(rowOp.RightRowAnchor),
                            rowOp.MoveGroupId is { } group
                                ? MoveId("row", part, source?.LeftRowAnchor, rowOp.RightRowAnchor, group)
                                : null,
                            RowValue(leftRow), RowValue(rightRow));
                        CompareRow(leftRow, rightRow, null, part);
                        break;
                    case IrRowOpKind.ModifyRow:
                        CompareRow(leftRow, rightRow, rowOp.CellOps, part);
                        break;
                }
            }
        }

        private void CompareRowsPositionally(IrTable left, IrTable right, string part)
        {
            int common = Math.Min(left.Rows.Count, right.Rows.Count);
            for (int index = 0; index < common; index++)
                CompareRow(left.Rows[index], right.Rows[index], null, part);
            for (int index = common; index < left.Rows.Count; index++)
                EmitWholeRowFeatures(left.Rows[index], null, part);
            for (int index = common; index < right.Rows.Count; index++)
                EmitWholeRowFeatures(null, right.Rows[index], part);
        }

        private void CompareRow(
            IrRow? left,
            IrRow? right,
            IReadOnlyList<IrCellOp>? cellOps,
            string part)
        {
            if (left is null || right is null) return;
            var la = left.Anchor.ToString();
            var ra = right.Anchor.ToString();
            RecordAlignment(la, ra);
            if (left.GridBefore != right.GridBefore || left.GridAfter != right.GridAfter)
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.TableSpan, part,
                    "table.row.grid_omissions", la, ra, left.Anchor.Scope, right.Anchor.Scope, null,
                    RowGridValue(left), RowGridValue(right));
            }
            if (left.TrPrDigest != right.TrPrDigest || left.FromTableSdt != right.FromTableSdt)
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.TableRow, part,
                    "table.row.properties", la, ra, left.Anchor.Scope, right.Anchor.Scope, null,
                    RowValue(left), RowValue(right));
            }
            IReadOnlyList<IrCellOp> effectiveCellOps = cellOps ?? PositionalCellOps(left, right);
            foreach (var cellOp in effectiveCellOps)
            {
                RecordAlignment(cellOp.LeftCellAnchor, cellOp.RightCellAnchor);
                var lc = left.Cells.FirstOrDefault(cell => cell.Anchor.ToString() == cellOp.LeftCellAnchor);
                var rc = right.Cells.FirstOrDefault(cell => cell.Anchor.ToString() == cellOp.RightCellAnchor);
                if (lc is null)
                {
                    Add(SemanticChangeOperation.Insert, SemanticChangeFamily.TableCell, part,
                        "table.cell", null, cellOp.RightCellAnchor, null, Scope(cellOp.RightCellAnchor), null,
                        SemanticValue.Absent, CellValue(rc));
                    if (rc is not null)
                        foreach (var block in rc.Blocks)
                            EmitWholeBlockFeatures(null, block, part, null, rc.Anchor.Scope);
                    continue;
                }
                if (rc is null)
                {
                    Add(SemanticChangeOperation.Delete, SemanticChangeFamily.TableCell, part,
                        "table.cell", cellOp.LeftCellAnchor, null, Scope(cellOp.LeftCellAnchor), null, null,
                        CellValue(lc), SemanticValue.Absent);
                    foreach (var block in lc.Blocks)
                        EmitWholeBlockFeatures(block, null, part, lc.Anchor.Scope, null);
                    continue;
                }
                if (lc.GridSpan != rc.GridSpan || lc.VMerge != rc.VMerge)
                {
                    Add(SemanticChangeOperation.Modify, SemanticChangeFamily.TableSpan, part,
                        "table.cell.span", lc.Anchor.ToString(), rc.Anchor.ToString(),
                        lc.Anchor.Scope, rc.Anchor.Scope, null, CellSpanValue(lc), CellSpanValue(rc));
                }
                if (lc.ShellDigest != rc.ShellDigest || lc.FromRowSdt != rc.FromRowSdt)
                {
                    Add(SemanticChangeOperation.Modify, SemanticChangeFamily.TableCell, part,
                        "table.cell.properties", lc.Anchor.ToString(), rc.Anchor.ToString(),
                        lc.Anchor.Scope, rc.Anchor.Scope, null, CellValue(lc), CellValue(rc));
                    CompareCellWidth(lc, rc, part);
                }
                var blockOps = cellOp.BlockOps;
                if (blockOps is null && lc.ContentHash != rc.ContentHash)
                {
                    var alignment = IrBlockAligner.AlignBlocks(lc.Blocks, rc.Blocks, _settings);
                    blockOps = IrNodeList.From(
                        IrEditScriptBuilder.ProjectAlignment(lc.Blocks, alignment, _settings));
                }
                if (blockOps is not null)
                    ProjectOps(blockOps, part, lc.Anchor.Scope, rc.Anchor.Scope);
            }
        }

        private static IReadOnlyList<IrCellOp> PositionalCellOps(IrRow left, IrRow right)
        {
            int common = Math.Min(left.Cells.Count, right.Cells.Count);
            var result = new List<IrCellOp>(Math.Max(left.Cells.Count, right.Cells.Count));
            for (int index = 0; index < common; index++)
                result.Add(new IrCellOp(
                    left.Cells[index].Anchor.ToString(),
                    right.Cells[index].Anchor.ToString(),
                    null));
            for (int index = common; index < left.Cells.Count; index++)
                result.Add(new IrCellOp(left.Cells[index].Anchor.ToString(), null, null));
            for (int index = common; index < right.Cells.Count; index++)
                result.Add(new IrCellOp(null, right.Cells[index].Anchor.ToString(), null));
            return result;
        }

        private void CompareSection(
            IrSectionFormat? left,
            IrSectionFormat? right,
            string part,
            string? leftAnchor,
            string? rightAnchor,
            string? leftScope,
            string? rightScope)
        {
            if (Equals(left, right)) return;
            var beforeSection = SectionValue(left);
            var afterSection = SectionValue(right);
            var operation = left is null ? SemanticChangeOperation.Insert
                : right is null ? SemanticChangeOperation.Delete
                : SemanticChangeOperation.Modify;
            if (ValueKey(beforeSection) != ValueKey(afterSection))
            {
                Add(operation, SemanticChangeFamily.Section, part,
                    "section", leftAnchor, rightAnchor, leftScope, rightScope, null,
                    beforeSection, afterSection);
            }
            var beforePage = PageSetupValue(left);
            var afterPage = PageSetupValue(right);
            if (ValueKey(beforePage) != ValueKey(afterPage))
            {
                Add(operation, SemanticChangeFamily.PageSetup, part,
                    "section.page_setup", leftAnchor, rightAnchor, leftScope, rightScope, null,
                    beforePage, afterPage);
            }
        }

        private void CompareStyles()
        {
            var part = _right.Styles.PartUri?.ToString()
                ?? _left.Styles.PartUri?.ToString()
                ?? "/word/styles.xml";
            var ids = _left.Styles.Styles.Keys
                .Concat(_right.Styles.Styles.Keys)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(id => id, StringComparer.Ordinal);
            foreach (var id in ids)
            {
                _left.Styles.Styles.TryGetValue(id, out var left);
                _right.Styles.Styles.TryGetValue(id, out var right);
                var before = StyleValue(left);
                var after = StyleValue(right);
                if (ValueKey(before) == ValueKey(after)) continue;
                Add(Operation(left, right), SemanticChangeFamily.Style, part,
                    $"style[{id}]", null, null, null, null, null, before, after);
            }
        }

        private void CompareNumbering()
        {
            var part = _right.Numbering.PartUri?.ToString()
                ?? _left.Numbering.PartUri?.ToString()
                ?? "/word/numbering.xml";
            var numIds = _left.Numbering.Nums.Keys.Concat(_right.Numbering.Nums.Keys)
                .Distinct().OrderBy(id => id);
            foreach (var id in numIds)
            {
                _left.Numbering.Nums.TryGetValue(id, out var left);
                _right.Numbering.Nums.TryGetValue(id, out var right);
                var before = NumValue(left);
                var after = NumValue(right);
                if (ValueKey(before) == ValueKey(after)) continue;
                Add(Operation(left, right), SemanticChangeFamily.Numbering, part,
                    $"numbering.instance[{id}]", null, null, null, null, null, before, after);
            }
            var abstractIds = _left.Numbering.AbstractNums.Keys.Concat(_right.Numbering.AbstractNums.Keys)
                .Distinct().OrderBy(id => id);
            foreach (var id in abstractIds)
            {
                _left.Numbering.AbstractNums.TryGetValue(id, out var left);
                _right.Numbering.AbstractNums.TryGetValue(id, out var right);
                var before = AbstractNumValue(left);
                var after = AbstractNumValue(right);
                if (ValueKey(before) == ValueKey(after)) continue;
                Add(Operation(left, right), SemanticChangeFamily.Numbering, part,
                    $"numbering.abstract[{id}]", null, null, null, null, null, before, after);
            }
        }

        private void EmitWholeBlockFeatures(
            IrBlock? left,
            IrBlock? right,
            string part,
            string? leftScope,
            string? rightScope)
        {
            if (left is IrParagraph || right is IrParagraph)
            {
                var lp = left as IrParagraph;
                var rp = right as IrParagraph;
                var operation = Operation(lp, rp);
                Add(operation, SemanticChangeFamily.Text, part, "paragraph.text",
                    lp?.Anchor.ToString(), rp?.Anchor.ToString(),
                    lp?.Anchor.Scope ?? leftScope, rp?.Anchor.Scope ?? rightScope, null,
                    lp is null ? SemanticValue.Absent : SemanticValue.String(PlainText(lp)),
                    rp is null ? SemanticValue.Absent : SemanticValue.String(PlainText(rp)));
                if (lp?.Format is not null || rp?.Format is not null)
                {
                    Add(operation, SemanticChangeFamily.ParagraphFormatting, part, "paragraph.format",
                        lp?.Anchor.ToString(), rp?.Anchor.ToString(), lp?.Anchor.Scope, rp?.Anchor.Scope, null,
                        lp is null ? SemanticValue.Absent : ParagraphFormatValue(lp.Format),
                        rp is null ? SemanticValue.Absent : ParagraphFormatValue(rp.Format));
                }
                if (lp?.List is not null || rp?.List is not null)
                {
                    Add(operation, SemanticChangeFamily.List, part, "paragraph.list",
                        lp?.Anchor.ToString(), rp?.Anchor.ToString(), lp?.Anchor.Scope, rp?.Anchor.Scope, null,
                        lp is null ? SemanticValue.Absent : ListValue(lp),
                        rp is null ? SemanticValue.Absent : ListValue(rp));
                }
                if (lp is not null)
                {
                    foreach (var span in RunFormatSpans(
                        FlattenInlines(lp.Inlines).OfType<IrTextRun>().ToArray()))
                    {
                        Add(SemanticChangeOperation.Delete, SemanticChangeFamily.RunFormatting,
                            part, $"paragraph.characters[{span.Start}:{span.End}].format",
                            lp.Anchor.ToString(), null, lp.Anchor.Scope, null, null,
                            RunFormatValue(span.Format), SemanticValue.Absent);
                    }
                }
                if (rp is not null)
                {
                    foreach (var span in RunFormatSpans(
                        FlattenInlines(rp.Inlines).OfType<IrTextRun>().ToArray()))
                    {
                        Add(SemanticChangeOperation.Insert, SemanticChangeFamily.RunFormatting,
                            part, $"paragraph.characters[{span.Start}:{span.End}].format",
                            null, rp.Anchor.ToString(), null, rp.Anchor.Scope, null,
                            SemanticValue.Absent, RunFormatValue(span.Format));
                    }
                }
                if (lp?.InlineSectionFormat is not null || rp?.InlineSectionFormat is not null)
                    CompareSection(lp?.InlineSectionFormat, rp?.InlineSectionFormat, part,
                        lp?.Anchor.ToString(), rp?.Anchor.ToString(),
                        lp?.Anchor.Scope ?? leftScope, rp?.Anchor.Scope ?? rightScope);
                var leftFeatures = lp is null ? Array.Empty<InlineFeature>() : InlineFeatures(lp.Inlines).ToArray();
                var rightFeatures = rp is null ? Array.Empty<InlineFeature>() : InlineFeatures(rp.Inlines).ToArray();
                foreach (var feature in leftFeatures)
                    Add(SemanticChangeOperation.Delete, feature.Family, part, $"paragraph.{feature.Path}",
                        lp?.Anchor.ToString(), null, lp?.Anchor.Scope, null, null,
                        feature.Value, SemanticValue.Absent);
                foreach (var feature in rightFeatures)
                    Add(SemanticChangeOperation.Insert, feature.Family, part, $"paragraph.{feature.Path}",
                        null, rp?.Anchor.ToString(), null, rp?.Anchor.Scope, null,
                        SemanticValue.Absent, feature.Value);
                return;
            }

            if (left is IrOpaqueBlock || right is IrOpaqueBlock)
            {
                var opaqueLeft = left as IrOpaqueBlock;
                var opaqueRight = right as IrOpaqueBlock;
                Add(Operation(opaqueLeft, opaqueRight), SemanticChangeFamily.OpaquePackagePart,
                    part, "block.opaque", opaqueLeft?.Anchor.ToString(),
                    opaqueRight?.Anchor.ToString(), opaqueLeft?.Anchor.Scope ?? leftScope,
                    opaqueRight?.Anchor.Scope ?? rightScope, null,
                    BlockValue(opaqueLeft), BlockValue(opaqueRight));
                return;
            }

            if (left is IrTable || right is IrTable)
            {
                var lt = left as IrTable;
                var rt = right as IrTable;
                var operation = Operation(lt, rt);
                Add(operation, SemanticChangeFamily.Table, part, "table",
                    lt?.Anchor.ToString(), rt?.Anchor.ToString(), lt?.Anchor.Scope, rt?.Anchor.Scope, null,
                    BlockValue(lt), BlockValue(rt));
                if (lt is not null)
                    foreach (var row in lt.Rows) EmitWholeRowFeatures(row, null, part);
                if (rt is not null)
                    foreach (var row in rt.Rows) EmitWholeRowFeatures(null, row, part);
                return;
            }

            if (left is IrSectionBreak || right is IrSectionBreak)
            {
                var ls = left as IrSectionBreak;
                var rs = right as IrSectionBreak;
                var operation = Operation(ls, rs);
                Add(operation, SemanticChangeFamily.Section, part, "section",
                    ls?.Anchor.ToString(), rs?.Anchor.ToString(), ls?.Anchor.Scope, rs?.Anchor.Scope, null,
                    ls is null ? SemanticValue.Absent : SectionValue(ls.Format),
                    rs is null ? SemanticValue.Absent : SectionValue(rs.Format));
                Add(operation, SemanticChangeFamily.PageSetup, part, "section.page_setup",
                    ls?.Anchor.ToString(), rs?.Anchor.ToString(), ls?.Anchor.Scope, rs?.Anchor.Scope, null,
                    ls is null ? SemanticValue.Absent : PageSetupValue(ls.Format),
                    rs is null ? SemanticValue.Absent : PageSetupValue(rs.Format));
                return;
            }

            if (left is IrSdtBlock || right is IrSdtBlock)
            {
                var lSdt = left as IrSdtBlock;
                var rSdt = right as IrSdtBlock;
                Add(Operation(lSdt, rSdt), SemanticChangeFamily.ContentControl, part,
                    "content_control", lSdt?.Anchor.ToString(), rSdt?.Anchor.ToString(),
                    lSdt?.Anchor.Scope, rSdt?.Anchor.Scope, null,
                    lSdt is null ? SemanticValue.Absent : BlockValue(lSdt),
                    rSdt is null ? SemanticValue.Absent : BlockValue(rSdt));
                if (lSdt is not null)
                    foreach (var block in lSdt.Blocks)
                        EmitWholeBlockFeatures(block, null, part, lSdt.Anchor.Scope, null);
                if (rSdt is not null)
                    foreach (var block in rSdt.Blocks)
                        EmitWholeBlockFeatures(null, block, part, null, rSdt.Anchor.Scope);
            }
        }

        private void EmitWholeRowFeatures(IrRow? left, IrRow? right, string part)
        {
            Add(Operation(left, right), SemanticChangeFamily.TableRow, part, "table.row",
                left?.Anchor.ToString(), right?.Anchor.ToString(), left?.Anchor.Scope, right?.Anchor.Scope,
                null, RowValue(left), RowValue(right));
            if (left is not null)
            {
                foreach (var cell in left.Cells)
                {
                    Add(SemanticChangeOperation.Delete, SemanticChangeFamily.TableCell, part, "table.cell",
                        cell.Anchor.ToString(), null, cell.Anchor.Scope, null, null,
                        CellValue(cell), SemanticValue.Absent);
                    foreach (var block in cell.Blocks)
                        EmitWholeBlockFeatures(block, null, part, cell.Anchor.Scope, null);
                }
            }
            if (right is not null)
            {
                foreach (var cell in right.Cells)
                {
                    Add(SemanticChangeOperation.Insert, SemanticChangeFamily.TableCell, part, "table.cell",
                        null, cell.Anchor.ToString(), null, cell.Anchor.Scope, null,
                        SemanticValue.Absent, CellValue(cell));
                    foreach (var block in cell.Blocks)
                        EmitWholeBlockFeatures(null, block, part, null, cell.Anchor.Scope);
                }
            }
        }

        private void CompareTableTypedProperties(IrTable left, IrTable right, string part)
        {
            var leftProperties = Child(left.Source.Element, "tblPr");
            var rightProperties = Child(right.Source.Element, "tblPr");
            var leftStyle = AttributeValue(leftProperties, "tblStyle", "val");
            var rightStyle = AttributeValue(rightProperties, "tblStyle", "val");
            if (leftStyle != rightStyle)
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.TableStyle, part,
                    "table.style", left.Anchor.ToString(), right.Anchor.ToString(),
                    left.Anchor.Scope, right.Anchor.Scope, null,
                    SemanticValue.String(leftStyle), SemanticValue.String(rightStyle));
            }
            var leftWidth = ElementValue(leftProperties, "tblW", "w", "type");
            var rightWidth = ElementValue(rightProperties, "tblW", "w", "type");
            if (ValueKey(leftWidth) != ValueKey(rightWidth))
            {
                Add(SemanticChangeOperation.Modify, SemanticChangeFamily.TableWidth, part,
                    "table.width", left.Anchor.ToString(), right.Anchor.ToString(),
                    left.Anchor.Scope, right.Anchor.Scope, null, leftWidth, rightWidth);
            }
        }

        private void CompareCellWidth(IrCell left, IrCell right, string part)
        {
            var before = ElementValue(Child(left.Source.Element, "tcPr"), "tcW", "w", "type");
            var after = ElementValue(Child(right.Source.Element, "tcPr"), "tcW", "w", "type");
            if (ValueKey(before) == ValueKey(after)) return;
            Add(SemanticChangeOperation.Modify, SemanticChangeFamily.TableWidth, part,
                "table.cell.width", left.Anchor.ToString(), right.Anchor.ToString(),
                left.Anchor.Scope, right.Anchor.Scope, null, before, after);
        }

        private void Add(
            SemanticChangeOperation operation,
            SemanticChangeFamily family,
            string partUri,
            string path,
            string? leftAnchor,
            string? rightAnchor,
            string? leftScope,
            string? rightScope,
            string? moveId,
            SemanticValue before,
            SemanticValue after)
        {
            // Some edit-script entries exist only to reconcile transient identifiers or renderer
            // topology. They are useful to the redline pipeline, but they are not semantic Modify
            // records when their canonical values are identical. Callers still project nested ops.
            if (operation == SemanticChangeOperation.Modify && SemanticValuesEqual(before, after))
                return;

            _changes.Add(new SemanticChangeDraft(
                operation, family, NormalizePartUri(partUri), path,
                leftAnchor, rightAnchor, leftScope, rightScope, moveId, before, after));
        }

        private static IrBlock? Find(IrDocument document, string? anchor) =>
            anchor is not null && document.AnchorIndex.TryGetValue(anchor, out var block) ? block : null;

        private static IrRow? FindRow(IrTable table, string? anchor) =>
            anchor is null ? null : table.Rows.FirstOrDefault(row => row.Anchor.ToString() == anchor);

        private static IrHeaderFooter? FindStory(
            IrDocument document,
            bool isHeader,
            Uri? partUri,
            string? scopeName)
        {
            if (partUri is null) return null;
            var stories = isHeader ? document.Headers : document.Footers;
            return stories.FirstOrDefault(story =>
                    story.Scope.PartUri == partUri
                    && (scopeName is null || story.ScopeName == scopeName))
                ?? stories.FirstOrDefault(story => story.Scope.PartUri == partUri);
        }

        private static SemanticChangeOperation NoteOperation(IrScope? left, IrScope? right) =>
            Operation(left, right);

        private static SemanticChangeOperation StoryOperation(IrHeaderFooterDiff story) =>
            story.LeftPartUri is null ? SemanticChangeOperation.Insert
            : story.RightPartUri is null ? SemanticChangeOperation.Delete
            : SemanticChangeOperation.Modify;
    }

    private sealed record RunFormatSpan(int Start, int End, IrRunFormat Format);

    private sealed record RunFormatDelta(
        int Start,
        int End,
        IrRunFormat Before,
        IrRunFormat After);

    private sealed record InlineFeature(
        SemanticChangeFamily Family,
        string Path,
        SemanticValue Value);

    private static string InlineFeatureKey(InlineFeature feature) => string.Join(
        "\u001f",
        ((int)feature.Family).ToString(System.Globalization.CultureInfo.InvariantCulture),
        feature.Path,
        ValueKey(feature.Value));

    private static IEnumerable<InlineFeature> InlineFeatures(IReadOnlyList<IrInline> inlines)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrHyperlink link:
                    yield return new InlineFeature(SemanticChangeFamily.Hyperlink, "hyperlink",
                        Obj(
                            ("target", SemanticValue.String(link.Target)),
                            ("internalTarget", SemanticValue.String(link.InternalTarget?.ToString())),
                            ("text", SemanticValue.String(PlainText(link.Inlines)))));
                    foreach (var nested in InlineFeatures(link.Inlines)) yield return nested;
                    break;
                case IrFieldRun field:
                    yield return new InlineFeature(SemanticChangeFamily.Field, "field",
                        Obj(
                            ("instruction", SemanticValue.String(field.Instruction)),
                            ("cachedResult", SemanticValue.String(PlainText(field.CachedResult))),
                            ("simple", SemanticValue.Boolean(field.IsSimpleField)),
                            ("scaffoldDigest", Digest(field.ScaffoldDigest,
                                "docxodus-ir-field-scaffold-v1"))));
                    foreach (var nested in InlineFeatures(field.CachedResult)) yield return nested;
                    break;
                case IrInlineImage image:
                    yield return new InlineFeature(SemanticChangeFamily.Image, "image",
                        Obj(
                            ("partUri", SemanticValue.String(image.PartUri.ToString())),
                            ("bytesDigest", Digest(image.ImageBytesHash, "raw-media-bytes")),
                            ("drawingDigest", Digest(image.DrawingDigest,
                                "docxodus-ir-drawing-v1")),
                            ("widthEmu", SemanticValue.IntegerFromDocument(image.WidthEmu)),
                            ("heightEmu", SemanticValue.IntegerFromDocument(image.HeightEmu)),
                            ("altText", SemanticValue.String(image.AltText))));
                    break;
                case IrOpaqueInline opaque:
                    yield return new InlineFeature(SemanticChangeFamily.OpaquePackagePart, "opaque_inline",
                        Obj(
                            ("elementName", SemanticValue.String(opaque.ElementName.ToString())),
                            ("digest", Digest(opaque.CanonicalHash,
                                "docxodus-ir-canonical-xml-v1"))));
                    break;
                case IrTextbox textbox:
                    yield return new InlineFeature(SemanticChangeFamily.BlockStructure, "textbox",
                        SemanticValue.Array(textbox.Blocks.Select(BlockValue)));
                    break;
            }
        }
    }

    private static IEnumerable<IrInline> FlattenInlines(IReadOnlyList<IrInline> inlines)
    {
        foreach (var inline in inlines)
        {
            yield return inline;
            if (inline is IrHyperlink link)
                foreach (var nested in FlattenInlines(link.Inlines)) yield return nested;
            else if (inline is IrFieldRun field)
                foreach (var nested in FlattenInlines(field.CachedResult)) yield return nested;
        }
    }

    private static SemanticValue BlockValue(IrBlock? block)
    {
        if (block is null) return SemanticValue.Absent;
        return Obj(
            ("kind", SemanticValue.String(BlockKind(block))),
            ("text", SemanticValue.String(PlainText(block))),
            ("contentDigest", Digest(block.ContentHash, "docxodus-ir-content-v1")),
            ("formatDigest", Digest(block.FormatFingerprint, "docxodus-ir-format-v1")));
    }

    private static string BlockKind(IrBlock block) => block switch
    {
        IrParagraph => "paragraph",
        IrTable => "table",
        IrSdtBlock => "content_control",
        IrSectionBreak => "section",
        IrOpaqueBlock opaque => "opaque:" + opaque.ElementName,
        _ => "unknown",
    };

    private static SemanticValue ScopeValue(IrScope? scope) => scope is null
        ? SemanticValue.Absent
        : Obj(
            ("name", SemanticValue.String(scope.Name)),
            ("partUri", SemanticValue.String(scope.PartUri?.ToString())),
            ("blocks", SemanticValue.Array(scope.Blocks.Select(BlockValue))));

    private static SemanticValue StoryValue(IrHeaderFooter? story) => story is null
        ? SemanticValue.Absent
        : Obj(
            ("partUri", SemanticValue.String(story.Scope.PartUri?.ToString())),
            ("scope", SemanticValue.String(story.ScopeName)),
            ("blocks", SemanticValue.Array(story.Scope.Blocks.Select(BlockValue))),
            ("bindings", SemanticValue.Array(story.References.Select(binding => Obj(
                ("sectionIndex", SemanticValue.Integer(binding.SectionIndex)),
                ("kind", SemanticValue.String(binding.Kind.ToString().ToLowerInvariant())))))));

    private static SemanticValue RowValue(IrRow? row) => row is null
        ? SemanticValue.Absent
        : Obj(
            ("contentDigest", Digest(row.ContentHash, "docxodus-ir-row-content-v1")),
            ("gridBefore", SemanticValue.Integer(row.GridBefore)),
            ("gridAfter", SemanticValue.Integer(row.GridAfter)),
            ("fromContentControl", SemanticValue.Boolean(row.FromTableSdt)),
            ("propertyDigest", Digest(row.TrPrDigest, "docxodus-ir-row-properties-v1")));

    private static SemanticValue RowGridValue(IrRow row) => Obj(
        ("gridBefore", SemanticValue.Integer(row.GridBefore)),
        ("gridAfter", SemanticValue.Integer(row.GridAfter)));

    private static SemanticValue CellValue(IrCell? cell) => cell is null
        ? SemanticValue.Absent
        : Obj(
            ("contentDigest", Digest(cell.ContentHash, "docxodus-ir-cell-content-v1")),
            ("gridSpan", SemanticValue.Integer(cell.GridSpan)),
            ("verticalMerge", SemanticValue.String(cell.VMerge.ToString().ToLowerInvariant())),
            ("fromContentControl", SemanticValue.Boolean(cell.FromRowSdt)),
            ("propertyDigest", Digest(cell.ShellDigest, "docxodus-ir-cell-properties-v1")));

    private static SemanticValue CellSpanValue(IrCell cell) => Obj(
        ("gridSpan", SemanticValue.Integer(cell.GridSpan)),
        ("verticalMerge", SemanticValue.String(cell.VMerge.ToString().ToLowerInvariant())));

    // These projections read int-typed modeled IR state, which cannot leave the v1 safe range, so
    // they call SemanticValue.Integer directly and keep its range check live as an assertion. Only
    // values parsed straight out of document bytes as long go through IntegerFromDocument.
    private static SemanticValue ParagraphFormatValue(IrParaFormat format) => Obj(
        ("styleId", SemanticValue.String(format.StyleId)),
        ("justification", SemanticValue.String(format.Justification?.ToString().ToLowerInvariant())),
        ("indentLeftTwips", SemanticValue.Integer(format.IndentLeftTwips)),
        ("indentRightTwips", SemanticValue.Integer(format.IndentRightTwips)),
        ("indentFirstLineTwips", SemanticValue.Integer(format.IndentFirstLineTwips)),
        ("spacingBeforeTwips", SemanticValue.Integer(format.SpacingBeforeTwips)),
        ("spacingAfterTwips", SemanticValue.Integer(format.SpacingAfterTwips)),
        ("outlineLevel", SemanticValue.Integer(format.OutlineLevel)),
        ("keepNext", SemanticValue.Boolean(format.KeepNext)),
        ("keepLines", SemanticValue.Boolean(format.KeepLines)),
        ("pageBreakBefore", SemanticValue.Boolean(format.PageBreakBefore)),
        ("numId", SemanticValue.Integer(format.NumId)),
        ("listLevel", SemanticValue.Integer(format.Ilvl)),
        ("unmodeledDigest", Digest(format.UnmodeledDigest,
            "docxodus-ir-paragraph-properties-v1")));

    private static SemanticValue RunFormatValue(IrRunFormat? format) => format is null
        ? SemanticValue.Absent
        : Obj(
            ("styleId", SemanticValue.String(format.StyleId)),
            ("bold", SemanticValue.Boolean(format.Bold)),
            ("italic", SemanticValue.Boolean(format.Italic)),
            ("underline", SemanticValue.String(format.Underline?.Kind.ToString().ToLowerInvariant())),
            ("strike", SemanticValue.Boolean(format.Strike)),
            ("doubleStrike", SemanticValue.Boolean(format.DoubleStrike)),
            ("verticalAlign", SemanticValue.String(format.VertAlign?.ToString().ToLowerInvariant())),
            ("fontAscii", SemanticValue.String(format.FontAscii)),
            ("sizeHalfPoints", SemanticValue.Integer(format.SizeHalfPoints)),
            ("color", SemanticValue.String(format.ColorHex)),
            ("highlight", SemanticValue.String(format.Highlight)),
            ("caps", SemanticValue.Boolean(format.Caps)),
            ("smallCaps", SemanticValue.Boolean(format.SmallCaps)),
            ("hidden", SemanticValue.Boolean(format.Vanish)),
            ("unmodeledDigest", Digest(format.UnmodeledDigest,
                "docxodus-ir-run-properties-v1")));

    private static SemanticValue TokenFormatValue(
        IReadOnlyList<IrDiffToken> tokens,
        int start,
        int end) => SemanticValue.Array(tokens.Skip(start).Take(end - start)
            .Select(token => RunFormatValue(token.Format)));

    private static SemanticValue TokenTextValue(
        IReadOnlyList<IrDiffToken> tokens,
        int start,
        int end) => Obj(
            ("text", SemanticValue.String(string.Concat(tokens.Skip(start).Take(end - start)
                .Select(token => token.Text)))),
            ("tokenKinds", SemanticValue.Array(tokens.Skip(start).Take(end - start)
                .Select(token => SemanticValue.String(token.Kind.ToString().ToLowerInvariant())))),
            ("atomicIdentities", SemanticValue.Array(tokens.Skip(start).Take(end - start)
                .Where(token => token.Kind is not (IrDiffTokenKind.Word or IrDiffTokenKind.Separator))
                .Select(token => Obj(
                    ("kind", SemanticValue.String(token.Kind.ToString().ToLowerInvariant())),
                    ("identity", SemanticValue.String(token.MatchKey.TrimStart('\u0001'))))))));

    private static SemanticValue ListValue(IrParagraph paragraph) => paragraph.List is null
        ? Obj(
            ("membership", SemanticValue.Absent),
            ("resolvedMarker", SemanticValue.String(paragraph.ResolvedListMarker)),
            ("layoutListItem", SemanticValue.Boolean(paragraph.IsListItemForLayout)))
        : Obj(
            ("numId", SemanticValue.Integer(paragraph.List.NumId)),
            ("abstractNumId", SemanticValue.Integer(paragraph.List.AbstractNumId)),
            ("level", SemanticValue.Integer(paragraph.List.Ilvl)),
            ("numberFormat", SemanticValue.String(paragraph.List.NumberFormat)),
            ("startOverride", SemanticValue.Integer(paragraph.List.StartOverride)),
            ("fromStyle", SemanticValue.Boolean(paragraph.List.FromStyle)),
            ("resolvedMarker", SemanticValue.String(paragraph.ResolvedListMarker)),
            ("layoutListItem", SemanticValue.Boolean(paragraph.IsListItemForLayout)));

    private static SemanticValue SectionValue(IrSectionFormat? format) => format is null
        ? SemanticValue.Absent
        : Obj(
            ("pageSetup", PageSetupValue(format)),
            ("sectionType", SemanticValue.String(format.SectionType)),
            ("unmodeledDigest", Digest(format.UnmodeledDigest,
                "docxodus-ir-section-properties-v1")));

    private static SemanticValue PageSetupValue(IrSectionFormat? format) => format is null
        ? SemanticValue.Absent
        : Obj(
            ("widthTwips", SemanticValue.Integer(format.PageWidthTwips)),
            ("heightTwips", SemanticValue.Integer(format.PageHeightTwips)),
            ("landscape", SemanticValue.Boolean(format.Landscape)),
            ("marginTopTwips", SemanticValue.Integer(format.MarginTopTwips)),
            ("marginBottomTwips", SemanticValue.Integer(format.MarginBottomTwips)),
            ("marginLeftTwips", SemanticValue.Integer(format.MarginLeftTwips)),
            ("marginRightTwips", SemanticValue.Integer(format.MarginRightTwips)));

    private static SemanticValue StyleValue(IrStyle? style) => style is null
        ? SemanticValue.Absent
        : Obj(
            ("id", SemanticValue.String(style.Id)),
            ("name", SemanticValue.String(style.Name)),
            ("basedOn", SemanticValue.String(style.BasedOn)),
            ("type", SemanticValue.String(style.Type)),
            ("default", SemanticValue.Boolean(style.IsDefault)),
            ("paragraphProperties", SemanticValue.String(CanonicalXml(style.PPr))),
            ("runProperties", SemanticValue.String(CanonicalXml(style.RPr))));

    private static SemanticValue NumValue(IrNum? num) => num is null
        ? SemanticValue.Absent
        : Obj(
            ("numId", SemanticValue.Integer(num.NumId)),
            ("abstractNumId", SemanticValue.Integer(num.AbstractNumId)),
            ("startOverrides", SemanticValue.Array(num.StartOverrides.OrderBy(pair => pair.Key)
                .Select(pair => Obj(
                    ("level", SemanticValue.Integer(pair.Key)),
                    ("start", SemanticValue.Integer(pair.Value)))))));

    private static SemanticValue AbstractNumValue(IrAbstractNum? num) => num is null
        ? SemanticValue.Absent
        : Obj(
            ("abstractNumId", SemanticValue.Integer(num.AbstractNumId)),
            ("levels", SemanticValue.Array(num.Levels.OrderBy(pair => pair.Key).Select(pair => Obj(
                ("level", SemanticValue.Integer(pair.Value.Ilvl)),
                ("format", SemanticValue.String(pair.Value.NumberFormat)),
                ("start", SemanticValue.Integer(pair.Value.Start)),
                ("text", SemanticValue.String(pair.Value.LvlText)),
                ("paragraphProperties", SemanticValue.String(CanonicalXml(pair.Value.PPr))))))));

    private static SemanticValue ThemeValue(IrThemeFonts fonts) => Obj(
        ("majorAscii", SemanticValue.String(fonts.MajorAscii)),
        ("minorAscii", SemanticValue.String(fonts.MinorAscii)));

    private static SemanticValue CommentValue(IrComment comment) => Obj(
        ("author", SemanticValue.String(comment.Author)),
        ("initials", SemanticValue.String(comment.Initials)),
        ("date", SemanticValue.String(comment.Date)),
        ("blocks", SemanticValue.Array(comment.Blocks.Select(BlockValue))),
        ("targets", SemanticValue.Array(comment.Targets.Select(target => Obj(
            ("anchor", SemanticValue.String(target.BlockAnchor.ToString())),
            ("startChar", SemanticValue.Integer(target.StartChar)),
            ("endChar", SemanticValue.Integer(target.EndChar)))))));

    private static SemanticValue TableGridValue(IrTable table)
    {
        var columns = (Child(table.Source.Element, "tblGrid")?.Elements()
                ?? Enumerable.Empty<XElement>())
            .Where(element => element.Name.LocalName == "gridCol")
            .Select(element => SemanticValue.IntegerFromDocument(ParseLong(element.Attributes()
                .FirstOrDefault(attribute => attribute.Name.LocalName == "w")?.Value)));
        return Obj(
            ("columnsTwips", SemanticValue.Array(columns)),
            ("digest", Digest(table.TblGridDigest, "docxodus-ir-table-grid-v1")));
    }

    /// <summary>
    /// The first direct child with this local name. Property projection never sweeps arbitrary
    /// descendants: a nested table carries its own <c>w:tblPr</c>/<c>w:tblGrid</c> and a row carries
    /// <c>w:tblPrEx</c>, so a descendant search would attribute an inner value to the outer table or
    /// cell whose anchor the change record names.
    /// </summary>
    private static XElement? Child(XElement? parent, string localName) =>
        parent?.Elements().FirstOrDefault(element => element.Name.LocalName == localName);

    private static SemanticValue ElementValue(XElement? parent, string child, params string[] attributes)
    {
        var element = Child(parent, child);
        if (element is null) return SemanticValue.Absent;
        return SemanticValue.Object(attributes.Select(name => new SemanticProperty(
            name,
            SemanticValue.String(element.Attributes().FirstOrDefault(attribute =>
                attribute.Name.LocalName == name)?.Value))));
    }

    private static string? AttributeValue(XElement? parent, string child, string attribute)
    {
        var element = Child(parent, child);
        return element?.Attributes().FirstOrDefault(item => item.Name.LocalName == attribute)?.Value;
    }

    private static long? ParseLong(string? value) =>
        long.TryParse(value, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var parsed) ? parsed : null;

    private static string PlainText(IrBlock block) => block switch
    {
        IrParagraph paragraph => PlainText(paragraph),
        IrTable table => string.Join("\n", table.Rows.SelectMany(row => row.Cells)
            .SelectMany(cell => cell.Blocks).Select(PlainText)),
        IrSdtBlock sdt => string.Join("\n", sdt.Blocks.Select(PlainText)),
        _ => string.Empty,
    };

    private static string PlainText(IrParagraph paragraph) => PlainText(paragraph.Inlines);

    private static string PlainText(IReadOnlyList<IrInline> inlines)
    {
        var parts = new List<string>();
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrTextRun run: parts.Add(run.Text); break;
                case IrTab: parts.Add("\t"); break;
                case IrBreak: parts.Add("\n"); break;
                case IrHyperlink link: parts.Add(PlainText(link.Inlines)); break;
                case IrFieldRun field: parts.Add(PlainText(field.CachedResult)); break;
                case IrTextbox textbox: parts.Add(string.Join("\n", textbox.Blocks.Select(PlainText))); break;
            }
        }
        return string.Concat(parts);
    }

    private static SemanticValue AnchorArray(IEnumerable<string?>? anchors) =>
        SemanticValue.Array((anchors ?? Enumerable.Empty<string?>())
            .Where(anchor => anchor is not null)
            .Select(anchor => SemanticValue.String(anchor)));

    private static SemanticValue Digest(IrHash hash, string profile) =>
        SemanticValue.Digest("SHA-256", hash.ToHex(), profile);

    private static SemanticValue Obj(params (string Name, SemanticValue Value)[] properties) =>
        SemanticValue.Object(properties.Select(property =>
            new SemanticProperty(property.Name, property.Value)));

    private static SemanticChangeOperation Operation<T>(T? left, T? right)
        where T : class =>
        left is null ? SemanticChangeOperation.Insert
        : right is null ? SemanticChangeOperation.Delete
        : SemanticChangeOperation.Modify;

    private static bool HasRightSide(IrEditOp operation) =>
        operation.RightAnchor is not null
        || operation.Kind == IrEditOpKind.SplitBlock
        || operation.CrossParagraphCells?.Any(cell => cell.RightAnchor is not null) == true;

    private static string MoveId(
        string kind,
        string partUri,
        string? leftAnchor,
        string? rightAnchor,
        int group) => string.Join(
            ":",
            kind,
            NormalizePartUri(partUri),
            leftAnchor ?? "-",
            rightAnchor ?? "-",
            group.ToString(System.Globalization.CultureInfo.InvariantCulture));

    private static string? Scope(string? anchor)
    {
        if (anchor is null) return null;
        int first = anchor.IndexOf(':');
        if (first < 0) return null;
        int second = anchor.IndexOf(':', first + 1);
        return second < 0 ? null : anchor.Substring(first + 1, second - first - 1);
    }

    private static string NormalizePartUri(string partUri) =>
        partUri.StartsWith("/", StringComparison.Ordinal) ? partUri : "/" + partUri;

    private static string CanonicalXml(XElement? element)
    {
        if (element is null) return string.Empty;
        return new XElement(
            XName.Get(element.Name.LocalName, element.Name.NamespaceName),
            element.Attributes()
                .Where(attribute => !attribute.IsNamespaceDeclaration)
                .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
                .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal)
                .Select(attribute => new XAttribute(
                    XName.Get(attribute.Name.LocalName, attribute.Name.NamespaceName), attribute.Value)),
            element.Nodes().Select(node => node is XElement child
                ? (XNode?)XElement.Parse(CanonicalXml(child), LoadOptions.PreserveWhitespace)
                : node is XText text ? new XText(text.Value) : null))
            .ToString(SaveOptions.DisableFormatting);
    }

    private static bool SemanticValuesEqual(SemanticValue left, SemanticValue right)
    {
        if (ReferenceEquals(left, right)) return true;
        if (left.Kind != right.Kind
            || left.StringValue != right.StringValue
            || left.BooleanValue != right.BooleanValue
            || left.IntegerValue != right.IntegerValue
            || left.DigestAlgorithm != right.DigestAlgorithm
            || left.DigestProfile != right.DigestProfile
            || left.DigestValue != right.DigestValue
            || left.Properties.Count != right.Properties.Count
            || left.Items.Count != right.Items.Count)
            return false;

        for (int index = 0; index < left.Properties.Count; index++)
        {
            var leftProperty = left.Properties[index];
            var rightProperty = right.Properties[index];
            if (leftProperty.Name != rightProperty.Name
                || !SemanticValuesEqual(leftProperty.Value, rightProperty.Value))
                return false;
        }

        for (int index = 0; index < left.Items.Count; index++)
            if (!SemanticValuesEqual(left.Items[index], right.Items[index]))
                return false;

        return true;
    }

    private static string ValueKey(SemanticValue value)
    {
        var change = new SemanticChangeSet(new[]
        {
            new SemanticChange
            {
                Id = "key",
                Operation = SemanticChangeOperation.Modify,
                Family = SemanticChangeFamily.Text,
                PartUri = "/",
                Path = "key",
                Before = value,
                After = SemanticValue.Absent,
            },
        });
        return change.ToJson(indented: false);
    }
}
