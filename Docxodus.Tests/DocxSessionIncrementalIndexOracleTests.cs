// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests;

/// <summary>
/// Differential oracles for issue #1022. After an edit the session refreshes its anchor index only
/// for the blocks whose trees changed, and builds its undo snapshot from frozen copies of the
/// blocks that did not. Both are only correct if they equal what the old whole-document work
/// produced. So after every op in a long mixed sequence (text, structural, notes, comments,
/// headers, styles, lists, tables, undo, redo, rollback, save, projection) over a sample of the
/// test corpus, these tests check that:
/// <list type="bullet">
/// <item>the maintained index equals a fresh full rebuild, key for key in document order, and that
/// rebuild had nothing left to assign (no element missing a Unid, no content-control identity
/// rewritten);</item>
/// <item>a snapshot taken now, materialized, equals a plain deep copy of every snapshot-scoped part.</item>
/// </list>
/// </summary>
public class DocxSessionIncrementalIndexOracleTests
{
    private readonly ITestOutputHelper _output;

    public DocxSessionIncrementalIndexOracleTests(ITestOutputHelper output) => _output = output;

    private static readonly WmlToMarkdownConverterSettings DefaultProjection = new();

    // ── Oracles ───────────────────────────────────────────────────────────

    /// <summary>The first difference between the session's index and a fresh full rebuild, or
    /// null when they agree. Also fails when the rebuild had to assign or rewrite any Unid, which
    /// would mean the incremental refresh left an element unassigned or wrongly assigned.</summary>
    private static string? IndexMismatch(DocxSession session)
    {
        var maintained = session.AnchorIndex();
        var settings = session.ProjectionSettings;
        var before = UnidFingerprint(session.LiveDocument, settings);
        var fresh = WmlToMarkdownConverter.BuildAnchorIndexOnly(session.LiveDocument, settings);
        var after = UnidFingerprint(session.LiveDocument, settings);
        if (before != after) return "a fresh rebuild assigned or rewrote Unids the session had left";

        var a = maintained.Select(Describe).ToList();
        var b = fresh.Select(Describe).ToList();
        for (int i = 0; i < Math.Max(a.Count, b.Count); i++)
        {
            var left = i < a.Count ? a[i] : "<end>";
            var right = i < b.Count ? b[i] : "<end>";
            if (left != right) return $"entry {i}: maintained {left}, rebuilt {right}";
        }
        if (maintained.Count != fresh.Count) return $"count {maintained.Count} vs {fresh.Count}";
        foreach (var key in fresh.Keys)
            if (!maintained.TryGetValue(key, out var hit) || Describe(new(key, hit)) != Describe(new(key, fresh[key])))
                return $"lookup of {key} disagrees";
        return null;

        static string Describe(KeyValuePair<string, AnchorTarget> pair) =>
            $"{pair.Key}|{pair.Value.Anchor.Id}|{pair.Value.Anchor.Kind}|{pair.Value.Anchor.Scope}|{pair.Value.Unid}|{pair.Value.PartUri}";
    }

    /// <summary>Every Unid in every projected part, in document order.</summary>
    private static string UnidFingerprint(WordprocessingDocument doc, WmlToMarkdownConverterSettings settings)
    {
        var sb = new StringBuilder();
        foreach (var (name, part) in WmlToMarkdownConverter.ProjectedScopes(doc.MainDocumentPart!, settings))
        {
            sb.Append('#').Append(name);
            foreach (var el in part.GetXDocument().Root!.DescendantsAndSelf())
                sb.Append(el.Name.LocalName).Append('=').Append((string?)el.Attribute(PtOpenXml.Unid)).Append(';');
        }
        return sb.ToString();
    }

    /// <summary>The first snapshot-scoped part whose snapshot, materialized, differs from a deep
    /// copy of the live part, or null.</summary>
    private static string? SnapshotMismatch(DocxSession session)
    {
        var snapshot = session.TakeSnapshot();
        var live = PartsByUri(session.LiveDocument);
        foreach (var part in snapshot.Parts)
        {
            if (!live.TryGetValue(part.PartUri, out var livePart)) return $"snapshot has {part.PartUri}, the package does not";
            var expected = new XDocument(livePart.GetXDocument());
            var actual = part.Materialize();
            if (!XNode.DeepEquals(expected, actual)) return $"{part.PartUri} differs from a deep copy";
            if (expected.Declaration?.ToString() != actual.Declaration?.ToString()) return $"{part.PartUri} declaration differs";
        }
        return null;
    }

    private static Dictionary<string, OpenXmlPart> PartsByUri(WordprocessingDocument doc)
    {
        var result = new Dictionary<string, OpenXmlPart>(StringComparer.Ordinal);
        var pending = new Stack<OpenXmlPart>(doc.Parts.Select(p => p.OpenXmlPart));
        while (pending.Count > 0)
        {
            var part = pending.Pop();
            if (!result.TryAdd(part.Uri.ToString(), part)) continue;
            foreach (var child in part.Parts) pending.Push(child.OpenXmlPart);
        }
        return result;
    }

    private static void AssertOracles(DocxSession session, string context)
    {
        if (IndexMismatch(session) is { } index) Assert.Fail($"{context}: anchor index: {index}");
        if (SnapshotMismatch(session) is { } snapshot) Assert.Fail($"{context}: snapshot: {snapshot}");
    }

    // ── Op mix ────────────────────────────────────────────────────────────

    private static List<string> Anchors(DocxSession session, params string[] kinds) =>
        session.AnchorIndex().Values
            .Where(t => kinds.Contains(t.Anchor.Kind))
            .Select(t => t.Anchor.Id)
            .Distinct()
            .ToList();

    /// <summary>Apply one pseudo-random op. Failures are fine: a refused op must leave both
    /// caches consistent too.</summary>
    /// <summary>The result of the last op <see cref="ApplyRandomOp"/> applied, when it returns one.</summary>
    [ThreadStatic]
    private static EditResult? LastResult;

    /// <summary>Every patch the last op returned, in order (several for a batch).</summary>
    private static IEnumerable<MarkdownPatch> LastPatches()
    {
        if (LastBatch is { } batch)
        {
            foreach (var step in batch.Steps)
                foreach (var result in step.Results)
                    if (result.Patch is { } stepPatch) yield return stepPatch;
        }
        else if (LastResult?.Patch is { } patch)
        {
            yield return patch;
        }
    }

    [ThreadStatic]
    private static MutationBatchResult? LastBatch;

    private static string ApplyRandomOp(DocxSession session, Random random, int step, bool historyOps = true)
    {
        LastResult = null;
        LastBatch = null;
        var paragraphs = Anchors(session, "p", "h", "li");
        string Pick(List<string> from) => from[random.Next(from.Count)];
        if (paragraphs.Count == 0)
        {
            if (!historyOps) return "no paragraphs";
            session.Undo();
            return "undo (no paragraphs)";
        }
        var p = Pick(paragraphs);
        var choice = random.Next(31);
        // Without history ops: no undo, redo, rollback, save, projection, batch or failing op — a
        // plain edit instead. Story-creating edits participate too: undo/redo must restore their
        // topology and every later edit to those parts (issue #1033).
        if (!historyOps && choice is 6 or 7 or 8 or 10 or 11 or 24 or 28 or 29) choice = 0;
        switch (choice)
        {
            case 0:
            case 1:
                LastResult = session.ReplaceText(p, $"Edited paragraph {step}.");
                return $"ReplaceText {p}";
            case 2:
                LastResult = session.ReplaceTextAtSpan(p, 0, 0, $"[{step}]");
                return $"ReplaceTextAtSpan {p}";
            case 3:
                LastResult = session.InsertParagraph(p, random.Next(2) == 0 ? Position.Before : Position.After, $"Inserted {step}");
                return $"InsertParagraph {p}";
            case 4:
                LastResult = session.DeleteBlock(p);
                return $"DeleteBlock {p}";
            case 5:
                LastResult = session.ApplyFormat(p, null, new FormatOp { Bold = true });
                return $"ApplyFormat {p}";
            case 6:
                session.Undo();
                return "Undo";
            case 7:
                session.Redo();
                return "Redo";
            case 8:
            {
                using var transaction = session.BeginTransaction();
                session.ReplaceText(p, $"Rolled back {step}");
                session.InsertParagraph(p, Position.After, "Rolled back too");
                transaction.Rollback();
                return $"rolled-back transaction {p}";
            }
            case 9:
                // Selective (non-package) transaction inside one op.
                LastResult = session.ReplaceTextAtSpanWithFormat(p, 0, 0, "Y", new FormatOp { Italic = true });
                return $"ReplaceTextAtSpanWithFormat {p}";
            case 10:
                session.Save();
                return "Save";
            case 11:
                session.Project();
                return "Project";
            case 12:
                LastResult = session.InsertFootnote(p, 0, $"Note {step}");
                return $"InsertFootnote {p}";
            case 13:
                LastResult = session.AddComment(p, null, "Oracle", $"Comment {step}");
                return $"AddComment {p}";
            case 14:
                LastResult = session.SetHeaderText(p, HeaderFooterKind.Default, $"Header {step}");
                return $"SetHeaderText {p}";
            case 15:
                LastResult = session.SetParagraphStyle(p, random.Next(2) == 0 ? "Heading1" : "Normal");
                return $"SetParagraphStyle {p}";
            case 16:
                LastResult = session.ApplyListFormat(p, random.Next(2) == 0 ? ListFormat.Decimal : ListFormat.None);
                return $"ApplyListFormat {p}";
            case 17:
                LastResult = session.MoveBlock(p, Pick(paragraphs), random.Next(2) == 0 ? Position.Before : Position.After);
                return $"MoveBlock {p}";
            case 18:
            {
                var cells = Anchors(session, "tc");
                if (cells.Count > 0)
                {
                    var cell = Pick(cells);
                    LastResult = session.InsertTableRow(cell, Position.After);
                    return $"InsertTableRow {cell}";
                }
                LastResult = session.InsertTable(p, Position.After, 2, 2);
                return $"InsertTable {p}";
            }
            case 19:
                LastResult = session.SplitParagraph(p, 0);
                return $"SplitParagraph {p}";
            case 20:
            {
                var controls = Anchors(session, "sdt");
                if (controls.Count == 0) goto case 0;
                var control = Pick(controls);
                LastResult = session.FillContentControlText(control, $"Filled {step}");
                return $"FillContentControlText {control}";
            }
            case 21:
                LastResult = session.AcceptAllRevisions();
                return "AcceptAllRevisions";
            case 22:
                LastResult = session.RejectAllRevisions();
                return "RejectAllRevisions";
            case 23:
            {
                var to = Pick(paragraphs);
                LastResult = session.DeleteRange(p, to);
                return $"DeleteRange {p}..{to}";
            }
            case 24:
            {
                var comments = Anchors(session, "cmt");
                if (comments.Count == 0) goto case 13;
                var comment = Pick(comments);
                LastResult = session.RemoveComment(comment);
                return $"RemoveComment {comment}";
            }
            case 25:
                LastResult = session.InsertImage(p, 0, TinyPng());
                return $"InsertImage {p}";
            case 26:
                LastResult = session.Raw.ReplaceXml(p,
                    "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Raw " + step + "</w:t></w:r></w:p>");
                return $"Raw.ReplaceXml {p}";
            case 28:
            {
                var second = Pick(paragraphs);
                LastBatch = session.ExecuteBatch(new[]
                {
                    new MutationBatchStep("docx_edit", "replace_text", s => s.ReplaceText(p, $"Batch one {step}.")),
                    new MutationBatchStep("docx_edit", "replace_text", s => s.ReplaceText(second, $"Batch two {step}.")),
                }, MutationBatchMode.BestEffort);
                return $"ExecuteBatch {p}, {second}";
            }
            case 29:
                // Invalid XML text is refused before mutation, with either patch setting.
                LastResult = session.ReplaceText(p, "bad\ud800payload");
                return $"failing ReplaceText {p}";
            case 30:
                // A mutation that returns no patch: the next patch must still cover it.
                LastResult = session.AddHyperlink(p, new CharSpan(0, 1), HyperlinkTarget.External("https://example.com/" + step));
                LastResult = null;
                return $"AddHyperlink {p}";
            default:
                // Inline code finds or creates a character style: a styles-part change.
                LastResult = session.ApplyFormat(p, null, new FormatOp { Code = true });
                return $"ApplyFormat code {p}";
        }
    }

    /// <summary>Run the mix, checking both oracles after every op. Returns how many ops changed the
    /// document (its version moved), so a caller can prove the run was not vacuous.</summary>
    private static int RunSequence(byte[] bytes, DocxSessionSettings settings, int seed, int steps, string label)
    {
        using var session = new DocxSession(bytes, settings);
        var random = new Random(seed);
        AssertOracles(session, $"{label} at open");
        int changed = 0;
        for (int step = 0; step < steps; step++)
        {
            var version = session.Version;
            var op = ApplyRandomOp(session, random, step);
            if (session.Version != version) changed++;
            AssertOracles(session, $"{label} step {step} ({op})");
        }
        return changed;
    }

    private static IEnumerable<FileInfo> CorpusSample()
    {
        var directory = new DirectoryInfo("../../../../TestFiles/");
        var files = directory.EnumerateFiles("*.docx", SearchOption.AllDirectories)
            .Where(f => f.Length < 1_000_000)
            .OrderBy(f => f.FullName, StringComparer.Ordinal)
            .ToList();
        // Every k-th file: a deterministic sample spread across every fixture family.
        int stride = Math.Max(1, files.Count / 48);
        for (int i = 0; i < files.Count; i += stride) yield return files[i];
    }

    private static int StableSeed(string name)
    {
        int hash = 17;
        foreach (var c in name) hash = unchecked(hash * 31 + c);
        return hash & 0xffff;
    }

    private static DocxSession? TryOpen(byte[] bytes, DocxSessionSettings settings)
    {
        try { return new DocxSession(bytes, settings); }
        catch (Exception) { return null; }
    }

    // ── Tests ─────────────────────────────────────────────────────────────

    [Fact]
    public void Corpus_incremental_index_and_shared_snapshots_match_full_rebuilds_after_every_op()
    {
        var settings = new DocxSessionSettings { EmitMarkdownPatch = false };
        int documents = 0, changedOps = 0;
        IncrementalAnchorIndex.RefreshesForTests = 0;
        IncrementalAnchorIndex.FallbacksForTests = 0;
        foreach (var file in CorpusSample())
        {
            var bytes = File.ReadAllBytes(file.FullName);
            using (var probe = TryOpen(bytes, settings))
            {
                if (probe is null) continue;
                try { probe.AnchorIndex(); }
                catch (Exception) { continue; }
            }
            changedOps += RunSequence(bytes, settings, seed: StableSeed(file.Name), steps: 24, file.Name);
            documents++;
        }
        _output.WriteLine($"{documents} documents, {changedOps} document-changing ops, " +
            $"{IncrementalAnchorIndex.RefreshesForTests} incremental refreshes, " +
            $"{IncrementalAnchorIndex.FallbacksForTests} fallbacks to a full rebuild");
        // The oracle must have exercised the incremental path, not only the fallback.
        Assert.True(IncrementalAnchorIndex.RefreshesForTests >= changedOps / 3,
            $"only {IncrementalAnchorIndex.RefreshesForTests} incremental refreshes");
        Assert.True(documents >= 30, $"only {documents} corpus documents opened");
        Assert.True(changedOps >= documents * 8, $"only {changedOps} ops changed a document");
    }

    /// <summary>With patches on, each op's projection carries its own full index and the session
    /// keeps no incremental one beside it; the snapshot side is shared all the same.</summary>
    [Fact]
    public void Patch_on_sessions_keep_both_caches_consistent()
    {
        var settings = new DocxSessionSettings();
        foreach (var file in CorpusSample().Take(8))
        {
            var bytes = File.ReadAllBytes(file.FullName);
            using (var probe = TryOpen(bytes, settings)) { if (probe is null) continue; }
            RunSequence(bytes, settings, seed: 1022, steps: 12, file.Name);
        }
    }

    // ── Block-scoped patches ─────────────────────────────────────────────

    /// <summary>What a client keeping one entry per block holds: per scope, each top-level block's
    /// anchor id and markdown, in order — built here from scratch.</summary>
    private static Dictionary<string, List<(string Id, string Markdown)>> BlockModel(DocxSession session)
    {
        // Deliberately no AnchorIndex() call: building the model must not refresh the session's
        // index, or it would mask an index the session failed to keep current.
        var settings = session.ProjectionSettings;
        var model = new Dictionary<string, List<(string, string)>>(StringComparer.Ordinal);
        foreach (var (name, part) in WmlToMarkdownConverter.ProjectedScopes(session.LiveDocument.MainDocumentPart!, settings))
        {
            var list = new List<(string, string)>();
            var root = part.GetXDocument().Root!;
            // A restored tree has no owner annotation until the production index/projection
            // prepares it. The independent emitter needs that owner to resolve hyperlinks too.
            root.RemoveAnnotations<OpenXmlPart>();
            root.AddAnnotation(part);
            // The projection leaves out a header or footer with no text at all.
            bool shown = !(name.StartsWith("hdr", StringComparison.Ordinal) || name.StartsWith("ftr", StringComparison.Ordinal))
                || root.Descendants(W.t).Any(t => !string.IsNullOrWhiteSpace(t.Value));
            if (shown)
            {
                foreach (var block in PartChangeTracker.ContainerOf(part.GetXDocument())!.Elements())
                {
                    if (WmlToMarkdownConverter.BlockAnchorId(block, name) is not { } id) continue;
                    var markdown = WmlToMarkdownConverter.EmitBlockMarkdown(session.LiveDocument, settings, name, block);
                    if (markdown.Length > 0) list.Add((id, markdown));
                }
            }
            model[name] = list;
        }
        return model;
    }

    /// <summary>The model checked against the full projection, emitted independently: every block's
    /// markdown must occur in it, in model order. A failure message, or null.</summary>
    private static string? ModelVersusProjection(Dictionary<string, List<(string Id, string Markdown)>> model, string projection)
    {
        int at = 0;
        foreach (var (scope, blocks) in model)
        {
            foreach (var (id, markdown) in blocks)
            {
                var found = projection.IndexOf(markdown, at, StringComparison.Ordinal);
                if (found < 0) return $"{scope} block {id} is not where the projection has it";
                at = found + markdown.Length;
            }
        }
        return null;
    }

    /// <summary>Apply a block-scoped patch the way a client would; a failure message, or null.</summary>
    private static string? ApplyPatch(Dictionary<string, List<(string Id, string Markdown)>> model, MarkdownPatch patch)
    {
        foreach (var id in patch.RemovedAnchorIds)
            foreach (var list in model.Values)
                list.RemoveAll(b => b.Id == id);
        foreach (var block in patch.Blocks)
        {
            var scope = block.AnchorId.Split(':')[1];
            if (!model.TryGetValue(scope, out var list)) model[scope] = list = new List<(string, string)>();
            list.RemoveAll(b => b.Id == block.AnchorId);
            int at = 0;
            if (block.AfterAnchorId is not null)
            {
                at = list.FindIndex(b => b.Id == block.AfterAnchorId) + 1;
                if (at == 0) return $"{block.AnchorId} follows {block.AfterAnchorId}, which the client does not have";
            }
            list.Insert(at, (block.AnchorId, block.Markdown));
        }
        return null;
    }

    private static string? ModelMismatch(Dictionary<string, List<(string Id, string Markdown)>> model,
        Dictionary<string, List<(string Id, string Markdown)>> fresh)
    {
        foreach (var (scope, blocks) in fresh)
        {
            model.TryGetValue(scope, out var held);
            held ??= new List<(string, string)>();
            for (int i = 0; i < Math.Max(held.Count, blocks.Count); i++)
            {
                var a = i < held.Count ? held[i] : ("<end>", "");
                var b = i < blocks.Count ? blocks[i] : ("<end>", "");
                if (a != b) return $"{scope} block {i}: client has {a.Item1} \"{Trim(a.Item2)}\", document has {b.Item1} \"{Trim(b.Item2)}\"";
            }
        }
        foreach (var scope in model.Keys)
            if (!fresh.ContainsKey(scope) && model[scope].Count > 0) return $"client still has scope {scope}";
        return null;

        static string Trim(string text) => text.Length > 80 ? text[..80] + "…" : text;
    }

    /// <summary>A client that applies every patch it receives holds, after each one, exactly the
    /// blocks a fresh projection has — whatever ops (including ones that return no patch, and undo,
    /// redo, rollback and save) ran in between.</summary>
    private static (int Scoped, int Full) RunPatchSequence(byte[] bytes, DocxSessionSettings settings, int seed, int steps, string label)
    {
        using var session = new DocxSession(bytes, settings);
        var random = new Random(seed);
        // A client starts from the projection, which is what assigns the opening Unids.
        session.Project();
        var model = BlockModel(session);
        int scoped = 0, full = 0;
        var trace = new List<string>();
        for (int step = 0; step < steps; step++)
        {
            var op = ApplyRandomOp(session, random, step);
            trace.Add($"{step} {op}: {LastResult?.Success} {LastResult?.Error?.Code}; patch={LastResult?.Patch?.IsFullDocument} blocks={LastResult?.Patch?.Blocks.Count}");
            // Undo, redo and rollback return no patch and move the document through its history;
            // a mirroring client re-reads the projection after them, as the patch contract says.
            if (op is "Undo" or "Redo" || op.StartsWith("rolled-back", StringComparison.Ordinal)
                || op.StartsWith("undo", StringComparison.Ordinal))
            {
                model = BlockModel(session);
                continue;
            }
            bool any = false;
            foreach (var patch in LastPatches())
            {
                any = true;
                if (patch.IsFullDocument)
                {
                    full++;
                    model = BlockModel(session);
                    continue;
                }
                scoped++;
                Assert.Equal(string.Concat(patch.Blocks.Select(b => b.Markdown)), patch.Markdown);
                if (ApplyPatch(model, patch) is { } applyError) Assert.Fail($"{label} step {step} ({op}): {applyError}");
            }
            if (!any) continue;
            var fresh = BlockModel(session);
            if (ModelMismatch(model, fresh) is { } mismatch) Assert.Fail($"{label} step {step} ({op}): {mismatch}\n" + string.Join("\n", trace));
            var projection = WmlToMarkdownConverter.Convert(session.LiveDocument, settings.ProjectionSettings).Markdown;
            if (ModelVersusProjection(fresh, projection) is { } projectionMismatch)
                Assert.Fail($"{label} step {step} ({op}): {projectionMismatch}");
            var last = LastPatches().Last();
            if (last.IsFullDocument) Assert.Equal(projection, last.Markdown);
        }
        return (scoped, full);
    }

    [Fact]
    public void Corpus_block_scoped_patches_keep_a_client_in_sync_after_every_op()
    {
        var settings = new DocxSessionSettings();
        int documents = 0, scoped = 0, full = 0;
        foreach (var file in CorpusSample())
        {
            var bytes = File.ReadAllBytes(file.FullName);
            using (var probe = TryOpen(bytes, settings))
            {
                if (probe is null) continue;
                try { probe.AnchorIndex(); } catch (Exception) { continue; }
            }
            var (s, f) = RunPatchSequence(bytes, settings, seed: StableSeed(file.Name) ^ 0x5a5a, steps: 24, file.Name);
            scoped += s;
            full += f;
            documents++;
        }
        _output.WriteLine($"{documents} documents, {scoped} block-scoped patches, {full} whole-document patches");
        Assert.True(documents >= 30, $"only {documents} documents");
        // The mix is heavy on ops that need the whole document (undo, redo, rollback, save, list,
        // style, note, comment and header changes); the plain edits between them must still be scoped.
        Assert.True(scoped > full, $"{scoped} block-scoped patches vs {full} whole-document ones");
    }

    [Theory]
    [InlineData(EmptyParagraphMode.Suppress, AnchorRenderMode.BlockAndInline)]
    [InlineData(EmptyParagraphMode.MarkedEmpty, AnchorRenderMode.None)]
    public void Block_scoped_patches_hold_under_other_projection_settings(EmptyParagraphMode empty, AnchorRenderMode anchors)
    {
        var settings = new DocxSessionSettings
        {
            ProjectionSettings = new WmlToMarkdownConverterSettings { EmptyParagraphs = empty, AnchorMode = anchors },
        };
        foreach (var file in CorpusSample().Take(12))
        {
            var bytes = File.ReadAllBytes(file.FullName);
            using (var probe = TryOpen(bytes, settings)) { if (probe is null) continue; }
            RunPatchSequence(bytes, settings, seed: 1033, steps: 16, file.Name);
        }
    }

    [Fact]
    public void A_text_edit_patch_carries_only_the_edited_block()
    {
        using var session = new DocxSession(Paragraphs(200));
        var anchors = Anchors(session, "p");
        var result = session.ReplaceText(anchors[50], "Only this changed.");
        Assert.True(result.Success);
        var patch = result.Patch!;
        Assert.False(patch.IsFullDocument);
        var block = Assert.Single(patch.Blocks);
        Assert.Equal(anchors[50], block.AnchorId);
        Assert.Equal(anchors[49], block.AfterAnchorId);
        Assert.Contains("Only this changed.", patch.Markdown);
        Assert.DoesNotContain("Paragraph 49", patch.Markdown);
        Assert.Empty(patch.RemovedAnchorIds);

        var inserted = session.InsertParagraph(anchors[10], Position.After, "New one.");
        Assert.Equal(anchors[10], Assert.Single(inserted.Patch!.Blocks).AfterAnchorId);
        var deleted = session.DeleteBlock(anchors[20]);
        Assert.Empty(deleted.Patch!.Blocks);
        Assert.Equal(anchors[20], Assert.Single(deleted.Patch.RemovedAnchorIds));
    }

    /// <summary>A patch depends on the op and the state it ran on, not on history: after an undo,
    /// the next op's patch is that op's own blocks, the same patch a fresh session over the
    /// restored document would produce.</summary>
    [Fact]
    public void The_patch_after_an_undo_covers_only_the_next_op()
    {
        var bytes = Paragraphs(20);
        using var session = new DocxSession(bytes);
        var anchors = Anchors(session, "p");
        Assert.False(session.ReplaceText(anchors[3], "One.").Patch!.IsFullDocument);
        Assert.True(session.Undo());
        var after = session.ReplaceText(anchors[4], "Two.").Patch!;
        Assert.False(after.IsFullDocument);
        Assert.Equal(anchors[4], Assert.Single(after.Blocks).AnchorId);

        using var fresh = new DocxSession(bytes);
        var same = fresh.ReplaceText(anchors[4], "Two.").Patch!;
        Assert.Equal(same.Blocks, after.Blocks);
        Assert.Equal(same.Markdown, after.Markdown);
    }

    /// <summary>The everyday agent flow — project, then edit — gets a scoped patch: projecting
    /// with patches on also brings up the index that watches the edit.</summary>
    [Fact]
    public void Projecting_then_editing_gets_a_scoped_patch()
    {
        using var session = new DocxSession(Paragraphs(20));
        var anchor = session.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:body:", StringComparison.Ordinal));
        var patch = session.ReplaceText(anchor, "Edited.").Patch!;
        Assert.False(patch.IsFullDocument);
        Assert.Equal(anchor, Assert.Single(patch.Blocks).AnchorId);
        Assert.True(session.Save().Length > 0);
        // Save leaves the content as it was, so nothing is owed and the next patch stays scoped.
        Assert.False(session.ReplaceText(anchor, "Again.").Patch!.IsFullDocument);
    }

    [Fact]
    public void Tracked_change_sessions_stay_consistent()
    {
        var settings = new DocxSessionSettings
        {
            EmitMarkdownPatch = false,
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Oracle",
        };
        foreach (var file in CorpusSample().Skip(10).Take(12))
        {
            var bytes = File.ReadAllBytes(file.FullName);
            using (var probe = TryOpen(bytes, settings)) { if (probe is null) continue; }
            RunSequence(bytes, settings, seed: 965, steps: 16, file.Name);
        }
    }

    /// <summary>Shared frozen blocks must never change under an older snapshot: after a run of
    /// edits, each undo restores exactly the XML the document had before that edit, and each redo
    /// exactly the XML it had after.</summary>
    [Fact]
    public void Undo_and_redo_restore_exactly_the_recorded_states()
    {
        var settings = new DocxSessionSettings { EmitMarkdownPatch = false, UndoDepth = 64 };
        int checkedFiles = 0;
        foreach (var file in CorpusSample().Take(16))
        {
            var bytes = File.ReadAllBytes(file.FullName);
            using var session = TryOpen(bytes, settings);
            if (session is null) continue;
            try { session.AnchorIndex(); } catch (Exception) { continue; }
            var random = new Random(StableSeed(file.Name));
            // Mirror the session's history: the state before each recorded edit, and the state
            // each redo should bring back. Edits after undos are what would expose a snapshot
            // whose frozen blocks had been aliased into the live tree and then edited.
            var undo = new Stack<string>();
            var redo = new Stack<string>();
            for (int step = 0; step < 30; step++)
            {
                int roll = random.Next(10);
                if (roll < 2 && undo.Count > 0)
                {
                    var current = State(session);
                    Assert.True(session.Undo(), $"{file.Name} step {step}: undo");
                    Assert.True(undo.Pop() == State(session), $"{file.Name} step {step}: undo restored a different state");
                    redo.Push(current);
                }
                else if (roll < 3 && redo.Count > 0)
                {
                    var current = State(session);
                    Assert.True(session.Redo(), $"{file.Name} step {step}: redo");
                    Assert.True(redo.Pop() == State(session), $"{file.Name} step {step}: redo restored a different state");
                    undo.Push(current);
                }
                else
                {
                    var before = State(session);
                    var version = session.Version;
                    ApplyRandomOp(session, random, step, historyOps: false);
                    if (session.Version != version)
                    {
                        undo.Push(before);
                        redo.Clear();
                    }
                }
            }
            while (undo.Count > 0)
            {
                Assert.True(session.Undo(), $"{file.Name}: final undo");
                Assert.True(undo.Pop() == State(session), $"{file.Name}: final undo restored a different state");
            }
            checkedFiles++;
        }
        Assert.True(checkedFiles >= 10, $"only {checkedFiles} files");

        static string State(DocxSession session)
        {
            // Index first: a lookup assigns Unids to new elements, and the next edit's undo snapshot
            // is taken after its own lookup, so that is the state an undo must come back to.
            session.AnchorIndex();
            var sb = new StringBuilder();
            // Parts in their role order, without URIs: redo re-creates a part the undo removed
            // under the same relationship id, which the package may give a different name.
            var live = PartsByUri(session.LiveDocument);
            foreach (var part in session.TakeSnapshot().Parts)
                sb.Append(live[part.PartUri].GetXDocument().ToString(SaveOptions.DisableFormatting)).Append('\n');
            return sb.ToString();
        }
    }

    [Fact]
    public void A_session_without_an_incremental_index_does_not_record_index_changes()
    {
        var settings = new DocxSessionSettings
        {
            EmitMarkdownPatch = false,
            ProjectionSettings = new WmlToMarkdownConverterSettings { AnchorIdRendering = AnchorIdRendering.Sequential },
        };
        using var session = new DocxSession(Paragraphs(50), settings);
        foreach (var anchor in Anchors(session, "p").Take(10))
            Assert.True(session.ReplaceText(anchor, "Changed.").Success);
        var tracker = PartChangeTracker.Existing(session.LiveDocument.MainDocumentPart!.GetXDocument());
        Assert.NotNull(tracker);
        Assert.False(tracker!.ForIndex.Active);
        Assert.Empty(tracker.ForIndex.Blocks);
    }

    [Theory]
    [InlineData(AnchorIdRendering.Abbreviated, EmptyParagraphMode.AnchorOnly)]
    [InlineData(AnchorIdRendering.FullUnid, EmptyParagraphMode.Suppress)]
    public void Other_projection_settings_stay_consistent(AnchorIdRendering rendering, EmptyParagraphMode empty)
    {
        var settings = new DocxSessionSettings
        {
            EmitMarkdownPatch = false,
            ProjectionSettings = new WmlToMarkdownConverterSettings { AnchorIdRendering = rendering, EmptyParagraphs = empty },
        };
        foreach (var file in CorpusSample().Take(10))
        {
            var bytes = File.ReadAllBytes(file.FullName);
            using (var probe = TryOpen(bytes, settings)) { if (probe is null) continue; }
            RunSequence(bytes, settings, seed: 7, steps: 12, file.Name);
        }
    }

    [Fact]
    public void Edits_after_the_first_refresh_the_index_without_a_rebuild()
    {
        using var session = new DocxSession(Paragraphs(50), new DocxSessionSettings { EmitMarkdownPatch = false });
        var first = session.AnchorIndex();
        Assert.IsType<IncrementalAnchorIndex>(first);
        var anchor = Anchors(session, "p")[10];
        Assert.True(session.ReplaceText(anchor, "Changed.").Success);
        Assert.True(session.InsertParagraph(anchor, Position.After, "Added.").Success);
        Assert.Same(first, session.AnchorIndex());
        Assert.Null(IndexMismatch(session));
    }

    [Fact]
    public void Undo_redo_and_rollback_rebuild_the_index_from_scratch()
    {
        using var session = new DocxSession(Paragraphs(20), new DocxSessionSettings { EmitMarkdownPatch = false });
        var anchor = Anchors(session, "p")[3];
        Assert.True(session.ReplaceText(anchor, "Changed.").Success);
        var beforeUndo = session.AnchorIndex();
        Assert.True(session.Undo());
        Assert.NotSame(beforeUndo, session.AnchorIndex());
        Assert.Null(IndexMismatch(session));
        var beforeRedo = session.AnchorIndex();
        Assert.True(session.Redo());
        Assert.NotSame(beforeRedo, session.AnchorIndex());
        Assert.Null(IndexMismatch(session));
        var beforeRollback = session.AnchorIndex();
        using (var transaction = session.BeginTransaction())
        {
            session.ReplaceText(anchor, "Speculative.");
            transaction.Rollback();
        }
        Assert.NotSame(beforeRollback, session.AnchorIndex());
        Assert.Null(IndexMismatch(session));
    }

    [Fact]
    public void Index_oracle_detects_a_block_the_refresh_skipped()
    {
        using var session = new DocxSession(Paragraphs(20), new DocxSessionSettings { EmitMarkdownPatch = false });
        var anchor = Anchors(session, "p")[5];
        Assert.True(session.ReplaceText(anchor, "Changed text.").Success);
        IncrementalAnchorIndex.SkipOneBlockForTests = true;
        try
        {
            Assert.NotNull(IndexMismatch(session));
        }
        finally
        {
            IncrementalAnchorIndex.SkipOneBlockForTests = false;
        }
    }

    [Fact]
    public void Snapshot_oracle_detects_a_frozen_block_that_was_not_evicted()
    {
        using var session = new DocxSession(Paragraphs(20), new DocxSessionSettings { EmitMarkdownPatch = false });
        var anchor = Anchors(session, "p")[5];
        Assert.Null(SnapshotMismatch(session));
        Assert.True(session.ReplaceText(anchor, "Changed text.").Success);
        PartSnapshotCache.SkipOneEvictionForTests = true;
        try
        {
            Assert.NotNull(SnapshotMismatch(session));
        }
        finally
        {
            PartSnapshotCache.SkipOneEvictionForTests = false;
        }
    }

    [Fact]
    public void A_text_edit_snapshot_shares_every_untouched_block_with_the_previous_one()
    {
        using var session = new DocxSession(Paragraphs(400), new DocxSessionSettings { EmitMarkdownPatch = false });
        var anchors = Anchors(session, "p");
        var first = session.TakeSnapshot();
        Assert.True(session.ReplaceText(anchors[200], "Changed.").Success);
        var second = session.TakeSnapshot();
        var main = (first.Parts[0], second.Parts[0]);
        var before = main.Item1.Chunks.SelectMany(c => c.Blocks).ToList();
        var after = main.Item2.Chunks.SelectMany(c => c.Blocks).ToList();
        Assert.Equal(before.Count, after.Count);
        Assert.Equal(1, before.Zip(after).Count(pair => !ReferenceEquals(pair.First, pair.Second)));
        // Only the chunk holding the edited block is new.
        Assert.Equal(1, main.Item2.Chunks.Count(c => !main.Item1.Chunks.Contains(c)));
        Assert.Same(main.Item1.Shell, main.Item2.Shell);
    }

    [Fact]
    public void The_first_snapshot_after_an_undo_reuses_the_restored_one()
    {
        using var session = new DocxSession(Paragraphs(400), new DocxSessionSettings { EmitMarkdownPatch = false });
        var anchors = Anchors(session, "p");
        Assert.True(session.ReplaceText(anchors[100], "Changed.").Success);
        Assert.True(session.Undo());
        var restored = session.TakeSnapshot().Parts[0];
        Assert.True(session.ReplaceText(anchors[300], "Changed too.").Success);
        var next = session.TakeSnapshot().Parts[0];
        // The restored tree is new live nodes, seeded with the snapshot's frozen blocks: one chunk
        // (the edited block's) is new, and the rest are the very chunks undo restored from.
        Assert.Equal(1, next.Chunks.Count(c => !restored.Chunks.Contains(c)));
    }

    [Fact]
    public void Block_scoped_unid_assignment_matches_the_whole_scope_walk()
    {
        // A paragraph whose runs lost their Unids (as after an edit that rebuilt them), in a body
        // whose other blocks are fully assigned.
        var xml =
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
            "<w:p><w:r><w:t>first</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>same</w:t></w:r><w:r><w:t>same</w:t></w:r>" +
            "<w:hyperlink><w:r><w:t>link</w:t></w:r></w:hyperlink><w:r><w:rPr><w:b/></w:rPr><w:t>same</w:t></w:r></w:p>" +
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "</w:body></w:document>";
        var whole = XDocument.Parse(xml);
        UnidHelper.AssignToAllElementsDeterministic(whole.Root!);
        foreach (int blockIndex in new[] { 1, 2 })
        {
            var expected = new XDocument(whole);
            var scoped = new XDocument(whole);
            foreach (var doc in new[] { expected, scoped })
            {
                var block = doc.Root!.Element(W.body)!.Elements().ElementAt(blockIndex);
                foreach (var d in block.Descendants()) d.Attribute(PtOpenXml.Unid)?.Remove();
            }
            UnidHelper.AssignToAllElementsDeterministic(expected.Root!);
            Assert.True(UnidHelper.AssignWithinBlock(scoped.Root!.Element(W.body)!.Elements().ElementAt(blockIndex)));
            Assert.True(XNode.DeepEquals(expected, scoped), $"block {blockIndex}");
        }
    }

    [Fact]
    public void Undo_retention_per_text_edit_no_longer_scales_with_the_document()
    {
        long PerStep(int paragraphs)
        {
            using var session = new DocxSession(Paragraphs(paragraphs), new DocxSessionSettings { EmitMarkdownPatch = false });
            var anchors = Anchors(session, "p");
            Assert.True(session.ReplaceText(anchors[0], "Warm-up.").Success);
            var afterOne = session.UndoMemoryBytes;
            for (int i = 1; i <= 10; i++)
                Assert.True(session.ReplaceText(anchors[i * 7], $"Edit {i}.").Success);
            return (session.UndoMemoryBytes - afterOne) / 10;
        }
        var small = PerStep(500);
        var large = PerStep(4000);
        _output.WriteLine($"per-step undo bytes: 500 paragraphs {small}, 4000 paragraphs {large}");
        // A step holds the edited block plus one chunk of block references (chunk boundaries vary
        // from run to run, so the figure does too), not a copy of every block: before, a step cost
        // about 1.3 KiB per paragraph, 650 KiB and 5.3 MiB here.
        Assert.True(small < 16 * 1024, $"{small} bytes per step at 500 paragraphs");
        Assert.True(large < 16 * 1024, $"{large} bytes per step at 4,000 paragraphs");
    }

    private static byte[] TinyPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    private static byte[] Paragraphs(int count)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            var body = new StringBuilder();
            for (var p = 0; p < count; p++)
                body.Append($"<w:p><w:r><w:t xml:space=\"preserve\">Paragraph {p} with some words in it.</w:t></w:r></w:p>");
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    $"<w:body>{body}<w:sectPr/></w:body></w:document>");
            doc.Save();
        }
        return stream.ToArray();
    }
}
