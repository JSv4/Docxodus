#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Tests.Ir;
using DocxodusDiffParityFixtures;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Comment structural integrity for <see cref="DocxDiff.Compare"/>, the comment counterpart to
/// <c>DocxDiffBookmarkStructureTests</c>. For every shape in <see cref="DocxDiffCommentFixtures"/> the
/// Compare output must be:
/// <list type="number">
/// <item><b>Schema valid</b> — no NEW OpenXml errors vs the inputs (catches the duplicate-comment-id
/// <c>Sem_UniqueAttributeValue</c>).</item>
/// <item><b>Comment-structurally sound in the intermediate</b> — every <c>w:commentReference</c> resolves
/// to exactly one <c>w:comment</c>; every <c>w:commentRangeStart</c> id is unique and pairs 1:1 with a
/// <c>w:commentRangeEnd</c> of the same id; no anchored id lacks a definition.</item>
/// <item><b>Round-trip at the comment-structure level</b> — the resolved-anchor projection (author + text
/// + threaded parent text) of <c>accept</c> equals RIGHT's and of <c>reject</c> equals LEFT's. A structural
/// assertion keyed by comment CONTENT, never a text multiset (ids may be renumbered by dedup).</item>
/// <item><b>Fine per-word markup</b> — an edited commented paragraph is NOT exploded into a whole-block
/// del-copy + ins-copy: for a pure single-paragraph text edit the output body keeps the right's paragraph
/// count.</item>
/// </list>
/// </summary>
public class DocxDiffCommentStructureTests
{
    public static IEnumerable<object[]> DuplicateReferenceCases()
    {
        foreach (var wrapper in new[] { "sdt", "customXml" })
        foreach (bool duplicate in new[] { false, true })
        foreach (bool referenceOnly in new[] { false, true })
        foreach (bool consolidate in new[] { false, true })
            yield return new object[] { wrapper, duplicate, referenceOnly, consolidate };
    }

    [Theory]
    [MemberData(nameof(DuplicateReferenceCases))]
    public void ReplacedWrapper_DuplicateCommentReferenceRemainsOneComment(
        string wrapper, bool duplicate, bool referenceOnly, bool consolidate)
    {
        WmlDocument Document(string tag, bool extra)
        {
            const string reference = "<w:r><w:commentReference w:id=\"0\"/></w:r>";
            string content = "<w:p>" + (referenceOnly ? "" : "<w:commentRangeStart w:id=\"0\"/>") +
                "<w:r><w:t>Cedar clause</w:t></w:r>" +
                (referenceOnly ? "" : "<w:commentRangeEnd w:id=\"0\"/>") + reference +
                (extra ? reference : "") + "</w:p>";
            string body = wrapper == "sdt"
                ? $"<w:sdt><w:sdtPr><w:tag w:val=\"{tag}\"/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>"
                : $"<w:customXml w:element=\"{tag}\">{content}</w:customXml>";
            return IrTestDocuments.WithComment("Cedar", "C", "2026-01-01T00:00:00Z", "Check this clause.", body);
        }

        var original = Document("Maple", duplicate);
        var revised = Document("Birch", duplicate);
        var redline = consolidate
            ? DocxDiff.Consolidate(original, new[] { new DocxDiffReviewer { Document = revised, Author = "Reviewer" } })
            : DocxCompare.Compare(original, revised);
        foreach (var (resolved, expected) in new[] {
            (RevisionProcessor.AcceptRevisions(redline), Document("Birch", false)),
            (RevisionProcessor.RejectRevisions(redline), Document("Maple", false)) })
        {
            Assert.Equal(AnchorProjection(expected), AnchorProjection(resolved));
            Assert.Equal(RangeCoverage(expected), RangeCoverage(resolved));
            Assert.Equal(BodyText(expected), BodyText(resolved));
            Assert.Empty(SchemaErrors(resolved));
        }
        Assert.Empty(SchemaErrors(redline));
    }

    [Theory]
    [InlineData(false, false, "body")]
    [InlineData(false, true, "body")]
    [InlineData(true, false, "body")]
    [InlineData(true, true, "body")]
    [InlineData(true, false, "table")]
    [InlineData(true, true, "table")]
    [InlineData(true, false, "sdt")]
    [InlineData(true, true, "sdt")]
    [InlineData(true, false, "body-mid")]
    [InlineData(true, false, "table-mid")]
    [InlineData(true, false, "hyperlink-mid")]
    public void ReviewerCommentInUnchangedText_SurvivesOnlyAccept(bool consolidate, bool control, string container)
    {
        const string plain = "<w:p><w:r><w:t>Maple clause</w:t></w:r></w:p>";
        string marked = "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t>Maple clause</w:t></w:r>" +
            "<w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>";
        if (container.EndsWith("-mid", StringComparison.Ordinal))
            marked = marked.Replace("<w:commentRangeStart w:id=\"0\"/><w:r><w:t>Maple clause",
                "<w:r><w:t xml:space=\"preserve\">Maple </w:t></w:r><w:commentRangeStart w:id=\"0\"/><w:r><w:t>clause");
        string Wrap(string paragraph) => container.Split('-')[0] switch
        {
            "hyperlink" => paragraph.Replace("<w:p>", "<w:p><w:hyperlink w:anchor=\"Cedar\">").Replace("</w:p>", "</w:hyperlink></w:p>"),
            "table" => "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc>" + paragraph + "</w:tc></w:tr></w:tbl>",
            "sdt" => "<w:sdt><w:sdtPr><w:tag w:val=\"clause\"/></w:sdtPr><w:sdtContent>" + paragraph + "</w:sdtContent></w:sdt>",
            _ => paragraph,
        };
        var original = IrTestDocuments.FromBodyXml(Wrap(plain));
        var revised = control ? original : IrTestDocuments.WithComment(
            "Cedar", "C", "2026-01-01T00:00:00Z", "Check the maple clause.", Wrap(marked));
        var redline = consolidate
            ? DocxDiff.Consolidate(original, new[] { new DocxDiffReviewer { Document = revised, Author = "Cedar" } })
            : DocxDiff.Compare(original, revised);
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Equal(AnchorProjection(revised), AnchorProjection(accepted));
        Assert.Equal(RangeCoverage(revised), RangeCoverage(accepted));
        Assert.Equal(AnchorProjection(original), AnchorProjection(rejected));
        Assert.Equal(BodyText(original), BodyText(accepted));
        Assert.Empty(SchemaErrors(redline).Except(SchemaErrors(original).Concat(SchemaErrors(revised))));
        using var stream = new MemoryStream(redline.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        var markers = package.MainDocumentPart!.GetXDocument().Descendants(W.ins).ToList();
        Assert.Equal(control, markers.Count == 0);
        if (consolidate)
            Assert.All(markers, m => Assert.Equal("Cedar", (string?)m.Attribute(W.author)));
    }

    public static IEnumerable<object[]> AllScenarios() =>
        DocxDiffCommentFixtures.Names().Select(n => new object[] { n });

    public static IEnumerable<object[]> SingleParaEditScenarios() =>
        DocxDiffCommentFixtures.SingleParaEditShapes().Select(n => new object[] { n });

    [Theory]
    [MemberData(nameof(AllScenarios))]
    public void Comment_Compare_IsSchemaValid(string scenario)
    {
        var (left, right) = DocxDiffCommentFixtures.Build(scenario);
        var result = DocxDiff.Compare(left, right);

        var baseErrors = SchemaErrors(left).Concat(SchemaErrors(right)).ToHashSet();
        var newErrors = SchemaErrors(result).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0,
            $"[{scenario}] Compare introduced {newErrors.Count} new schema error(s): {string.Join(" | ", newErrors.Take(5))}");
    }

    [Theory]
    [MemberData(nameof(AllScenarios))]
    public void Comment_Compare_HasResolvedPairedReferences(string scenario)
    {
        var (left, right) = DocxDiffCommentFixtures.Build(scenario);
        var result = DocxDiff.Compare(left, right);
        var s = CommentStructure(result);

        // Every commentRangeStart id is unique (no two starts share an id).
        Assert.True(s.RangeStartIds.Count == s.RangeStartIds.Distinct().Count(),
            $"[{scenario}] duplicate commentRangeStart id(s): [{string.Join(",", s.RangeStartIds)}]");

        // 1:1 range pairing: the multiset of start ids equals the multiset of end ids.
        var onlyStart = s.RangeStartIds.Except(s.RangeEndIds).ToList();
        var onlyEnd = s.RangeEndIds.Except(s.RangeStartIds).ToList();
        Assert.True(onlyStart.Count == 0, $"[{scenario}] commentRangeStart id(s) without matching End: [{string.Join(",", onlyStart)}]");
        Assert.True(onlyEnd.Count == 0, $"[{scenario}] commentRangeEnd id(s) without matching Start: [{string.Join(",", onlyEnd)}]");

        // Every anchored id (reference / range-start / range-end) resolves to exactly one definition.
        Assert.True(s.UnresolvedIds.Count == 0,
            $"[{scenario}] comment id(s) with no/ambiguous definition: [{string.Join(",", s.UnresolvedIds)}]");
    }

    [Theory]
    [MemberData(nameof(AllScenarios))]
    public void Comment_RoundTrips_AtStructureLevel(string scenario)
    {
        var (left, right) = DocxDiffCommentFixtures.Build(scenario);
        var result = DocxDiff.Compare(left, right);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);

        // body text round-trip (the WmlComparer output contract)
        Assert.Equal(BodyText(right), BodyText(accepted));
        Assert.Equal(BodyText(left), BodyText(rejected));

        // comment-structure round-trip: accept ≡ right, reject ≡ left (anchored-comment projection by content)
        Assert.Equal(AnchorProjection(right), AnchorProjection(accepted));
        Assert.Equal(AnchorProjection(left), AnchorProjection(rejected));

        // anchor SPAN coverage round-trip: the text each comment physically brackets must match the side it
        // resolves to (a surviving range marker placed on the wrong side of an edit leaves the wrong content —
        // or NOTHING — anchored, which the resolution/pairing checks above cannot see).
        Assert.Equal(RangeCoverage(right), RangeCoverage(accepted));
        Assert.Equal(RangeCoverage(left), RangeCoverage(rejected));
    }

    [Theory]
    [MemberData(nameof(SingleParaEditScenarios))]
    public void Comment_EditedParagraph_RendersFineMarkup(string scenario)
    {
        var (left, right) = DocxDiffCommentFixtures.Build(scenario);
        var result = DocxDiff.Compare(left, right);

        // A whole-block bail explodes the edited commented paragraph into a del-copy + ins-copy (paragraph
        // count grows by one per such paragraph). Fine per-word markup keeps the body paragraph count equal
        // to the right document's.
        Assert.Equal(BodyParagraphCount(right), BodyParagraphCount(result));

        // And the edited commented paragraph carries BOTH a run-level w:ins and a run-level w:del (genuine
        // per-word redline) rather than a whole-paragraph mark deletion.
        Assert.True(HasRunLevelInsAndDel(result),
            $"[{scenario}] expected fine per-word w:ins + w:del markup in the commented paragraph");
    }

    [Theory]
    [MemberData(nameof(AllScenarios))]
    public void Comment_Consolidate_IsSchemaValidAndSound(string scenario)
    {
        // The composite/Consolidate path now reconciles comments too (comment markers are AlwaysKeep, so they
        // ride the composite token diff exactly as in two-way Compare). Treat the fixture's LEFT as the shared
        // base and RIGHT as one reviewer: the consolidated output must be schema-valid (no duplicate comment ids
        // / dangling references) with unique, 1:1-paired, fully-resolved comment markers.
        var (left, right) = DocxDiffCommentFixtures.Build(scenario);
        var result = DocxDiff.Consolidate(left,
            new[] { new DocxDiffReviewer { Document = right, Author = "Reviewer" } });

        var baseErrors = SchemaErrors(left).Concat(SchemaErrors(right)).ToHashSet();
        var newErrors = SchemaErrors(result).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0,
            $"[{scenario}] Consolidate introduced {newErrors.Count} new schema error(s): {string.Join(" | ", newErrors.Take(5))}");

        var s = CommentStructure(result);
        Assert.True(s.RangeStartIds.Count == s.RangeStartIds.Distinct().Count(),
            $"[{scenario}] duplicate commentRangeStart id(s): [{string.Join(",", s.RangeStartIds)}]");
        Assert.True(s.RangeStartIds.OrderBy(x => x).SequenceEqual(s.RangeEndIds.OrderBy(x => x)),
            $"[{scenario}] commentRange start/end not 1:1 paired");
        Assert.True(s.UnresolvedIds.Count == 0,
            $"[{scenario}] comment id(s) with no/ambiguous definition: [{string.Join(",", s.UnresolvedIds)}]");
    }

    [Fact]
    public void Comment_IdentityDiff_IsCleanNoOp()
    {
        foreach (var name in DocxDiffCommentFixtures.Names())
        {
            var (left, _) = DocxDiffCommentFixtures.Build(name);
            var result = DocxDiff.Compare(left, left);
            Assert.Empty(SchemaErrors(result).Where(e => !SchemaErrors(left).Contains(e)));
            Assert.Equal(AnchorProjection(left), AnchorProjection(RevisionProcessor.AcceptRevisions(result)));
            Assert.Equal(AnchorProjection(left), AnchorProjection(RevisionProcessor.RejectRevisions(result)));
        }
    }

    [Theory]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, false)]
    [InlineData(true, true, true)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    [InlineData(false, false, true)]
    [InlineData(true, true, false, "author")]
    [InlineData(true, true, true, "author")]
    [InlineData(true, true, false, "date")]
    [InlineData(true, true, true, "date")]
    [InlineData(true, false, false, "duplicate")]
    [InlineData(true, false, true, "duplicate")]
    [InlineData(false, true, false, "text", true)]
    [InlineData(false, false, false, "text", true)]
    [InlineData(false, true, true, "text", true)]
    [InlineData(false, false, true, "text", true)]
    [InlineData(true, true, false, "annotation-only")]
    [InlineData(true, true, true, "annotation-only")]
    [InlineData(false, true, false, "annotation-only", true)]
    [InlineData(false, true, true, "annotation-only", true)]
    [InlineData(false, true, false, "orphan", true)]
    [InlineData(false, true, true, "orphan", true)]
    public void Comment_IdentitySurvivesReusedIds(bool keepOld, bool collision, bool consolidate,
        string identityChange = "text", bool keepRemovedText = false)
    {
        WmlDocument Document(bool revised)
        {
            string Note(string id, bool added)
            {
                string author = added && identityChange == "author" ? "Willow" : "Cedar";
                string date = added && identityChange == "date" ? "2026-02-01" : "2026-01-01";
                string text = added && identityChange is "text" or "annotation-only" or "orphan" ? "New willow note" : "Old cedar note";
                return $"<w:comment w:id=\"{id}\" w:author=\"{author}\" w:date=\"{date}T00:00:00Z\">"
                    + $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>";
            }
            string oldId = revised && collision ? "1" : "0", newId = collision ? "0" : "2";
            var doc = Ir.IrTestDocuments.FromBodyXml("<w:p>" + CommentRun("AlderFirst", revised ? newId : null)
                + "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>"
                + CommentRun(revised && !keepOld && !keepRemovedText ? "WillowReplacement" : "BirchSecond", !revised || keepOld ? oldId : null)
                + "<w:r><w:t xml:space=\"preserve\">" + (revised && identityChange != "annotation-only" ? " tail revised" : " tail") + "</w:t></w:r></w:p>");
            return AddComments(doc, (!revised || keepOld || identityChange == "orphan" ? Note(oldId, false) : "")
                + (revised ? Note(newId, true) : ""));
        }
        var left = Document(false);
        var right = Document(true);
        var leftBytes = left.DocumentByteArray.ToArray();
        var rightBytes = right.DocumentByteArray.ToArray();
        var result = consolidate
            ? DocxCompare.Consolidate(left, new[] { new DocxDiffReviewer { Document = right, Author = "Reviewer" } })
            : DocxCompare.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        foreach (var doc in new[] { left, right, result, accepted, rejected }) Assert.Empty(SchemaErrors(doc));
        Assert.Empty(CommentStructure(result).UnresolvedIds);
        Assert.Equal(RangeCoverage(right, true), RangeCoverage(accepted, true));
        Assert.Equal(RangeCoverage(left, true), RangeCoverage(rejected, true));
        Assert.Equal(AnchorProjection(right), AnchorProjection(accepted));
        Assert.Equal(AnchorProjection(left), AnchorProjection(rejected));
        Assert.Equal(BodyText(right), BodyText(accepted));
        Assert.Equal(BodyText(left), BodyText(rejected));
        if (!consolidate)
        {
            var snapshot = DocxDiff.CreateSnapshot(right);
            var comparison = DocxDiff.CreateComparison(DocxDiff.CreateSnapshot(left), snapshot);
            var throughSnapshot = comparison.ToRedline();
            Assert.Same(throughSnapshot, comparison.ToRedline());
            Assert.Equal(RangeCoverage(right, true), RangeCoverage(RevisionProcessor.AcceptRevisions(throughSnapshot), true));
            Assert.Equal(RangeCoverage(left, true), RangeCoverage(RevisionProcessor.RejectRevisions(throughSnapshot), true));
            Assert.Equal(RangeCoverage(right, true), RangeCoverage(DocxDiff.CreateComparison(snapshot, snapshot).ToRedline(), true));
        }
        Assert.Equal(leftBytes, left.DocumentByteArray);
        Assert.Equal(rightBytes, right.DocumentByteArray);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Comment_IndependentReviewersReusingOneIdKeepBothDefinitions(bool identicalNotes)
    {
        WmlDocument Document(bool first, bool second)
        {
            string Paragraph(bool changed, string word) => "<w:p>" + CommentRun(word + (changed ? " revised" : ""), changed ? "0" : null) + "</w:p>";
            var doc = Ir.IrTestDocuments.FromBodyXml(Paragraph(first, "Alder") + Paragraph(second, "Birch"));
            if (!first && !second) return doc;
            return AddComments(doc, "<w:comment w:id=\"0\" w:author=\"Cedar\" w:date=\"2026-01-01T00:00:00Z\">"
                + "<w:p><w:r><w:t>" + (identicalNotes ? "Check clause" : first ? "Check alder" : "Check birch")
                + "</w:t></w:r></w:p></w:comment>");
        }
        var original = Document(false, false);
        var first = Document(true, false);
        var second = Document(false, true);
        var result = DocxCompare.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Author = "First reviewer", Document = first },
            new DocxDiffReviewer { Author = "Second reviewer", Document = second },
        });
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        foreach (var doc in new[] { original, first, second, result, accepted, rejected }) Assert.Empty(SchemaErrors(doc));
        Assert.Empty(CommentStructure(result).UnresolvedIds);
        Assert.Equal(2, CommentStructure(accepted).RangeStartIds.Count);
        Assert.Equal(2, CommentStructure(accepted).RangeStartIds.Distinct().Count());
        foreach (var source in new[] { first, second })
            Assert.Contains(RangeCoverage(source, true)[5..^1], RangeCoverage(accepted, true));
        Assert.Equal(RangeCoverage(original, true), RangeCoverage(rejected, true));
    }

    private static string CommentRun(string text, string? id) => id is null ? $"<w:r><w:t>{text}</w:t></w:r>"
        : $"<w:commentRangeStart w:id=\"{id}\"/><w:r><w:t>{text}</w:t></w:r>"
            + $"<w:commentRangeEnd w:id=\"{id}\"/><w:r><w:commentReference w:id=\"{id}\"/></w:r>";

    private static WmlDocument AddComments(WmlDocument document, string comments)
    {
        using var stream = new MemoryStream();
        stream.Write(document.DocumentByteArray);
        using (var package = WordprocessingDocument.Open(stream, true))
        {
            var part = package.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>();
            using var writer = new StreamWriter(part.GetStream(FileMode.Create));
            writer.Write("<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + comments + "</w:comments>");
        }
        return new WmlDocument("comment-identity.docx", stream.ToArray());
    }

    // ---- comment structure observers ------------------------------------------------------------

    private sealed record CmtStructure(
        List<string> RangeStartIds, List<string> RangeEndIds, List<string> UnresolvedIds);

    /// <summary>The intermediate-document comment structure: range-start/end id lists plus the set of
    /// anchored ids (reference OR range marker) that do NOT resolve to exactly one comment definition.</summary>
    private static CmtStructure CommentStructure(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var body = main.Document?.Body;
        List<string> Ids(string n) => body?.Descendants().Where(e => e.LocalName == n)
            .Select(e => AttrValue(e, "id")).ToList() ?? new();

        var startIds = Ids("commentRangeStart");
        var endIds = Ids("commentRangeEnd");
        var refIds = Ids("commentReference");

        // definition id → count (to catch missing AND ambiguous-duplicate definitions)
        var defCounts = new Dictionary<string, int>();
        var croot = main.WordprocessingCommentsPart?.GetXDocument().Root;
        if (croot != null)
            foreach (var c in croot.Elements().Where(e => e.Name.LocalName == "comment"))
            {
                var id = AttrLocal(c, "id");
                defCounts[id] = defCounts.TryGetValue(id, out var n) ? n + 1 : 1;
            }

        var anchored = startIds.Concat(endIds).Concat(refIds).ToHashSet();
        var unresolved = anchored.Where(id => !defCounts.TryGetValue(id, out var n) || n != 1).ToList();
        return new CmtStructure(startIds, endIds, unresolved);
    }

    /// <summary>
    /// A renderer-independent snapshot of a document's ANCHORED comments: for every comment id referenced
    /// (or ranged) in the body that resolves to a definition, the (author, normalized-text, parent-text)
    /// triple. Keyed by CONTENT, so dedup-renumbered ids still compare equal. Threaded reply links are
    /// resolved through <c>commentsExtended</c> (paraId → paraIdParent → parent comment text).
    /// </summary>
    private static string AnchorProjection(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var body = main.Document?.Body;
        if (body is null) return "anchors=[]";

        // id → (author, text, paraId)
        var defs = new Dictionary<string, (string Author, string Text, string ParaId)>();
        var croot = main.WordprocessingCommentsPart?.GetXDocument().Root;
        if (croot != null)
            foreach (var c in croot.Elements().Where(e => e.Name.LocalName == "comment"))
            {
                var id = AttrLocal(c, "id");
                var author = AttrLocal(c, "author");
                var text = string.Concat(c.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value));
                var firstP = c.Descendants().FirstOrDefault(e => e.Name.LocalName == "p");
                var paraId = firstP != null ? AttrLocal(firstP, "paraId") : "";
                defs[id] = (author, text, paraId);
            }

        // paraId → paraIdParent (commentsExtended threading)
        var parentOfPara = new Dictionary<string, string>();
        var exroot = main.WordprocessingCommentsExPart?.GetXDocument().Root;
        if (exroot != null)
            foreach (var ex in exroot.Elements().Where(e => e.Name.LocalName == "commentEx"))
            {
                var pid = AttrLocal(ex, "paraId");
                var parent = AttrLocal(ex, "paraIdParent");
                if (pid.Length > 0 && parent.Length > 0) parentOfPara[pid] = parent;
            }
        string ParentText(string paraId)
        {
            if (!parentOfPara.TryGetValue(paraId, out var parentPara)) return "";
            var hit = defs.Values.FirstOrDefault(d => d.ParaId == parentPara);
            return hit.Text ?? "";
        }

        // anchored ids = those with a commentReference in the body (the live anchor)
        var refIds = body.Descendants().Where(e => e.LocalName == "commentReference")
            .Select(e => AttrValue(e, "id")).ToList();

        var anchors = refIds.Where(id => defs.ContainsKey(id))
            .Select(id => $"{defs[id].Author}|{defs[id].Text}|parent={ParentText(defs[id].ParaId)}")
            .OrderBy(s => s, StringComparer.Ordinal).ToList();
        return $"anchors=[{string.Join(" ;; ", anchors)}]";
    }

    /// <summary>The actual visible text each comment physically brackets (between its commentRangeStart and
    /// commentRangeEnd in document order), keyed by comment CONTENT so dedup-renumbered ids still compare. Run
    /// on a MATERIALIZED accept/reject document (revisions resolved) it is the span-coverage counterpart to
    /// <see cref="AnchorProjection"/>: it catches a surviving range marker placed on the wrong side of an
    /// edited word, where accept/reject bracket different text than the side they must equal.</summary>
    private static string RangeCoverage(WmlDocument doc, bool includeDate = false)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var body = main.Document?.Body;
        if (body is null) return "cov=[]";

        var content = new Dictionary<string, string>();
        var croot = main.WordprocessingCommentsPart?.GetXDocument().Root;
        if (croot != null)
            foreach (var c in croot.Elements().Where(e => e.Name.LocalName == "comment"))
                content[AttrLocal(c, "id")] = AttrLocal(c, "author") + "|" + (includeDate ? AttrLocal(c, "date") + "|" : "") +
                    string.Concat(c.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value));

        var open = new Dictionary<string, System.Text.StringBuilder>();
        var done = new List<string>();
        foreach (var e in body.Descendants())
        {
            if (e.LocalName == "commentRangeStart")
                open[AttrValue(e, "id")] = new System.Text.StringBuilder();
            else if (e.LocalName == "commentRangeEnd")
            {
                var id = AttrValue(e, "id");
                if (open.TryGetValue(id, out var sb))
                {
                    // Trim the bracketed text: a deletion whose internal whitespace token-aligns with adjacent
                    // context can migrate a boundary space into/out of a wholly-deleted comment's span (a
                    // token-diff quality artifact, body text still round-trips). We assert CONTENT coverage —
                    // a real wrong-side marker (e.g. "" vs "contested") still diverges after trimming.
                    done.Add($"{(content.TryGetValue(id, out var k) ? k : id)}=>[{sb.ToString().Trim()}]");
                    open.Remove(id);
                }
            }
            else if (e.LocalName == "t")
                foreach (var sb in open.Values) sb.Append(e.InnerText);
        }
        return "cov=[" + string.Join(" ;; ", done.OrderBy(s => s, StringComparer.Ordinal)) + "]";
    }

    private static int BodyParagraphCount(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        // top-level paragraphs only (exclude comment-definition paragraphs, which live in another part)
        return body?.Elements().Count(e => e.LocalName == "p") ?? 0;
    }

    private static bool HasRunLevelInsAndDel(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        if (body is null) return false;
        // A run-level w:ins/w:del (inside a paragraph) — NOT the paragraph-mark marker in w:pPr/w:rPr.
        bool RunLevel(string local) => body.Descendants().Any(e => e.LocalName == local &&
            e.Parent?.LocalName != "rPr" && e.Ancestors().Any(a => a.LocalName == "p"));
        return RunLevel("ins") && RunLevel("del");
    }

    private static string AttrLocal(XElement e, string localName) =>
        (string?)e.Attributes().FirstOrDefault(a => a.Name.LocalName == localName) ?? "";

    private static string AttrValue(OpenXmlElement e, string localName) =>
        e.GetAttributes().FirstOrDefault(a => a.LocalName == localName).Value ?? "";

    private static string BodyText(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        return body is null ? "" : string.Concat(body.Descendants<Text>().Select(t => t.Text));
    }

    // Key on Id + part + value-normalized description (numeric ids → '#') so a legitimate comment-id
    // renumber is not miscounted as a NEW error (mirrors DocxDiffBookmarkRealDocTests.SchemaErrors).
    private static HashSet<string> SchemaErrors(WmlDocument doc) =>
        PackageValidation.ErrorKeys(doc, FileFormatVersions.Office2019,
            PackageValidation.Keys.IdAtPartNumberNormalizedDescription);
}
