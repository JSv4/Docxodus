#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Tests.Ir;
using Xunit;
using WordType = DocumentFormat.OpenXml.WordprocessingDocumentType;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Tests for <see cref="DocxDiffSettings.PreserveInputRevisions"/> — the Word-parity flag that carries the
/// inputs' PRE-EXISTING tracked revisions (original author/date markup) through into the compare output,
/// while the text diff is still computed over the accepted view. Word's Compare behaves exactly this way:
/// an input's own <c>w:ins</c>/<c>w:del</c> markup rides through verbatim alongside the fresh compare
/// revisions.
///
/// <para><b>V1 semantics pinned here.</b> EQUAL blocks and whole-block INSERTS preserve the right input's
/// foreign markup (foreign wrappers are never re-wrapped, so no nested <c>w:ins</c>-in-<c>w:ins</c>);
/// MODIFIED/format-only/split/merge/move paths render over the accepted view (foreign markup there is
/// flattened — scoped out in v1). The round-trip contract under the flag is one-sided, as in Word:
/// <c>accept(output)</c> content-equals <c>accept(right)</c>, but <c>reject(output)</c> does NOT equal
/// <c>left</c> where foreign markup exists (rejecting a foreign deletion restores its text).</para>
/// </summary>
public class DocxDiffPreserveInputRevisionsTests
{
    private const string Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = Wns;

    // ----------------------------------------------------------------- fixture builders

    private static string R(string text) =>
        $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static string Ins(string author, string text, int id = 900) =>
        $"<w:ins w:id=\"{id}\" w:author=\"{author}\" w:date=\"2020-01-01T00:00:00Z\">" +
        $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:ins>";

    private static string InsWithPrefix(string prefix, string author, string text, int id = 900) =>
        $"<{prefix}:ins xmlns:{prefix}=\"{Wns}\" {prefix}:id=\"{id}\" {prefix}:author=\"{author}\" {prefix}:date=\"2020-01-01T00:00:00Z\">" +
        $"<{prefix}:r><{prefix}:t xml:space=\"preserve\">{text}</{prefix}:t></{prefix}:r></{prefix}:ins>";

    private static string Del(string author, string text, int id = 901) =>
        $"<w:del w:id=\"{id}\" w:author=\"{author}\" w:date=\"2020-01-01T00:00:00Z\">" +
        $"<w:r><w:delText xml:space=\"preserve\">{text}</w:delText></w:r></w:del>";

    private static string Table(string text, bool bidiVisual = false) =>
        "<w:tbl><w:tblPr>" + (bidiVisual ? "<w:bidiVisual/>" : string.Empty) + "</w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
        $"<w:tr><w:tc><w:tcPr/><w:p>{R(text)}</w:p></w:tc></w:tr></w:tbl>";

    /// <summary>Build a minimal DOCX whose body is the given paragraphs (each an inner-XML string of w:p content).</summary>
    private static WmlDocument BodyDoc(params string[] paragraphInnerXml)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            using var s = main.GetStream(FileMode.Create, FileAccess.Write);
            using var w = new StreamWriter(s);
            w.Write(
                $"<w:document xmlns:w=\"{Wns}\"><w:body>" +
                string.Concat(paragraphInnerXml.Select(p => $"<w:p>{p}</w:p>")) +
                "</w:body></w:document>");
        }
        return new WmlDocument("preserve-fixture.docx", ms.ToArray());
    }

    /// <summary>Add a header story carrying arbitrary WordprocessingML. The preservation gate must inspect
    /// this relationship just as <see cref="RevisionProcessor.HasTrackedRevisions(WmlDocument)"/> does; a
    /// body-only scan would preserve the right side of an asymmetric dirty pair.</summary>
    private static WmlDocument WithHeader(WmlDocument source, string headerInnerXml)
    {
        // The package gains a part/relationship, so use an expandable stream rather than the fixed-capacity
        // MemoryStream(byte[]) constructor.
        using var ms = new MemoryStream();
        ms.Write(source.DocumentByteArray, 0, source.DocumentByteArray.Length);
        ms.Position = 0;
        using (var wDoc = WordprocessingDocument.Open(ms, true))
        {
            var header = wDoc.MainDocumentPart!.AddNewPart<HeaderPart>();
            using var stream = header.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream);
            writer.Write($"<w:hdr xmlns:w=\"{Wns}\">{headerInnerXml}</w:hdr>");
        }
        return new WmlDocument("preserve-header-fixture.docx", ms.ToArray());
    }

    // ----------------------------------------------------------------- projections

    private static XDocument MainXDoc(WmlDocument d)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        using var s = doc.MainDocumentPart!.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(s);
    }

    /// <summary>Visible body text: concatenated <c>w:t</c> (deleted text lives in <c>w:delText</c> and is excluded).</summary>
    private static string BodyText(WmlDocument d) =>
        string.Concat(MainXDoc(d).Descendants(W + "t").Select(t => t.Value));

    private static string AcceptedBodyText(WmlDocument d) =>
        BodyText(RevisionProcessor.AcceptRevisions(d));

    /// <summary>Every revision-wrapper element (ins/del) whose author attribute matches.</summary>
    private static XElement[] RevisionWrappersBy(WmlDocument d, string author) =>
        MainXDoc(d).Descendants()
            .Where(e => (e.Name == W + "ins" || e.Name == W + "del") &&
                        (string?)e.Attribute(W + "author") == author)
            .ToArray();

    /// <summary>Asserts the output never nests a revision wrapper inside a same-kind wrapper
    /// (<c>w:ins</c> in <c>w:ins</c> or <c>w:del</c> in <c>w:del</c>) — the schema/Word-cleanliness bar.</summary>
    private static void AssertNoSameKindNesting(WmlDocument d)
    {
        var xd = MainXDoc(d);
        foreach (var name in new[] { W + "ins", W + "del" })
        {
            var nested = xd.Descendants(name).Where(e => e.Ancestors(name).Any()).ToList();
            Assert.True(nested.Count == 0,
                $"output nests {nested.Count} <{name.LocalName}> element(s) inside another <{name.LocalName}>.");
        }
    }

    private static DocxDiffSettings Preserve() =>
        new() { PreserveInputRevisions = true, AuthorForRevisions = "TheDiff" };

    // ----------------------------------------------------------------- 1: equal block, foreign w:ins

    [Fact]
    public void Equal_paragraph_preserves_foreign_ins_verbatim()
    {
        // Right paragraph 1 carries a foreign insertion by "Reviewer B"; its ACCEPTED text equals left
        // paragraph 1, so the pair aligns as EqualBlock. Paragraph 2 differs so the diff has real work.
        var left = BodyDoc(R("Shared lead REVTEXT end"), R("Alpha"));
        var right = BodyDoc(R("Shared lead ") + Ins("Reviewer B", "REVTEXT") + R(" end"), R("Beta"));

        var result = DocxDiff.Compare(left, right, Preserve());

        var foreign = RevisionWrappersBy(result, "Reviewer B");
        Assert.Contains(foreign, e => e.Name == W + "ins" &&
            string.Concat(e.Descendants(W + "t").Select(t => t.Value)) == "REVTEXT");
        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
    }

    // ----------------------------------------------------------------- 2: equal block, foreign w:del

    [Fact]
    public void Equal_paragraph_preserves_foreign_del_and_accept_removes_it()
    {
        var left = BodyDoc(R("Keep tail"), R("Alpha"));
        var right = BodyDoc(R("Keep ") + Del("Reviewer B", "GONE") + R("tail"), R("Beta"));

        var result = DocxDiff.Compare(left, right, Preserve());

        var foreign = RevisionWrappersBy(result, "Reviewer B");
        Assert.Contains(foreign, e => e.Name == W + "del" &&
            string.Concat(e.Descendants(W + "delText").Select(t => t.Value)) == "GONE");

        // accept(output) removes the foreign deletion — content-equals accept(right).
        var accepted = AcceptedBodyText(result);
        Assert.DoesNotContain("GONE", accepted);
        Assert.Equal(AcceptedBodyText(right), accepted);

        // The documented Word-identical caveat: reject(output) RESTORES the foreign deletion's text,
        // so reject does NOT reproduce left. Pin the deviation so it is deliberate, not accidental.
        Assert.Contains("GONE", BodyText(RevisionProcessor.RejectRevisions(result)));
    }

    // ----------------------------------------------------------------- 3: modified block stays schema-clean

    [Fact]
    public void Modified_paragraph_with_foreign_ins_is_schema_clean_and_round_trips()
    {
        // Right paragraph 1 both differs from left ("quick" → "slow", a real compare edit) AND carries a
        // foreign insertion — the ModifyBlock path. V1 renders modified paragraphs over the ACCEPTED view,
        // so the foreign markup is flattened there; the bar is NO nested same-kind wrappers and the
        // accept-side round trip.
        var left = BodyDoc(R("The quick brown fox"), R("Alpha"));
        var right = BodyDoc(R("The slow brown fox") + Ins("Reviewer B", " jumps"), R("Alpha"));

        var result = DocxDiff.Compare(left, right, Preserve());

        AssertNoSameKindNesting(result);
        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
    }

    // ----------------------------------------------------------------- inserted block preserves markup

    [Fact]
    public void Inserted_block_preserves_foreign_markup_without_nesting()
    {
        // Right-only paragraph (whole-block INSERT) carrying foreign ins AND del. Word preserves both; our
        // wrapper marks only the plain runs as inserted-by-the-diff, leaving the foreign wrappers as-is.
        var left = BodyDoc(R("Common intro"));
        var right = BodyDoc(
            R("Common intro"),
            R("New para ") + Ins("Reviewer B", "added") + Del("Reviewer B", "cut"));

        var result = DocxDiff.Compare(left, right, Preserve());

        var foreign = RevisionWrappersBy(result, "Reviewer B");
        Assert.Contains(foreign, e => e.Name == W + "ins");
        Assert.Contains(foreign, e => e.Name == W + "del");
        AssertNoSameKindNesting(result);

        // The paragraph's PLAIN text is attributed to this diff (it is inserted relative to left).
        Assert.Contains(RevisionWrappersBy(result, "TheDiff"), e => e.Name == W + "ins" &&
            string.Concat(e.Descendants(W + "t").Select(t => t.Value)).Contains("New para "));

        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
        Assert.DoesNotContain("cut", AcceptedBodyText(result));
    }

    // ----------------------------------------------------------------- fully-deleted paragraph rides along

    [Fact]
    public void Fully_deleted_paragraph_is_preserved_and_vanishes_on_accept()
    {
        // A paragraph whose CONTENT and MARK are both revision-deleted vanishes on accept, so it aligns
        // with nothing — the document-level accept merges it into the NEXT paragraph. Word keeps it in the
        // compare output; the preservation walk maps the {deleted para, next para} group onto the next
        // block and emits both.
        const string date = "2020-01-01T00:00:00Z";
        string fullyDeleted =
            $"<w:pPr><w:rPr><w:del w:id=\"910\" w:author=\"Reviewer B\" w:date=\"{date}\"/></w:rPr></w:pPr>" +
            Del("Reviewer B", "Removed by B", 911);

        var left = BodyDoc(R("First"), R("Second"));
        var right = BodyDoc(R("First"), fullyDeleted, R("Second"));

        var result = DocxDiff.Compare(left, right, Preserve());

        // The deleted paragraph's content AND its deleted mark survive, attributed to Reviewer B.
        var xd = MainXDoc(result);
        Assert.Contains("Removed by B", string.Concat(xd.Descendants(W + "delText").Select(t => t.Value)));
        Assert.Contains(xd.Descendants(W + "pPr").Elements(W + "rPr").Elements(W + "del"),
            e => (string?)e.Attribute(W + "author") == "Reviewer B");

        // Preserved wrappers get FRESH ids from the render's counter (the originals' ids would collide
        // with this diff's own) — no duplicate w:id anywhere in the body.
        var ids = xd.Descendants()
            .Where(e => e.Name == W + "ins" || e.Name == W + "del")
            .Select(e => (string?)e.Attribute(W + "id"))
            .Where(v => v != null)
            .ToList();
        Assert.Equal(ids.Count, ids.Distinct().Count());

        // accept ≡ accept(right): the deleted paragraph vanishes again.
        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
        Assert.DoesNotContain("Removed by B", AcceptedBodyText(result));
    }

    [Fact]
    public void Preserved_duplicate_foreign_wrapper_ids_are_normalized_independently()
    {
        // Some real-world documents reuse a w:id across unrelated w:ins elements. Each output revision
        // wrapper needs its own fresh annotation id; treating the source id as a global key duplicates ids
        // in the result and causes OOXML validation/Word-cleanliness problems.
        var left = BodyDoc(R("Shared first second"), R("Alpha"));
        var right = BodyDoc(
            R("Shared ") + Ins("Reviewer B", "first", 900) + R(" ") + Ins("Reviewer B", "second", 900),
            R("Beta"));

        var result = DocxDiff.Compare(left, right, Preserve());
        var foreignIds = MainXDoc(result).Descendants(W + "ins")
            .Where(e => (string?)e.Attribute(W + "author") == "Reviewer B")
            .Select(e => (string?)e.Attribute(W + "id"))
            .Where(id => id != null)
            .ToList();

        Assert.Equal(2, foreignIds.Count);
        Assert.Equal(foreignIds.Count, foreignIds.Distinct().Count());
        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
    }

    [Fact]
    public void Preserved_range_endpoints_share_an_id_but_wrapper_does_not()
    {
        // A range start/end pair must stay linked, even across preserved-clone normalization. Its moveTo
        // wrapper is a separate revision annotation though, despite deliberately reusing id 900 in the
        // malformed source document.
        const string moveStart =
            "<w:moveToRangeStart w:id=\"900\" w:author=\"Reviewer B\" w:date=\"2020-01-01T00:00:00Z\" w:name=\"move900\"/>";
        const string move =
            "<w:moveTo w:id=\"900\" w:author=\"Reviewer B\" w:date=\"2020-01-01T00:00:00Z\">" +
            "<w:r><w:t>moved</w:t></w:r></w:moveTo>";
        const string moveEnd = "<w:moveToRangeEnd w:id=\"900\"/>";
        var left = BodyDoc(R("Shared moved end"), R("Alpha"));
        var right = BodyDoc(R("Shared ") + moveStart + move + moveEnd + R(" end"), R("Beta"));

        var result = DocxDiff.Compare(left, right, Preserve());
        var xd = MainXDoc(result);
        var start = xd.Descendants(W + "moveToRangeStart")
            .Single(e => (string?)e.Attribute(W + "author") == "Reviewer B");
        var end = xd.Descendants(W + "moveToRangeEnd").Single();
        var wrapper = xd.Descendants(W + "moveTo")
            .Single(e => (string?)e.Attribute(W + "author") == "Reviewer B");

        Assert.Equal((string?)start.Attribute(W + "id"), (string?)end.Attribute(W + "id"));
        Assert.NotEqual((string?)start.Attribute(W + "id"), (string?)wrapper.Attribute(W + "id"));
        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
    }

    // ----------------------------------------------------------------- note-scope preservation

    /// <summary>A document whose body paragraph optionally references footnote 1; the note body carries the
    /// given inner XML (so the right side's copy can carry foreign revision markup).</summary>
    private static WmlDocument NoteDoc(string bodyText, string? footnoteInnerXml)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            string bodyRef = string.Empty;
            if (footnoteInnerXml != null)
            {
                var fnPart = main.AddNewPart<FootnotesPart>();
                using (var fs = fnPart.GetStream(FileMode.Create, FileAccess.Write))
                using (var fw = new StreamWriter(fs))
                {
                    fw.Write(
                        $"<w:footnotes xmlns:w=\"{Wns}\">" +
                        "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                        "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                        $"<w:footnote w:id=\"1\"><w:p>{footnoteInnerXml}</w:p></w:footnote>" +
                        "</w:footnotes>");
                }
                bodyRef = "<w:r><w:footnoteReference w:id=\"1\"/></w:r>";
            }

            using var s = main.GetStream(FileMode.Create, FileAccess.Write);
            using var w = new StreamWriter(s);
            w.Write(
                $"<w:document xmlns:w=\"{Wns}\"><w:body>" +
                $"<w:p><w:r><w:t xml:space=\"preserve\">{bodyText}</w:t></w:r>{bodyRef}</w:p>" +
                "</w:body></w:document>");
        }
        return new WmlDocument("preserve-note-fixture.docx", ms.ToArray());
    }

    private static string FootnotesXml(WmlDocument d)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var part = doc.MainDocumentPart!.FootnotesPart;
        if (part == null) return string.Empty;
        using var s = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(s).ToString();
    }

    [Fact]
    public void Inserted_note_preserves_foreign_markup_in_the_note_body()
    {
        // The RIGHT side introduces a footnote whose body carries a foreign insertion AND deletion by
        // "Reviewer B" — Word keeps both in the compare output's footnotes part. The note renderer routes
        // through the same block dispatch as the body, so note-scope preservation rides the same hooks.
        var left = NoteDoc("Text", footnoteInnerXml: null);
        var right = NoteDoc("Text",
            R("Note lead ") + Ins("Reviewer B", "NOTEADD", 920) + Del("Reviewer B", "NOTECUT", 921));

        var result = DocxDiff.Compare(left, right, Preserve());

        var fn = FootnotesXml(result);
        Assert.Contains("NOTEADD", fn);
        Assert.Contains("NOTECUT", fn);
        Assert.Contains("Reviewer B", fn);

        // accept(output)'s footnote text equals accept(right)'s (foreign del gone, foreign ins kept).
        var acceptedFn = FootnotesXml(RevisionProcessor.AcceptRevisions(result));
        Assert.Contains("NOTEADD", acceptedFn);
        Assert.DoesNotContain("NOTECUT", acceptedFn);
    }

    // ----------------------------------------------------------------- 4: flag interactions / guards

    [Fact]
    public void Left_revision_with_an_alternate_wordprocessingml_prefix_disables_right_preservation()
    {
        // The LEFT's x:ins is exactly the same OOXML element as w:ins. It must disable preservation of
        // RIGHT-only markup, otherwise the render flattens the left revision but retains Reviewer B's,
        // producing the asymmetry this guard is intended to prevent.
        var left = BodyDoc(
            R("Left ") + InsWithPrefix("x", "Reviewer A", "revision"),
            R("Shared lead REVTEXT end"),
            R("Alpha"));
        var right = BodyDoc(
            R("Left revision"),
            R("Shared lead ") + Ins("Reviewer B", "REVTEXT") + R(" end"),
            R("Beta"));

        var result = DocxDiff.Compare(left, right, Preserve());

        Assert.Empty(RevisionWrappersBy(result, "Reviewer B"));
        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
    }

    [Fact]
    public void Left_property_revision_disables_right_preservation()
    {
        // pPrChange has no w:ins/w:del wrapper, but RevisionProcessor treats it as a tracked revision.
        // The current (accepted) centered pPr matches RIGHT exactly; the old left alignment lives only in
        // pPrChange, isolating this test to the gate rather than a format diff.
        const string centered = "<w:pPr><w:jc w:val=\"center\"/></w:pPr>";
        const string centeredWithChange =
            "<w:pPr><w:jc w:val=\"center\"/><w:pPrChange w:id=\"902\" w:author=\"Reviewer A\" w:date=\"2020-01-01T00:00:00Z\">" +
            "<w:pPr><w:jc w:val=\"left\"/></w:pPr></w:pPrChange></w:pPr>";
        var left = BodyDoc(
            centeredWithChange + R("Stable"),
            R("Shared lead REVTEXT end"),
            R("Alpha"));
        var right = BodyDoc(
            centered + R("Stable"),
            R("Shared lead ") + Ins("Reviewer B", "REVTEXT") + R(" end"),
            R("Beta"));

        var result = DocxDiff.Compare(left, right, Preserve());

        Assert.Empty(RevisionWrappersBy(result, "Reviewer B"));
        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
    }

    [Fact]
    public void Left_input_insertion_deleted_by_the_compare_keeps_its_deletion_provenance()
    {
        // Word projects a left-side pre-existing insertion onto the comparison's DELETE side rather than
        // flattening it and creating a new deletion under the comparer author. The paragraph mark matters:
        // accept must remove the paragraph entirely while reject restores the left accepted view.
        const string author = "Reviewer A";
        const string insertedMark =
            "<w:ins w:id=\"910\" w:author=\"Reviewer A\" w:date=\"2020-01-01T00:00:00Z\"/>";
        const string insertedRuns =
            "<w:ins w:id=\"911\" w:author=\"Reviewer A\" w:date=\"2020-01-01T00:00:00Z\">" +
            "<w:r><w:t>Removed input </w:t></w:r><w:r><w:t>insertion</w:t></w:r></w:ins>";
        var left = BodyDoc(
            R("Common"),
            $"<w:pPr><w:rPr>{insertedMark}</w:rPr></w:pPr>{insertedRuns}");
        var right = BodyDoc(R("Common"));

        var result = DocxDiff.Compare(left, right, Preserve());
        var xd = MainXDoc(result);
        var projected = xd.Descendants(W + "del")
            .Where(e => (string?)e.Attribute(W + "author") == author)
            .ToList();

        var contentDeletion = Assert.Single(projected, e => e.Descendants(W + "delText").Any());
        Assert.Equal("Removed input insertion", string.Concat(contentDeletion.Descendants(W + "delText").Select(t => t.Value)));
        Assert.Equal(2, contentDeletion.Elements(W + "r").Count());
        Assert.Contains(xd.Descendants(W + "pPr").Elements(W + "rPr").Elements(W + "del"),
            e => (string?)e.Attribute(W + "author") == author);
        Assert.DoesNotContain(RevisionWrappersBy(result, author), e => e.Name == W + "ins");
        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
        Assert.Equal(AcceptedBodyText(left), BodyText(RevisionProcessor.RejectRevisions(result)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Left_input_insertion_after_table_normalization_keeps_its_deletion_provenance(bool secondTableIsBidiVisual)
    {
        // RevisionProcessor normalizes adjacent same-bidi tables into one accepted table and discards direct
        // body leaf markers. The preservation index must resynchronize past both before a later input
        // insertion; the false case is the comment-heavy benchmark shape, true proves a bidi boundary stays
        // on the ordinary strict one-to-one path.
        const string author = "Reviewer A";
        const string insertion =
            "<w:p><w:pPr><w:rPr><w:ins w:id=\"930\" w:author=\"Reviewer A\" w:date=\"2020-01-01T00:00:00Z\"/>" +
            "</w:rPr></w:pPr><w:ins w:id=\"931\" w:author=\"Reviewer A\" w:date=\"2020-01-01T00:00:00Z\">" +
            "<w:r><w:t>Later input insertion</w:t></w:r></w:ins></w:p>";
        string tablesAndMarker =
            $"{Table("First table")}{Table("Second table", secondTableIsBidiVisual)}<w:commentRangeEnd w:id=\"42\"/>";
        var left = IrTestDocuments.FromBodyXml($"<w:p>{R("Shared")}</w:p>{tablesAndMarker}{insertion}");
        // Make RIGHT dirty too so both IR reads execute RevisionProcessor's table normalization. Its harmless
        // foreign insertion is intentionally flattened because LEFT is dirty; it isolates the map behavior.
        var right = IrTestDocuments.FromBodyXml(
            $"<w:p>{Ins("Reviewer B", "Shared", 940)}</w:p>{tablesAndMarker}");

        var result = DocxDiff.Compare(left, right, Preserve());
        var projected = RevisionWrappersBy(result, author);

        Assert.Contains(projected, e => e.Name == W + "del" &&
            string.Concat(e.Descendants(W + "delText").Select(t => t.Value)) == "Later input insertion");
        Assert.DoesNotContain(projected, e => e.Name == W + "ins");
        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
        Assert.Equal(AcceptedBodyText(left), BodyText(RevisionProcessor.RejectRevisions(result)));
    }

    [Fact]
    public void Left_mixed_revision_group_falls_back_to_the_accepted_view()
    {
        // A pre-existing deletion is invisible in LEFT's accepted working copy. It cannot safely be projected
        // through the narrow w:ins→w:del path, so the old hidden text must not leak into the comparison delete.
        var left = BodyDoc(R("Common"), R("Visible left text") + Del("Reviewer A", "Hidden old deletion", 920));
        var right = BodyDoc(R("Common"));

        var result = DocxDiff.Compare(left, right, Preserve());
        var xml = MainXDoc(result);
        var allText = string.Concat(xml.Descendants(W + "t").Concat(xml.Descendants(W + "delText")).Select(t => t.Value));

        Assert.DoesNotContain("Hidden old deletion", allText);
        Assert.Contains("Visible left text", allText);
        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
        Assert.Equal(AcceptedBodyText(left), BodyText(RevisionProcessor.RejectRevisions(result)));
    }

    [Fact]
    public void Left_header_revision_disables_right_preservation()
    {
        // RevisionProcessor scans headers as well as the main story. A body-only preservation gate would
        // flatten this left-side foreign insertion while retaining Reviewer B's right-side one, recreating
        // the asymmetry that PreserveInputRevisions intentionally avoids.
        var left = WithHeader(BodyDoc(
            R("Left revision"),
            R("Shared lead REVTEXT end"),
            R("Alpha")),
            $"<w:p>{Ins("Reviewer A", "header revision")}</w:p>");
        var right = BodyDoc(
            R("Left revision"),
            R("Shared lead ") + Ins("Reviewer B", "REVTEXT") + R(" end"),
            R("Beta"));

        var result = DocxDiff.Compare(left, right, Preserve());

        Assert.Empty(RevisionWrappersBy(result, "Reviewer B"));
        Assert.Equal(AcceptedBodyText(right), AcceptedBodyText(result));
    }

    [Fact]
    public void Clean_inputs_flag_on_is_byte_identical_to_flag_off()
    {
        // With no foreign markup anywhere, the flag must be a byte-level no-op (deterministic default dates
        // make the outputs directly comparable).
        var left = BodyDoc(R("One two three"), R("Alpha"));
        var right = BodyDoc(R("One two four"), R("Alpha"));

        var off = DocxDiff.Compare(left, right, new DocxDiffSettings { AuthorForRevisions = "TheDiff" });
        var on = DocxDiff.Compare(left, right, Preserve());

        Assert.Equal(off.DocumentByteArray, on.DocumentByteArray);
    }

    [Fact]
    public void Preserve_wins_when_both_flags_are_set()
    {
        var left = BodyDoc(R("Keep tail"), R("Alpha"));
        var right = BodyDoc(R("Keep ") + Del("Reviewer B", "GONE") + R("tail"), R("Beta"));

        var both = DocxDiff.Compare(left, right, new DocxDiffSettings
        {
            PreserveInputRevisions = true,
            PreAcceptInputRevisions = true,
            AuthorForRevisions = "TheDiff",
        });
        var preserveOnly = DocxDiff.Compare(left, right, Preserve());

        // Preserve wins: pre-accept is skipped, byte-identical to preserve-only.
        Assert.Equal(preserveOnly.DocumentByteArray, both.DocumentByteArray);
        Assert.NotEmpty(RevisionWrappersBy(both, "Reviewer B"));
    }

    [Fact]
    public void DocxCompare_docxdiff_branch_enables_preserve()
    {
        // The engine-selector mapping (CLI/bench path) opts into Word-parity preservation.
        Assert.True(DocxCompare.ToDocxDiffSettings(new WmlComparerSettings()).PreserveInputRevisions);
    }
}
