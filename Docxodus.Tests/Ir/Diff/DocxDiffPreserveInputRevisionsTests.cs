#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
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

    private static string Del(string author, string text, int id = 901) =>
        $"<w:del w:id=\"{id}\" w:author=\"{author}\" w:date=\"2020-01-01T00:00:00Z\">" +
        $"<w:r><w:delText xml:space=\"preserve\">{text}</w:delText></w:r></w:del>";

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
