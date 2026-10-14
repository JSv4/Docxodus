#nullable enable
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// N-way NOTE-scope merge in the consolidate engine: reviewers' footnote edits against the shared base
/// note store compose per-block exactly like the body (disjoint edits land, consensus dedupes, contested
/// edits conflict per policy), reviewer-INSERTED notes land under fresh output ids with every body
/// reference resolvable, and reject restores the base notes — closing the last WmlComparer consolidate gap
/// (was: NotSupportedException for any N≥2 note edit).
/// </summary>
public class IrCompositeNoteTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // ------------------------------------------------------------------ fixtures

    /// <summary>A document whose body carries one paragraph per entry; an entry with a non-null Note gets a
    /// trailing footnote reference, ids assigned 1..k in body order (the Word convention).</summary>
    private static WmlDocument NoteDoc(params (string Para, string? Note)[] paras)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(new RunPropertiesDefault(
                new RunPropertiesBaseStyle(new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var notes = new System.Text.StringBuilder();
            notes.Append($"<w:footnotes xmlns:w=\"{W}\">")
                .Append("<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>")
                .Append("<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>");
            var body = new System.Text.StringBuilder();
            int nextId = 1;
            foreach (var (para, note) in paras)
            {
                body.Append($"<w:p><w:r><w:t xml:space=\"preserve\">{para}</w:t></w:r>");
                if (note != null)
                {
                    body.Append($"<w:r><w:footnoteReference w:id=\"{nextId}\"/></w:r>");
                    notes.Append($"<w:footnote w:id=\"{nextId}\"><w:p><w:r><w:t xml:space=\"preserve\">{note}</w:t></w:r></w:p></w:footnote>");
                    nextId++;
                }
                body.Append("</w:p>");
            }
            notes.Append("</w:footnotes>");

            var fn = main.AddNewPart<FootnotesPart>();
            WritePartXml(fn, notes.ToString());
            WritePartXml(main,
                $"<w:document xmlns:w=\"{W}\"><w:body>{body}" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("notes.docx", ms.ToArray());
    }

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
        writer.Write(xml);
    }

    private static WmlDocument Consolidate(
        WmlDocument baseDoc, ConflictResolution policy, params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.Consolidate(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    private static IReadOnlyList<DocxDiffConflict> Conflicts(
        WmlDocument baseDoc, ConflictResolution policy, params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.GetConflicts(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    // ------------------------------------------------------------------ oracles

    /// <summary>The footnote texts referenced from the body, in body-reference order; "(unresolved)" for a
    /// dangling reference — so both resolution AND content are checked in one shape.</summary>
    private static List<string> ReferencedNoteTexts(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        XNamespace ns = W;
        var defText = new Dictionary<string, string>();
        var fnRoot = main.FootnotesPart?.GetXDocument().Root;
        if (fnRoot != null)
            foreach (var f in fnRoot.Elements(ns + "footnote"))
                if ((string?)f.Attribute(ns + "id") is { } id)
                    defText[id] = string.Concat(f.Descendants(ns + "t").Select(t => t.Value));
        var result = new List<string>();
        var body = main.GetXDocument().Root?.Element(ns + "body");
        if (body != null)
            foreach (var r in body.Descendants(ns + "footnoteReference"))
            {
                var id = (string?)r.Attribute(ns + "id");
                result.Add(id != null && defText.TryGetValue(id, out var t) ? t : "(unresolved)");
            }
        return result;
    }

    private static string BodyText(WmlDocument doc) => Docs.PlainText(doc);

    // ------------------------------------------------------------------ 1. disjoint note edits compose

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Disjoint_note_edits_compose_both_land(ConflictResolution policy)
    {
        var b = NoteDoc(("First paragraph.", "note one text"), ("Second paragraph.", "note two text"));
        var alice = NoteDoc(("First paragraph.", "note one ALICE"), ("Second paragraph.", "note two text"));
        var bob = NoteDoc(("First paragraph.", "note one text"), ("Second paragraph.", "note two BOB"));

        Assert.Empty(Conflicts(b, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(b, policy, ("Alice", alice), ("Bob", bob));
        var accepted = RevisionAccepter.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        Assert.Equal(new List<string> { "note one ALICE", "note two BOB" }, ReferencedNoteTexts(accepted));
        Assert.Equal(new List<string> { "note one text", "note two text" }, ReferencedNoteTexts(rejected));
        Assert.Equal(BodyText(b), BodyText(rejected));
    }

    // ------------------------------------------------------------------ 2. same-note conflicting edits

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Same_note_conflicting_edits_resolve_per_policy(ConflictResolution policy)
    {
        var b = NoteDoc(("The paragraph.", "the base word here"));
        var alice = NoteDoc(("The paragraph.", "the ALICE word here"));
        var bob = NoteDoc(("The paragraph.", "the BOB word here"));

        var conflicts = Conflicts(b, policy, ("Alice", alice), ("Bob", bob));
        Assert.NotEmpty(conflicts);
        var authors = conflicts.SelectMany(c => c.Competitors.Select(x => x.Author)).Distinct().ToList();
        Assert.Contains("Alice", authors);
        Assert.Contains("Bob", authors);

        var merged = Consolidate(b, policy, ("Alice", alice), ("Bob", bob));
        var acceptedNote = ReferencedNoteTexts(RevisionAccepter.AcceptRevisions(merged)).Single();
        switch (policy)
        {
            case ConflictResolution.BaseWins:
                Assert.DoesNotContain("ALICE", acceptedNote);
                Assert.DoesNotContain("BOB", acceptedNote);
                break;
            case ConflictResolution.FirstReviewerWins:
                Assert.Contains("ALICE", acceptedNote);
                Assert.DoesNotContain("BOB", acceptedNote);
                break;
            case ConflictResolution.StackAll:
                Assert.Contains("ALICE", acceptedNote);
                Assert.Contains("BOB", acceptedNote);
                break;
        }

        Assert.Equal(new List<string> { "the base word here" },
            ReferencedNoteTexts(RevisionProcessor.RejectRevisions(merged)));
    }

    // ------------------------------------------------------------------ 3. same-note disjoint words compose

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Same_note_disjoint_word_edits_compose(ConflictResolution policy)
    {
        var b = NoteDoc(("The paragraph.", "alpha beta gamma delta"));
        var alice = NoteDoc(("The paragraph.", "ALICE beta gamma delta"));
        var bob = NoteDoc(("The paragraph.", "alpha beta gamma BOB"));

        Assert.Empty(Conflicts(b, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(b, policy, ("Alice", alice), ("Bob", bob));
        var acceptedNote = ReferencedNoteTexts(RevisionAccepter.AcceptRevisions(merged)).Single();
        Assert.Contains("ALICE", acceptedNote);
        Assert.Contains("BOB", acceptedNote);
        Assert.Contains("beta gamma", acceptedNote);

        Assert.Equal(new List<string> { "alpha beta gamma delta" },
            ReferencedNoteTexts(RevisionProcessor.RejectRevisions(merged)));
    }

    // ------------------------------------------------------------------ 4. note delete vs edit conflicts

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Note_delete_vs_edit_records_conflict_and_rejects_to_base(ConflictResolution policy)
    {
        var b = NoteDoc(("Kept paragraph.", "the disputed note"), ("Tail paragraph.", null));
        // Alice deletes the note (reference AND definition).
        var alice = NoteDoc(("Kept paragraph.", null), ("Tail paragraph.", null));
        // Bob edits the note's text.
        var bob = NoteDoc(("Kept paragraph.", "the disputed note EDITED"), ("Tail paragraph.", null));

        Assert.NotEmpty(Conflicts(b, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(b, policy, ("Alice", alice), ("Bob", bob));
        var rejected = RevisionProcessor.RejectRevisions(merged);
        Assert.Equal(BodyText(b), BodyText(rejected));
        Assert.Equal(new List<string> { "the disputed note" }, ReferencedNoteTexts(rejected));
    }

    // ------------------------------------------------------------------ 5. reviewer-inserted note

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Reviewer_inserted_note_lands_with_resolvable_refs(ConflictResolution policy)
    {
        var b = NoteDoc(("First paragraph.", "existing note"), ("Second paragraph.", null));
        // Alice adds a NEW footnote on the second paragraph (her ids: existing=1, new=2).
        var alice = NoteDoc(("First paragraph.", "existing note"), ("Second paragraph.", "brand new alice note"));
        // Bob edits body text elsewhere.
        var bob = NoteDoc(("First paragraph EDITED.", "existing note"), ("Second paragraph.", null));

        Assert.Empty(Conflicts(b, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(b, policy, ("Alice", alice), ("Bob", bob));
        var accepted = RevisionAccepter.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        // Accept: both notes referenced and resolvable, Bob's body edit landed.
        Assert.Equal(new List<string> { "existing note", "brand new alice note" }, ReferencedNoteTexts(accepted));
        Assert.Contains("First paragraph EDITED.", BodyText(accepted));

        // Reject: only the base note remains referenced; body ≡ base.
        Assert.Equal(new List<string> { "existing note" }, ReferencedNoteTexts(rejected));
        Assert.Equal(BodyText(b), BodyText(rejected));
    }

    /// <summary>
    /// The id-shift case: Alice inserts a note BEFORE the existing one, so her copy of the EXISTING note
    /// carries a SHIFTED id (2 instead of 1) — and she also edits the existing note's text, so her note diff
    /// is a matched modify under shifted ids. Every reference must still resolve to the right text on both
    /// accept and reject (the reviewer-sourced reference rewrite + renumber pass under test).
    /// </summary>
    [Fact]
    public void Inserted_note_shifting_ids_keeps_all_references_resolvable()
    {
        var b = NoteDoc(("First paragraph.", null), ("Second paragraph.", "the original note"));
        var alice = NoteDoc(("First paragraph.", "alice new note"), ("Second paragraph.", "the original note EDITED"));

        var merged = Consolidate(b, ConflictResolution.BaseWins, ("Alice", alice));
        var accepted = RevisionAccepter.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        Assert.Equal(new List<string> { "alice new note", "the original note EDITED" }, ReferencedNoteTexts(accepted));
        Assert.Equal(new List<string> { "the original note" }, ReferencedNoteTexts(rejected));
        Assert.Equal(BodyText(b), BodyText(rejected));
    }

    // ------------------------------------------------------------------ 6. two reviewers insert notes

    [Fact]
    public void Two_reviewers_inserting_notes_both_land_with_distinct_ids()
    {
        var b = NoteDoc(("First paragraph.", null), ("Second paragraph.", null), ("Anchor.", "base note"));
        var alice = NoteDoc(("First paragraph.", "alice note"), ("Second paragraph.", null), ("Anchor.", "base note"));
        var bob = NoteDoc(("First paragraph.", null), ("Second paragraph.", "bob note"), ("Anchor.", "base note"));

        Assert.Empty(Conflicts(b, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(b, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));
        var accepted = RevisionAccepter.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        Assert.Equal(new List<string> { "alice note", "bob note", "base note" }, ReferencedNoteTexts(accepted));
        Assert.Equal(new List<string> { "base note" }, ReferencedNoteTexts(rejected));
        Assert.Equal(BodyText(b), BodyText(rejected));
    }

    // ------------------------------------------------------------------ 7. consolidated revisions cover notes

    [Fact]
    public void Consolidated_revisions_surface_note_edits_with_attribution()
    {
        var b = NoteDoc(("The paragraph.", "note base text"));
        var alice = NoteDoc(("The paragraph.", "note ALICE text"));
        var bob = NoteDoc(("The paragraph EDITED.", "note base text"));

        var revs = DocxDiff.GetConsolidatedRevisions(b, new[]
        {
            new DocxDiffReviewer { Document = alice, Author = "Alice" },
            new DocxDiffReviewer { Document = bob, Author = "Bob" },
        });

        // Alice's note edit is a visible, attributed revision; Bob's body edit likewise.
        Assert.Contains(revs, r => r.Author == "Alice" && (r.Text.Contains("ALICE") || r.Text.Contains("base")));
        Assert.Contains(revs, r => r.Author == "Bob" && r.Text.Contains("EDITED"));
    }
}
