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
/// N-way ENDNOTE-scope merge in the consolidate engine — the endnote sibling of <see cref="IrCompositeNoteTests"/>
/// (which is footnote-only). The consolidate acceptance criterion names "footnote/ENDNOTE text", and both kinds
/// run the SAME <c>IrCompositeMerger.MergeNoteScopes</c> kind-loop (<c>{ Footnote, Endnote }</c>) + the SAME
/// <c>ApplyCompositeNoteDiffs</c> render dispatch, so this pins that the endnote branch composes across reviewers
/// exactly like the footnote branch: disjoint endnote edits land, same-endnote edits conflict per policy, and
/// reject restores the base endnotes under every policy.
/// </summary>
public class IrCompositeEndnoteTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // ------------------------------------------------------------------ fixture

    /// <summary>A document whose body carries one paragraph per entry; an entry with a non-null Note gets a
    /// trailing endnote reference, ids assigned 1..k in body order (the Word convention).</summary>
    private static WmlDocument EndnoteDoc(params (string Para, string? Note)[] paras)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(new RunPropertiesDefault(
                new RunPropertiesBaseStyle(new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var notes = new System.Text.StringBuilder();
            notes.Append($"<w:endnotes xmlns:w=\"{W}\">")
                .Append("<w:endnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>")
                .Append("<w:endnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>");
            var body = new System.Text.StringBuilder();
            int nextId = 1;
            foreach (var (para, note) in paras)
            {
                body.Append($"<w:p><w:r><w:t xml:space=\"preserve\">{para}</w:t></w:r>");
                if (note != null)
                {
                    body.Append($"<w:r><w:endnoteReference w:id=\"{nextId}\"/></w:r>");
                    notes.Append($"<w:endnote w:id=\"{nextId}\"><w:p><w:r><w:t xml:space=\"preserve\">{note}</w:t></w:r></w:p></w:endnote>");
                    nextId++;
                }
                body.Append("</w:p>");
            }
            notes.Append("</w:endnotes>");

            WritePartXml(main.AddNewPart<EndnotesPart>(), notes.ToString());
            WritePartXml(main,
                $"<w:document xmlns:w=\"{W}\"><w:body>{body}" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("endnotes.docx", ms.ToArray());
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

    // ------------------------------------------------------------------ oracle

    /// <summary>The endnote texts referenced from the body, in body-reference order; "(unresolved)" for a
    /// dangling reference — so both resolution AND content are checked in one shape.</summary>
    private static List<string> ReferencedEndnoteTexts(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        XNamespace ns = W;
        var defText = new Dictionary<string, string>();
        var enRoot = main.EndnotesPart?.GetXDocument().Root;
        if (enRoot != null)
            foreach (var e in enRoot.Elements(ns + "endnote"))
                if ((string?)e.Attribute(ns + "id") is { } id)
                    defText[id] = string.Concat(e.Descendants(ns + "t").Select(t => t.Value));
        var result = new List<string>();
        var body = main.GetXDocument().Root?.Element(ns + "body");
        if (body != null)
            foreach (var r in body.Descendants(ns + "endnoteReference"))
            {
                var id = (string?)r.Attribute(ns + "id");
                result.Add(id != null && defText.TryGetValue(id, out var t) ? t : "(unresolved)");
            }
        return result;
    }

    private static string BodyText(WmlDocument doc) => Docs.PlainText(doc);

    // ------------------------------------------------------------------ 1. disjoint endnote edits compose

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Disjoint_endnote_edits_compose_both_land(ConflictResolution policy)
    {
        var b = EndnoteDoc(("First paragraph.", "end one text"), ("Second paragraph.", "end two text"));
        var alice = EndnoteDoc(("First paragraph.", "end one ALICE"), ("Second paragraph.", "end two text"));
        var bob = EndnoteDoc(("First paragraph.", "end one text"), ("Second paragraph.", "end two BOB"));

        Assert.Empty(Conflicts(b, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(b, policy, ("Alice", alice), ("Bob", bob));
        var accepted = RevisionAccepter.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        Assert.Equal(new List<string> { "end one ALICE", "end two BOB" }, ReferencedEndnoteTexts(accepted));
        Assert.Equal(new List<string> { "end one text", "end two text" }, ReferencedEndnoteTexts(rejected));
        Assert.Equal(BodyText(b), BodyText(rejected));
    }

    // ------------------------------------------------------------------ 2. same-endnote conflicting edits

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Same_endnote_conflicting_edits_resolve_per_policy(ConflictResolution policy)
    {
        var b = EndnoteDoc(("The paragraph.", "the base word here"));
        var alice = EndnoteDoc(("The paragraph.", "the ALICE word here"));
        var bob = EndnoteDoc(("The paragraph.", "the BOB word here"));

        var conflicts = Conflicts(b, policy, ("Alice", alice), ("Bob", bob));
        Assert.NotEmpty(conflicts);
        var authors = conflicts.SelectMany(c => c.Competitors.Select(x => x.Author)).Distinct().ToList();
        Assert.Contains("Alice", authors);
        Assert.Contains("Bob", authors);

        var merged = Consolidate(b, policy, ("Alice", alice), ("Bob", bob));
        var acceptedNote = ReferencedEndnoteTexts(RevisionAccepter.AcceptRevisions(merged)).Single();
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
            ReferencedEndnoteTexts(RevisionProcessor.RejectRevisions(merged)));
    }

    // ------------------------------------------------------------------ 3. consolidated revisions surface endnote edits

    [Fact]
    public void Consolidated_revisions_surface_endnote_edits_with_attribution()
    {
        var b = EndnoteDoc(("The paragraph.", "note base text"));
        var alice = EndnoteDoc(("The paragraph.", "note ALICE text"));
        var bob = EndnoteDoc(("The paragraph EDITED.", "note base text"));

        var revs = DocxDiff.GetConsolidatedRevisions(b, new[]
        {
            new DocxDiffReviewer { Document = alice, Author = "Alice" },
            new DocxDiffReviewer { Document = bob, Author = "Bob" },
        });

        // Alice's endnote edit is a visible, attributed revision; Bob's body edit likewise.
        Assert.Contains(revs, r => r.Author == "Alice" && (r.Text.Contains("ALICE") || r.Text.Contains("base")));
        Assert.Contains(revs, r => r.Author == "Bob" && r.Text.Contains("EDITED"));
    }
}
