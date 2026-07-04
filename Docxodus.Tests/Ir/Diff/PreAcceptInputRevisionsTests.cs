#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;
using static Docxodus.Tests.Ir.Diff.RevisionsInInputFixtures;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Tests for <see cref="DocxDiffSettings.PreAcceptInputRevisions"/> — the opt-in accept-all pre-flatten that
/// makes "diff the accepted view of both inputs" a first-class setting. The flag, when set, runs each input
/// through <see cref="RevisionProcessor.AcceptRevisions(WmlDocument)"/> before diffing, so the comparison —
/// and the OUTPUT PACKAGE the markup renderer clones from those inputs — is revision-free on BOTH sides.
///
/// <para>These are the FLAG-ON oracles (b)-(f). The DEFAULT (flag-off) behavior is pinned separately in
/// <see cref="RevisionsInInputDefaultTests"/>: the engine already diffs the accepted view at the
/// edit-script/body level, but carries pre-existing revision markup in non-body parts (headers/footers,
/// unchanged notes) straight through — the leak this flag eliminates.</para>
/// </summary>
public class PreAcceptInputRevisionsTests
{
    private static int SchemaErrorCount(WmlDocument doc)
    {
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        ms.Position = 0;
        using var wd = WordprocessingDocument.Open(ms, false);
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        return validator.Validate(wd).Count(e =>
            e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema &&
            !OxPt.WcTests.ExpectedErrors.Contains(e.Description));
    }

    // ---- oracle (b): the flag IS the wrapper ----------------------------------------------------

    [Fact]
    public void Flag_on_is_byte_identical_to_accept_all_then_compare_wrapper()
    {
        var left = MultiScopeRevisionDoc(
            "Body alpha", "Alice", "alpha-ins", "alpha-del", "PriorBob", "hdr-prior", "fn-prior");
        var right = MultiScopeRevisionDoc(
            "Body gamma", "Alice", "gamma-ins", "gamma-del", "PriorBob", "hdr-prior", "fn-prior");

        var flagOn = DocxDiff.Compare(left, right,
            new DocxDiffSettings { PreAcceptInputRevisions = true, AuthorForRevisions = "NewDiff" });
        var wrapper = DocxDiff.Compare(AcceptAll(left), AcceptAll(right),
            new DocxDiffSettings { AuthorForRevisions = "NewDiff" });

        // The flag is, by construction, exactly "accept-all both inputs, then compare" — so the produced
        // tracked-changes packages are byte-for-byte identical.
        Assert.Equal(wrapper.DocumentByteArray, flagOn.DocumentByteArray);
    }

    // ---- oracle (c): no pre-existing input revision markup leaks into the output ----------------

    [Fact]
    public void Flag_on_eliminates_stale_input_revision_authorship_in_every_scope()
    {
        var left = MultiScopeRevisionDoc(
            "Body alpha", "PriorAlice", "ins-a", "del-a", "PriorBob", "hdr-prior", "fn-prior");
        var right = MultiScopeRevisionDoc(
            "Body gamma", "PriorAlice", "ins-g", "del-g", "PriorBob", "hdr-prior", "fn-prior");

        // DEFAULT (flag off): the prior reviewers' markup in the carried-over parts LEAKS straight through —
        // the contrast that proves the flag does real work (pinned formally in the characterization suite).
        var leaked = DocxDiff.Compare(left, right, new DocxDiffSettings { AuthorForRevisions = "NewDiff" });
        Assert.Contains("PriorBob", RevisionAuthorsAllScopes(leaked));

        // FLAG ON: every revision in the result is attributable to THIS diff's author — none is a stale input
        // revision passed through (no "PriorAlice"/"PriorBob" anywhere, in any scope).
        var clean = DocxDiff.Compare(left, right,
            new DocxDiffSettings { PreAcceptInputRevisions = true, AuthorForRevisions = "NewDiff" });
        Assert.Equal(new HashSet<string> { "NewDiff" }, RevisionAuthorsAllScopes(clean));
    }

    // ---- oracle (d): round-trip against the ACCEPTED view of each side, in every scope ----------

    [Fact]
    public void Flag_on_round_trips_reject_to_accepted_left_and_accept_to_accepted_right()
    {
        var left = MultiScopeRevisionDoc(
            "Body alpha", "PriorAlice", "ins-a", "del-a", "PriorBob", "hdr-prior", "fn-prior");
        var right = MultiScopeRevisionDoc(
            "Body gamma", "PriorAlice", "ins-g", "del-g", "PriorBob", "hdr-prior", "fn-prior");

        var result = DocxDiff.Compare(left, right,
            new DocxDiffSettings { PreAcceptInputRevisions = true, AuthorForRevisions = "NewDiff" });

        // reject(result) ≡ accepted-view(left); accept(result) ≡ accepted-view(right) — across body, header,
        // and footnote scopes (the carried-over parts the flag cleans).
        Assert.Equal(
            AllScopesText(AcceptAll(left)),
            AllScopesText(RevisionProcessor.RejectRevisions(result)));
        Assert.Equal(
            AllScopesText(AcceptAll(right)),
            AllScopesText(RevisionAccepter.AcceptRevisions(result)));
    }

    // ---- oracle (e): schema validity -----------------------------------------------------------

    [Fact]
    public void Flag_on_output_introduces_no_new_schema_errors()
    {
        var left = MultiScopeRevisionDoc(
            "Body alpha", "PriorAlice", "ins-a", "del-a", "PriorBob", "hdr-prior", "fn-prior");
        var right = MultiScopeRevisionDoc(
            "Body gamma", "PriorAlice", "ins-g", "del-g", "PriorBob", "hdr-prior", "fn-prior");

        var result = DocxDiff.Compare(left, right,
            new DocxDiffSettings { PreAcceptInputRevisions = true, AuthorForRevisions = "NewDiff" });

        var baseline = Math.Max(SchemaErrorCount(AcceptAll(left)), SchemaErrorCount(AcceptAll(right)));
        Assert.True(SchemaErrorCount(result) <= baseline,
            $"flag-on output has more schema errors ({SchemaErrorCount(result)}) than accepted inputs ({baseline}).");
    }

    // ---- oracle (f): a redline-of-a-redline real contract (multi-author, incl. a move) ----------

    [Fact]
    public void Multi_author_redline_of_a_redline_round_trips_cleanly_with_flag()
    {
        // Build each side as a REAL multi-author redline via Consolidate of two reviewers off a shared base,
        // so each input carries UN-accepted w:ins/w:del from two distinct authors.
        var baseDoc = Docs.Para(
            "Alpha bravo charlie delta echo.",
            "Foxtrot golf hotel india juliet.",
            "Kilo lima mike november oscar.");

        var leftRedline = DocxDiff.Consolidate(baseDoc, new[]
        {
            new DocxDiffReviewer { Author = "Alice", Document = Docs.Para(
                "Alpha BRAVO charlie delta echo.", "Foxtrot golf hotel india juliet.", "Kilo lima mike november oscar.") },
            new DocxDiffReviewer { Author = "Bob", Document = Docs.Para(
                "Alpha bravo charlie delta echo.", "Foxtrot golf hotel india juliet.", "Kilo lima mike NOVEMBER oscar.") },
        });
        var rightRedline = DocxDiff.Consolidate(baseDoc, new[]
        {
            new DocxDiffReviewer { Author = "Carol", Document = Docs.Para(
                "Alpha bravo CHARLIE delta echo.", "Foxtrot golf hotel india juliet.", "Kilo lima mike november oscar.") },
            new DocxDiffReviewer { Author = "Dave", Document = Docs.Para(
                "Alpha bravo charlie delta echo.", "Foxtrot GOLF hotel india juliet.", "Kilo lima mike november oscar.") },
        });

        // Guard: the inputs genuinely carry multi-author, un-accepted tracked changes.
        Assert.True(RevisionAuthorsAllScopes(leftRedline).Count >= 2, "left input should carry ≥2 authors");
        Assert.True(RevisionAuthorsAllScopes(rightRedline).Count >= 2, "right input should carry ≥2 authors");

        var result = DocxDiff.Compare(leftRedline, rightRedline,
            new DocxDiffSettings { PreAcceptInputRevisions = true, AuthorForRevisions = "Final" });

        // (c) no stale authorship; (e) schema valid; well-formed (unique ids, no double-marking, paired moves).
        Assert.Equal(new HashSet<string> { "Final" }, RevisionAuthorsAllScopes(result));
        Assert.Equal(0, SchemaErrorCount(result));
        AssertWellFormedRevisions(result);

        // (d) round-trip against the accepted views (body structural projection).
        Assert.Equal(
            Docs.StructuralBody(AcceptAll(leftRedline)),
            Docs.StructuralBody(RevisionProcessor.RejectRevisions(result)));
        Assert.Equal(
            Docs.StructuralBody(AcceptAll(rightRedline)),
            Docs.StructuralBody(RevisionAccepter.AcceptRevisions(result)));
    }

    [Fact]
    public void Move_bearing_redline_input_round_trips_cleanly_with_flag()
    {
        // A redline whose input carries native w:moveFrom/w:moveTo (a relocation), produced by the engine.
        var baseDoc = Docs.Para(
            "Alpha bravo charlie delta echo.",
            "Foxtrot golf hotel india juliet.",
            "Kilo lima mike november oscar papa.");
        var moved = Docs.Para(
            "Kilo lima mike november oscar papa.",
            "Alpha bravo charlie delta echo.",
            "Foxtrot golf hotel india juliet.");

        var leftRedline = DocxDiff.Compare(baseDoc, moved, new DocxDiffSettings { AuthorForRevisions = "Mover" });
        var rightRedline = DocxDiff.Compare(baseDoc,
            Docs.Para("Alpha bravo charlie delta echo.",
                      "Foxtrot golf hotel india JULIET.",
                      "Kilo lima mike november oscar papa."),
            new DocxDiffSettings { AuthorForRevisions = "Editor" });

        // Guard: the left input genuinely carries native move markup.
        Assert.Contains(Wn + "moveFrom", AllRevisionElementNames(leftRedline));

        var result = DocxDiff.Compare(leftRedline, rightRedline,
            new DocxDiffSettings { PreAcceptInputRevisions = true, AuthorForRevisions = "Final" });

        Assert.Equal(new HashSet<string> { "Final" }, RevisionAuthorsAllScopes(result));
        Assert.Equal(0, SchemaErrorCount(result));
        AssertWellFormedRevisions(result);
        Assert.Equal(
            Docs.PlainText(AcceptAll(leftRedline)),
            Docs.PlainText(RevisionProcessor.RejectRevisions(result)));
        Assert.Equal(
            Docs.PlainText(AcceptAll(rightRedline)),
            Docs.PlainText(RevisionAccepter.AcceptRevisions(result)));
    }

    // ---- the flag flows through Consolidate too (base + every reviewer accepted) ----------------

    [Fact]
    public void Flag_on_cleans_carried_over_scopes_in_Consolidate()
    {
        // Base and a reviewer (derived from it) both carry a prior reviewer's header+footnote insertion
        // ("GhostRev"), identical on both → a carry-over leak. Their bodies differ so Consolidate has work.
        var baseDoc = MultiScopeRevisionDoc(
            "Base body one", "BodyPrior", "bi", "bd", "GhostRev", "hdr-g", "fn-g");
        var reviewerDoc = MultiScopeRevisionDoc(
            "Reviewer body two", "BodyPrior", "bi", "bd", "GhostRev", "hdr-g", "fn-g");
        var reviewers = new[] { new DocxDiffReviewer { Author = "Rev1", Document = reviewerDoc } };

        // Flag OFF: the prior "GhostRev" markup leaks through the consolidated output.
        var leaked = DocxDiff.Consolidate(baseDoc, reviewers,
            new DocxDiffConsolidateSettings { Diff = new DocxDiffSettings { AuthorForRevisions = "Rev1" } });
        Assert.Contains("GhostRev", RevisionAuthorsAllScopes(leaked));

        // Flag ON: base and reviewer are accepted first, so no stale authorship survives.
        var clean = DocxDiff.Consolidate(baseDoc, reviewers,
            new DocxDiffConsolidateSettings
            {
                Diff = new DocxDiffSettings { PreAcceptInputRevisions = true, AuthorForRevisions = "Rev1" },
            });
        Assert.DoesNotContain("GhostRev", RevisionAuthorsAllScopes(clean));

        // Consolidate round-trip: reject ≡ base — with the flag, the ACCEPTED view of the base.
        Assert.Equal(
            Docs.PlainText(AcceptAll(baseDoc)),
            Docs.PlainText(RevisionProcessor.RejectRevisions(clean)));
    }

    // ---- the flag's effect is scoped to the rendered package, NOT the consumer revision list ----

    [Fact]
    public void Flag_does_not_change_the_consumer_revision_list()
    {
        // GetRevisions already reads the accepted view, so the flag (which only additionally cleans the
        // carried-over OUTPUT parts) leaves the consumer revision list identical.
        var left = MultiScopeRevisionDoc(
            "Body alpha", "PriorAlice", "ins-a", "del-a", "PriorBob", "hdr-prior", "fn-prior");
        var right = MultiScopeRevisionDoc(
            "Body gamma", "PriorAlice", "ins-g", "del-g", "PriorBob", "hdr-prior", "fn-prior");

        var off = DocxDiff.GetRevisions(left, right, new DocxDiffSettings { AuthorForRevisions = "X" });
        var on = DocxDiff.GetRevisions(left, right,
            new DocxDiffSettings { PreAcceptInputRevisions = true, AuthorForRevisions = "X" });

        Assert.Equal(
            off.Select(r => $"{r.Type}:{r.Text}").ToList(),
            on.Select(r => $"{r.Type}:{r.Text}").ToList());
    }

    [Fact]
    public void Flag_does_not_change_the_edit_script_json()
    {
        // GetEditScriptJson is also a diff-as-data view built off the (already-accepted) edit script, so the flag
        // is a no-op there too — exercises the GetEditScriptJson PreAccept wiring and pins the invariance.
        var left = MultiScopeRevisionDoc(
            "Body alpha", "PriorAlice", "ins-a", "del-a", "PriorBob", "hdr-prior", "fn-prior");
        var right = MultiScopeRevisionDoc(
            "Body gamma", "PriorAlice", "ins-g", "del-g", "PriorBob", "hdr-prior", "fn-prior");

        var off = DocxDiff.GetEditScriptJson(left, right, new DocxDiffSettings { AuthorForRevisions = "X" });
        var on = DocxDiff.GetEditScriptJson(left, right,
            new DocxDiffSettings { PreAcceptInputRevisions = true, AuthorForRevisions = "X" });

        Assert.Equal(off, on);
    }

    // ---- well-formedness helper ----------------------------------------------------------------

    /// <summary>
    /// Assert NECESSARY (not exhaustive) well-formedness conditions on the produced revisions: every
    /// <c>w:ins</c>/<c>w:del</c> id is unique across the whole document (a duplicate id is the double-marking
    /// signature), no <c>w:ins</c> nests directly inside another <c>w:ins</c> (nor del-in-del), and the set of
    /// <c>w:moveFromRangeStart</c> names equals the set of <c>w:moveToRangeStart</c> names (no orphaned move
    /// range on either side). These catch the concrete malformations a stale-input-revision leak would produce;
    /// schema validity (oracle e) is asserted separately.
    /// </summary>
    private static void AssertWellFormedRevisions(WmlDocument d)
    {
        var insDelIds = new List<string>();
        var fromNames = new HashSet<string>();
        var toNames = new HashSet<string>();
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        foreach (var part in ContentParts(doc))
        {
            XDocument xd;
            using (var s = part.GetStream(FileMode.Open, FileAccess.Read)) xd = XDocument.Load(s);
            foreach (var el in xd.Descendants())
            {
                if (el.Name == Wn + "ins" || el.Name == Wn + "del")
                {
                    var id = (string?)el.Attribute(Wn + "id");
                    if (id != null) insDelIds.Add(id);
                    Assert.DoesNotContain(el.Ancestors(), a => a.Name == el.Name); // no ins-in-ins / del-in-del
                }
            }
            foreach (var el in xd.Descendants(Wn + "moveFromRangeStart"))
            {
                var n = (string?)el.Attribute(Wn + "name");
                if (n != null) fromNames.Add(n);
            }
            foreach (var el in xd.Descendants(Wn + "moveToRangeStart"))
            {
                var n = (string?)el.Attribute(Wn + "name");
                if (n != null) toNames.Add(n);
            }
        }

        Assert.Equal(insDelIds.Count, insDelIds.Distinct().Count()); // unique ins/del ids — no double-marking
        Assert.Equal(fromNames, toNames);                            // every move-from range pairs with a move-to
    }
}
