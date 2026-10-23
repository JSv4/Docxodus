using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DocxSession.ReplaceTextAtSpan"/> inside and beside existing tracked revisions: the
/// span path the browser editor sends every keystroke through. Each test reads the paragraph back
/// as a compact shape — <c>"text"</c> for a run, <c>ins:Author[...]</c> for a revision envelope —
/// so the markup a reviewer would see is the assertion.
/// </summary>
public class DocxSessionSpanRevisionTests
{
    private const string Date = "2026-05-01T09:00:00Z";

    // ─── Deleting from the session author's own pending insertion (issue #1316) ───

    [Fact]
    public void Backspace_InOwnPendingInsertion_RemovesTheCharacter()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Parties "));
        Assert.True(session.ReplaceTextAtSpan(p, 8, 0, "bindingg").Success);

        var edit = session.ReplaceTextAtSpan(p, 15, 1, "");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("\"Parties \" ins:Pat[\"binding\"]", Shape(session, p));
        Assert.Equal(new[] { "insert:Pat:binding" }, Revisions(session));
        Assert.Equal("Parties binding", Accepted(session));
        Assert.Equal("Parties ", Rejected(session));
    }

    [Fact]
    public void Correction_InOwnPendingInsertion_StaysOneInsertion()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Fee: "));
        Assert.True(session.ReplaceTextAtSpan(p, 5, 0, "teh clause").Success);

        var edit = session.ReplaceTextAtSpan(p, 6, 2, "he");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("\"Fee: \" ins:Pat[\"the clause\"]", Shape(session, p));
        Assert.Equal(new[] { "insert:Pat:the clause" }, Revisions(session));
    }

    [Fact]
    public void DeletingAllOfOwnPendingInsertion_LeavesNoRevision()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Parties "));
        Assert.True(session.ReplaceTextAtSpan(p, 8, 0, "abc").Success);

        Assert.True(session.ReplaceTextAtSpan(p, 8, 3, "").Success);

        Assert.Equal("\"Parties \"", Shape(session, p));
        Assert.Empty(Revisions(session));
    }

    [Fact]
    public void Replacement_AcrossOriginalTextAndOwnInsertion_DeletesOnlyTheOriginal()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat",
            Run("Parties "), Ins("Pat", 1, Run("binding")));

        // "s bi": the original "s " is deleted under tracking, Pat's own "bi" is simply removed.
        var edit = session.ReplaceTextAtSpan(p, 6, 4, "X");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("\"Partie\" del:Pat[\"s \"] ins:Pat[\"X\"] ins:Pat[\"nding\"]", Shape(session, p));
        Assert.Equal("PartieXnding", Accepted(session));
        Assert.Equal("Parties ", Rejected(session));
    }

    [Fact]
    public void Control_BackspaceInAnotherAuthorsInsertion_IsADeletionInsideIt()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat",
            Run("Parties "), Ins("Alex", 1, Run("bindingg")));

        Assert.True(session.ReplaceTextAtSpan(p, 15, 1, "").Success);

        Assert.Equal("\"Parties \" ins:Alex[\"binding\" del:Pat[\"g\"]]", Shape(session, p));
        Assert.Equal("Parties binding", Accepted(session));
    }

    [Fact]
    public void Control_BackspaceInOwnMoveDestination_IsADeletionInsideIt()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat",
            Run("Parties "), new XElement(W.moveTo, Stamp("Pat", 1), Run("bindingg")));

        Assert.True(session.ReplaceTextAtSpan(p, 15, 1, "").Success);

        Assert.Equal("\"Parties \" moveTo:Pat[\"binding\" del:Pat[\"g\"]]", Shape(session, p));
    }

    [Theory]
    [InlineData(4, 1, "", "\"Pa \" ins:Pat[\"a⇥cd\"]")]
    [InlineData(5, 1, "X", "\"Pa \" ins:Pat[\"ab⇥Xd\"]")]
    public void EditInOwnPendingInsertion_KeepsATabBetweenTwoTextNodesInPlace(int start, int length, string replace, string shape)
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Pa "), Ins("Pat", 1,
            new XElement(W.r, new XElement(W.t, "ab"), new XElement(W.tab), new XElement(W.t, "cd"))));

        Assert.True(session.ReplaceTextAtSpan(p, start, length, replace).Success);

        Assert.Equal(shape, Shape(session, p));
    }

    [Fact]
    public void EditInOwnPendingInsertion_KeepsABreakBetweenTwoTextNodesInPlace()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Pa "), Ins("Pat", 1,
            new XElement(W.r, new XElement(W.t, "ab"), new XElement(W.br), new XElement(W.t, "cd"))));

        Assert.True(session.ReplaceTextAtSpan(p, 3, 2, "").Success);

        Assert.Equal("\"Pa \" ins:Pat[\"↵cd\"]", Shape(session, p));
    }

    [Fact]
    public void ReplaceTextRangeAndReplaceMatch_InOwnPendingInsertion_EditItInPlace()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Fee: "));
        Assert.True(session.ReplaceTextAtSpan(p, 5, 0, "teh clause").Success);

        Assert.True(Assert.Single(session.ReplaceTextRange(p, "teh", "the")).Success);
        Assert.True(session.ReplaceMatch(Assert.Single(session.Grep("clause")), "term").Success);

        Assert.Equal("\"Fee: \" ins:Pat[\"the term\"]", Shape(session, p));
        Assert.Equal(new[] { "insert:Pat:the term" }, Revisions(session));
    }

    [Fact]
    public void BackspaceInOwnPendingInsertion_UndoesAndRedoes()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Parties "));
        Assert.True(session.ReplaceTextAtSpan(p, 8, 0, "bindingg").Success);
        var typed = Shape(session, p);

        Assert.True(session.ReplaceTextAtSpan(p, 15, 1, "").Success);
        Assert.True(session.Undo());
        Assert.Equal(typed, Shape(session, p));
        Assert.True(session.Redo());

        Assert.Equal("\"Parties \" ins:Pat[\"binding\"]", Shape(session, p));
    }

    [Fact]
    public void BackspaceInOwnPendingInsertionInsideAHyperlink_RemovesTheCharacter()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("See "),
            new XElement(W.hyperlink, new XAttribute(W.anchor, "terms"), Ins("Pat", 1, Run("termss"))));

        Assert.True(session.ReplaceTextAtSpan(p, 9, 1, "").Success);

        Assert.Equal("\"See \" hyperlink[ins:Pat[\"terms\"]]", Shape(session, p));
    }

    [Fact]
    public void EditInOwnPendingInsertion_KeepsABookmarkInside()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Pa "),
            Ins("Pat", 1, Bookmark(W.bookmarkStart), Run("bindingg"), Bookmark(W.bookmarkEnd)));

        Assert.True(session.ReplaceTextAtSpan(p, 10, 1, "").Success);

        Assert.Equal("\"Pa \" ins:Pat[bookmarkStart[] \"binding\" bookmarkEnd[]]", Shape(session, p));
    }

    [Fact]
    public void Control_RemovingAPictureFromAnotherAuthorsInsertion_IsRefused()
    {
        // A control for the image ops' own-insertion check, now OwnPendingInsertion: another
        // author's insertion is refused before that check runs, so this passed before too.
        // Alex inserts a picture under tracking; Pat then tries to remove it with tracking on.
        byte[] bytes;
        using (var alex = new DocxSession(Build(Run("Logo: ")), new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Alex",
        }))
        {
            var anchor = alex.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id;
            Assert.True(alex.InsertImage(anchor, 6, DocxSessionImageTests.Png(2, 3)).Success);
            bytes = alex.Save();
        }
        using var pat = new DocxSession(bytes, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Pat",
        });
        var before = pat.GetPackageContentHash();

        var removed = pat.RemoveImage(Assert.Single(pat.ListImages()).Id);

        Assert.False(removed.Success);
        Assert.Equal(before, pat.GetPackageContentHash());
    }

    // ─── Untracked typing strictly inside a tracked insertion (issue #1319) ───

    /// <summary>Alex's insertion <c>amount</c> as one run, or as <c>amo</c> + bold <c>unt</c>.</summary>
    private static object[] FeeAmountDue(string author, bool twoRuns) => new object[]
    {
        Run("Fee "),
        Ins(author, 1, twoRuns ? new object[] { Run("amo"), Run("unt", bold: true) } : new object[] { Run("amount") }),
        Run(" due"),
    };

    [Theory]
    [InlineData("Alex", false)]
    [InlineData("Alex", true)]
    [InlineData("Pat", false)]
    [InlineData("Pat", true)]
    public void Untracked_TextTypedInsideAnInsertion_SplitsItAroundAPlainRun(string insertionAuthor, bool twoRuns)
    {
        var (session, p) = Open(TrackedChangeMode.Accept, "Pat", FeeAmountDue(insertionAuthor, twoRuns));

        var edit = session.ReplaceTextAtSpan(p, 7, 0, "XY");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal($"\"Fee \" ins:{insertionAuthor}[\"amo\"] \"XY\" ins:{insertionAuthor}[\"unt\"] \" due\"",
            Shape(session, p));
        var halves = XElement.Parse(session.Raw.GetXml(p)).Elements(W.ins).ToArray();
        Assert.Equal(Date, (string?)halves[1].Attribute(W.date));
        Assert.NotEqual((string?)halves[0].Attribute(W.id), (string?)halves[1].Attribute(W.id));
        Assert.Equal("Fee amoXYunt due", Accepted(session));
        Assert.Equal("Fee XY due", Rejected(session));
    }

    [Fact]
    public void Untracked_SplitInsideAnInsertion_KeepsEachHalfsFormatting()
    {
        var (session, p) = Open(TrackedChangeMode.Accept, "Pat", FeeAmountDue("Alex", twoRuns: true));

        Assert.True(session.ReplaceTextAtSpan(p, 7, 0, "XY").Success);

        var halves = XElement.Parse(session.Raw.GetXml(p)).Elements(W.ins).ToArray();
        Assert.Null(halves[0].Element(W.r)!.Element(W.rPr));
        Assert.NotNull(halves[1].Element(W.r)!.Element(W.rPr)?.Element(W.b));
    }

    [Theory]
    [InlineData(4, "\"Fee XY\" ins:Alex[\"amount\"] \" due\"")]
    [InlineData(10, "\"Fee \" ins:Alex[\"amount\"] \"XY\" \" due\"")]
    public void Control_UntrackedTextAtAnInsertionsEdge_StaysOutsideIt(int offset, string shape)
    {
        var (session, p) = Open(TrackedChangeMode.Accept, "Pat", FeeAmountDue("Alex", twoRuns: false));

        Assert.True(session.ReplaceTextAtSpan(p, offset, 0, "XY").Success);

        Assert.Equal(shape, Shape(session, p));
    }


    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Untracked_TextTypedInsideAnInsertionInAHyperlink_SplitsOnlyTheInsertion(bool twoRuns)
    {
        var amount = (XElement)FeeAmountDue("Alex", twoRuns)[1];
        var (session, p) = Open(TrackedChangeMode.Accept, "Pat",
            Run("Fee "), new XElement(W.hyperlink, new XAttribute(W.anchor, "fees"), amount), Run(" due"));

        Assert.True(session.ReplaceTextAtSpan(p, 7, 0, "XY").Success);

        Assert.Equal("\"Fee \" hyperlink[ins:Alex[\"amo\"] \"XY\" ins:Alex[\"unt\"]] \" due\"", Shape(session, p));
        Assert.Equal("Fee XY due", Rejected(session));
    }

    [Fact]
    public void Untracked_TextTypedInsideAnInsertionInAContentControl_SplitsOnlyTheInsertion()
    {
        var amount = (XElement)FeeAmountDue("Alex", twoRuns: true)[1];
        var (session, p) = Open(TrackedChangeMode.Accept, "Pat",
            Run("Fee "), new XElement(W.sdt, new XElement(W.sdtPr), new XElement(W.sdtContent, amount)), Run(" due"));

        Assert.True(session.ReplaceTextAtSpan(p, 7, 0, "XY").Success);

        Assert.Equal("\"Fee \" sdt[sdtPr[] sdtContent[ins:Alex[\"amo\"] \"XY\" ins:Alex[\"unt\"]]] \" due\"", Shape(session, p));
    }

    [Fact]
    public void Untracked_TextTypedInsideAContentControlInAnInsertion_StaysWhereItLands()
    {
        // Nothing between the text and the insertion can be cut, so it behaves as before.
        var (session, p) = Open(TrackedChangeMode.Accept, "Pat", Run("Fee "), Ins("Alex", 1,
            new XElement(W.sdt, new XElement(W.sdtPr), new XElement(W.sdtContent, Run("amo"), Run("unt", bold: true)))), Run(" due"));

        Assert.True(session.ReplaceTextAtSpan(p, 7, 0, "XY").Success);

        Assert.Equal("\"Fee \" ins:Alex[sdt[sdtPr[] sdtContent[\"amo\" \"XY\" \"unt\"]]] \" due\"", Shape(session, p));
    }

    [Fact]
    public void Untracked_ReplacementInsideAnInsertion_GoesOutsideIt()
    {
        var (session, p) = Open(TrackedChangeMode.Accept, "Pat", FeeAmountDue("Alex", twoRuns: false));

        Assert.True(session.ReplaceTextAtSpan(p, 6, 2, "XY").Success);

        Assert.Equal("\"Fee \" ins:Alex[\"am\"] \"XY\" ins:Alex[\"nt\"] \" due\"", Shape(session, p));
        Assert.Equal("Fee amXYnt due", Accepted(session));
        Assert.Equal("Fee XY due", Rejected(session));
    }

    [Fact]
    public void Untracked_ReplacementInsideAnInsertion_KeepsATabOnItsSide()
    {
        // The editor's fallback for an offset the zero-length path refuses: a run with a tab after text.
        var (session, p) = Open(TrackedChangeMode.Accept, "Pat", Run("Fee "), Ins("Alex", 1,
            new XElement(W.r, new XElement(W.t, "amo"), new XElement(W.tab), new XElement(W.t, "unt"))), Run(" due"));

        Assert.True(session.ReplaceTextAtSpan(p, 5, 1, "mXY").Success);

        Assert.Equal("\"Fee \" ins:Alex[\"a\"] \"mXY\" ins:Alex[\"o⇥unt\"] \" due\"", Shape(session, p));
    }

    [Fact]
    public void Untracked_ReplacementInAFieldResultInsideAnInsertion_SurvivesItsReject()
    {
        var (session, p) = Open(TrackedChangeMode.Accept, "Pat", Run("Fee "), Ins("Alex", 1, Run("amo"),
            FieldChar("begin"), new XElement(W.r, new XElement(W.instrText, " PAGE ")), FieldChar("separate"),
            Run("12"), FieldChar("end")), Run(" due"));

        Assert.True(session.ReplaceTextAtSpan(p, 7, 1, "1XY").Success);

        Assert.Equal("Fee 1XY due", Rejected(session));
    }

    [Fact]
    public void Untracked_TextAtTheStartOfAnInsertionAfterAnEmptyRun_LeavesNoEmptyHalf()
    {
        var (session, p) = Open(TrackedChangeMode.Accept, "Pat",
            Ins("Alex", 1, new XElement(W.r, new XElement(W.rPr, new XElement(W.b))), Run("amount")), Run(" due"));

        Assert.True(session.ReplaceTextAtSpan(p, 0, 0, "XY").Success);

        Assert.Equal(new[] { "insert:Alex:amount" }, Revisions(session));
    }

    private static object[] FeeMovedAmountDue(bool twoRuns) => new object[]
    {
        Run("Fee "),
        new XElement(W.moveToRangeStart, Stamp("Alex", 10), new XAttribute(W.name, "move1")),
        new XElement(W.moveTo, Stamp("Alex", 11),
            twoRuns ? new object[] { Run("amo"), Run("unt", bold: true) } : new object[] { Run("amount") }),
        new XElement(W.moveToRangeEnd, new XAttribute(W.id, 10)),
        Run(" due"),
    };

    [Theory]
    [InlineData(TrackedChangeMode.Accept, false, 7, 0, "XY")]
    [InlineData(TrackedChangeMode.Accept, true, 7, 0, "XY")]
    [InlineData(TrackedChangeMode.Accept, false, 6, 1, "oXY")]
    [InlineData(TrackedChangeMode.RenderInline, false, 7, 0, "XY")]
    [InlineData(TrackedChangeMode.RenderInline, true, 7, 0, "XY")]
    [InlineData(TrackedChangeMode.RenderInline, false, 6, 1, "oXY")]
    public void TextInsideAMoveDestination_IsRefused(TrackedChangeMode mode, bool twoRuns, int start, int length, string text)
    {
        var (session, p) = Open(mode, "Pat", FeeMovedAmountDue(twoRuns));
        var before = session.GetPackageContentHash();

        var edit = session.ReplaceTextAtSpan(p, start, length, text);

        Assert.False(edit.Success);
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary, edit.Error!.Code);
        Assert.Equal(before, session.GetPackageContentHash());
    }

    [Theory]
    [InlineData(TrackedChangeMode.Accept)]
    [InlineData(TrackedChangeMode.RenderInline)]
    public void TextAtTheEndOfAMoveDestination_LandsOutsideTheMovedRange(TrackedChangeMode mode)
    {
        var (session, p) = Open(mode, "Pat", FeeMovedAmountDue(twoRuns: false));

        Assert.True(session.ReplaceTextAtSpan(p, 10, 0, "XY").Success);

        var shape = Shape(session, p);
        Assert.True(shape.IndexOf("moveToRangeEnd", StringComparison.Ordinal) < shape.IndexOf("XY", StringComparison.Ordinal), shape);
        Assert.Equal("Fee amountXY due", Accepted(session));
        // Untracked, the text is nobody's change and survives rejecting the move.
        if (mode == TrackedChangeMode.Accept) Assert.Equal("Fee XY due", Rejected(session));
    }

    // ─── Tracked typing strictly inside an insertion (issue #1318) ───

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Tracked_TextTypedInsideAnotherAuthorsInsertion_SplitsItAroundOurs(bool twoRuns)
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", FeeAmountDue("Alex", twoRuns));

        var edit = session.ReplaceTextAtSpan(p, 7, 0, "XY");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("\"Fee \" ins:Alex[\"amo\"] ins:Pat[\"XY\"] ins:Alex[\"unt\"] \" due\"", Shape(session, p));
        var insertions = XElement.Parse(session.Raw.GetXml(p)).Elements(W.ins).ToArray();
        Assert.Equal(3, insertions.Select(i => (string?)i.Attribute(W.id)).Distinct().Count());
        Assert.Equal(("Alex", Date), ((string?)insertions[2].Attribute(W.author), (string?)insertions[2].Attribute(W.date)));
        Assert.Equal(new[] { "insert:Alex:amo", "insert:Pat:XY", "insert:Alex:unt" }, Revisions(session));
        Assert.Equal("Fee amoXYunt due", Accepted(session));
        Assert.Equal("Fee  due", Rejected(session));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Tracked_TextTypedInsideAnotherAuthorsInsertionInAHyperlink_SplitsOnlyTheirInsertion(bool twoRuns)
    {
        var amount = (XElement)FeeAmountDue("Alex", twoRuns)[1];
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat",
            Run("Fee "), new XElement(W.hyperlink, new XAttribute(W.anchor, "fees"), amount), Run(" due"));

        Assert.True(session.ReplaceTextAtSpan(p, 7, 0, "XY").Success);

        Assert.Equal("\"Fee \" hyperlink[ins:Alex[\"amo\"] ins:Pat[\"XY\"] ins:Alex[\"unt\"]] \" due\"", Shape(session, p));
        Assert.Contains("insert:Pat:XY", Revisions(session));
    }

    [Fact]
    public void Tracked_TextTypedInsideAnotherAuthorsInsertionInAContentControl_SplitsOnlyTheirInsertion()
    {
        var amount = (XElement)FeeAmountDue("Alex", twoRuns: true)[1];
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat",
            Run("Fee "), new XElement(W.sdt, new XElement(W.sdtPr), new XElement(W.sdtContent, amount)), Run(" due"));

        Assert.True(session.ReplaceTextAtSpan(p, 7, 0, "XY").Success);

        Assert.Equal("\"Fee \" sdt[sdtPr[] sdtContent[ins:Alex[\"amo\"] ins:Pat[\"XY\"] ins:Alex[\"unt\"]]] \" due\"", Shape(session, p));
    }

    [Fact]
    public void Tracked_ReplacementInsideAnotherAuthorsInsertion_PutsOursBesideTheirs()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", FeeAmountDue("Alex", twoRuns: false));

        Assert.True(session.ReplaceTextAtSpan(p, 6, 2, "XY").Success);

        Assert.Equal("\"Fee \" ins:Alex[\"am\" del:Pat[\"ou\"]] ins:Pat[\"XY\"] ins:Alex[\"nt\"] \" due\"", Shape(session, p));
        Assert.Equal("Fee amXYnt due", Accepted(session));
        Assert.Equal("Fee  due", Rejected(session));
    }

    [Theory]
    [InlineData(false, "\"Fee \" ins:Pat[\"amoXYunt\"] \" due\"")]
    [InlineData(true, "\"Fee \" ins:Pat[\"amoXY\" \"unt\"] \" due\"")]
    public void Tracked_TextTypedInsideOwnInsertion_ExtendsIt(bool twoRuns, string shape)
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", FeeAmountDue("Pat", twoRuns));

        var edit = session.ReplaceTextAtSpan(p, 7, 0, "XY");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal(shape, Shape(session, p));
        Assert.Equal(new[] { "insert:Pat:amoXYunt" }, Revisions(session));
    }

    [Fact]
    public void Tracked_TextTypedInsideOwnInsertion_DoesNotTakeAnotherReviewersFormatChange()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Fee "), Ins("Pat", 1,
            new XElement(W.r, new XElement(W.rPr, new XElement(W.b), new XElement(W.rPrChange, Stamp("Bob", 2), new XElement(W.rPr))),
                new XElement(W.t, "amount"))), Run(" due"));

        Assert.True(session.ReplaceTextAtSpan(p, 7, 0, "XY").Success);

        Assert.Equal("\"Fee \" ins:Pat[\"amo\" \"XY\" \"unt\"] \" due\"", Shape(session, p));
        var xy = XElement.Parse(session.Raw.GetXml(p)).Descendants(W.r).Single(r => r.Value == "XY");
        Assert.NotNull(xy.Element(W.rPr)?.Element(W.b));
        Assert.Null(xy.Descendants(W.rPrChange).FirstOrDefault());
    }

    [Fact]
    public void Tracked_FormattedTextTypedInsideAnotherAuthorsInsertion_FormatsOnlyTheNewText()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", FeeAmountDue("Alex", twoRuns: false));

        var edit = session.ReplaceTextAtSpanWithFormat(p, 7, 0, "XY", new FormatOp { Bold = true });

        Assert.True(edit.Success, edit.Error?.Message);
        var runs = session.GetFormatting(p)!.Runs;
        Assert.Equal("Fee amoXYunt due", string.Concat(runs.Select(r => r.Text)));
        Assert.Equal("XY", string.Concat(runs.Where(r => r.Effective.Bold is true).Select(r => r.Text)));
        Assert.Equal("\"Fee \" ins:Alex[\"amo\"] ins:Pat[\"XY\"] ins:Alex[\"unt\"] \" due\"", Shape(session, p));
    }

    [Fact]
    public void Control_TrackedTextAtTheEndOfAnotherAuthorsInsertion_IsASiblingInsertion()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", FeeAmountDue("Alex", twoRuns: false));

        Assert.True(session.ReplaceTextAtSpan(p, 10, 0, "XY").Success);

        Assert.Equal("\"Fee \" ins:Alex[\"amount\"] ins:Pat[\"XY\"] \" due\"", Shape(session, p));
    }

    [Theory]
    [InlineData(TrackedChangeMode.Accept, "Alex", "XY", "\"Fee\" bookmarkStart[] \"XY\" bookmarkEnd[] \" due\"")]
    [InlineData(TrackedChangeMode.RenderInline, "Pat", "", "\"Fee\" bookmarkStart[] bookmarkEnd[] \" due\"")]
    public void RemovingAllOfAnInsertionAroundABookmark_LeavesNoEmptyInsertion(
        TrackedChangeMode mode, string insertionAuthor, string replace, string shape)
    {
        // Untracked, Alex's whole insertion is typed over; tracked, Pat deletes all of her own.
        var (session, p) = Open(mode, "Pat", Run("Fee"),
            Ins(insertionAuthor, 1, Bookmark(W.bookmarkStart), Run(" amount"), Bookmark(W.bookmarkEnd)), Run(" due"));

        var edit = session.ReplaceTextAtSpan(p, 3, 7, replace);

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal(shape, Shape(session, p));
        Assert.DoesNotContain(Revisions(session), r => r.StartsWith("insert:", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(TrackedChangeMode.Accept, "Alex", false)]
    [InlineData(TrackedChangeMode.Accept, "Alex", true)]
    [InlineData(TrackedChangeMode.Accept, "Pat", false)]
    [InlineData(TrackedChangeMode.Accept, "Pat", true)]
    [InlineData(TrackedChangeMode.RenderInline, "Alex", false)]
    [InlineData(TrackedChangeMode.RenderInline, "Alex", true)]
    [InlineData(TrackedChangeMode.RenderInline, "Pat", false)]
    [InlineData(TrackedChangeMode.RenderInline, "Pat", true)]
    public void TypingInsideAnInsertionInASimpleField_IsRefusedWithoutMutation(
        TrackedChangeMode tracking, string insertionAuthor, bool twoRuns)
    {
        var insertion = (XElement)FeeAmountDue(insertionAuthor, twoRuns)[1];
        var (session, p) = Open(tracking, "Pat", Run("Fee "),
            new XElement(W.fldSimple, new XAttribute(W.instr, " DOCPROPERTY Title "), insertion), Run(" due"));
        var before = session.Raw.GetXml(p);
        var version = session.Version;
        var revisions = Revisions(session);

        var edit = session.ReplaceTextAtSpan(p, 7, 0, "XY");

        Assert.False(edit.Success);
        Assert.Equal(EditErrorCode.OffsetOutOfRange, edit.Error?.Code);
        Assert.Equal(before, session.Raw.GetXml(p));
        Assert.Equal(version, session.Version);
        Assert.Equal(revisions, Revisions(session));
    }

    [Theory]
    [InlineData(TrackedChangeMode.Accept, 4)]
    [InlineData(TrackedChangeMode.Accept, 10)]
    [InlineData(TrackedChangeMode.RenderInline, 4)]
    [InlineData(TrackedChangeMode.RenderInline, 10)]
    public void TypingAtASimpleFieldEdge_GoesOutsideTheCachedResult(TrackedChangeMode tracking, int offset)
    {
        var field = new XElement(W.fldSimple, new XAttribute(W.instr, " DOCPROPERTY Title "),
            Ins("Alex", 1, Run("amount")));
        var (session, p) = Open(tracking, "Pat", Run("Fee "), field, Run(" due"));

        var edit = session.ReplaceTextAtSpan(p, offset, 0, "XY");

        Assert.True(edit.Success, edit.Error?.Message);
        var paragraph = XElement.Parse(session.Raw.GetXml(p));
        Assert.Equal("amount", paragraph.Element(W.fldSimple)?.Value);
        Assert.Equal(offset == 4 ? "Fee XYamount due" : "Fee amountXY due", Accepted(session));
        if (tracking == TrackedChangeMode.RenderInline)
            Assert.Equal("Fee  due", Rejected(session));
        else
            Assert.Equal("Fee XY due", Rejected(session));
    }

    // ─── Which side of tracked deleted text a zero-length insertion goes on (issue #1290) ───

    internal static object[] TenantPromptlyRepaint() => new object[]
    {
        Run("The tenant shall "), Del("A", 1, "promptly "), Run("repaint"),
    };

    [Theory]
    [InlineData(null, "\"The tenant shall \" ins:Pat[\"Y\"] del:A[\"promptly \"] \"repaint\"")]
    [InlineData(DeletedTextSide.Before, "\"The tenant shall \" ins:Pat[\"Y\"] del:A[\"promptly \"] \"repaint\"")]
    [InlineData(DeletedTextSide.After, "\"The tenant shall \" del:A[\"promptly \"] ins:Pat[\"Y\"] \"repaint\"")]
    public void Tracked_InsertionBesideADeletion_GoesOnTheRequestedSide(DeletedTextSide? side, string shape)
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", TenantPromptlyRepaint());

        var edit = side is { } s ? session.ReplaceTextAtSpan(p, 17, 0, "Y", s) : session.ReplaceTextAtSpan(p, 17, 0, "Y");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal(shape, Shape(session, p));
        Assert.Equal("The tenant shall Yrepaint", Accepted(session));
        Assert.Equal("The tenant shall promptly repaint", Rejected(session));
    }

    [Theory]
    [InlineData(DeletedTextSide.Before, "\"The tenant shall Y\" del:A[\"promptly \"] \"repaint\"")]
    [InlineData(DeletedTextSide.After, "\"The tenant shall \" del:A[\"promptly \"] \"Yrepaint\"")]
    public void Untracked_InsertionBesideADeletion_ExtendsTheRunOnTheRequestedSide(DeletedTextSide side, string shape)
    {
        var (session, p) = Open(TrackedChangeMode.Accept, "Pat", TenantPromptlyRepaint());

        Assert.True(session.ReplaceTextAtSpan(p, 17, 0, "Y", side).Success);

        Assert.Equal(shape, Shape(session, p));
    }

    [Theory]
    [InlineData(null, "ins:Pat[\"Y\"] bookmarkStart[] del:A[\"Gone heading\"] bookmarkEnd[]")]
    [InlineData(DeletedTextSide.Before, "bookmarkStart[] ins:Pat[\"Y\"] del:A[\"Gone heading\"] bookmarkEnd[]")]
    [InlineData(DeletedTextSide.After, "bookmarkStart[] del:A[\"Gone heading\"] ins:Pat[\"Y\"] bookmarkEnd[]")]
    public void Tracked_InsertionIntoADeletedHeadingInsideABookmark_StaysInTheBookmarkWhenASideIsGiven(
        DeletedTextSide? side, string shape)
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat",
            Bookmark(W.bookmarkStart), Del("A", 1, "Gone heading"), Bookmark(W.bookmarkEnd));

        var edit = side is { } s ? session.ReplaceTextAtSpan(p, 0, 0, "Y", s) : session.ReplaceTextAtSpan(p, 0, 0, "Y");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal(shape, Shape(session, p));
    }

    [Theory]
    [InlineData(DeletedTextSide.Before, "\"See \" commentRangeStart[] ins:Pat[\"Y\"] del:A[\"old \"] \"text\"")]
    [InlineData(DeletedTextSide.After, "\"See \" commentRangeStart[] del:A[\"old \"] ins:Pat[\"Y\"] \"text\"")]
    public void Tracked_InsertionBesideADeletion_LeavesARangeMarkerWithTheLiveText(DeletedTextSide side, string shape)
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat",
            Run("See "), new XElement(W.commentRangeStart, new XAttribute(W.id, 0)), Del("A", 1, "old "), Run("text"));

        Assert.True(session.ReplaceTextAtSpan(p, 4, 0, "Y", side).Success);

        Assert.Equal(shape, Shape(session, p));
    }

    [Theory]
    [InlineData(DeletedTextSide.Before)]
    [InlineData(DeletedTextSide.After)]
    public void Control_WithNoDeletionAtTheOffset_TheSideChangesNothing(DeletedTextSide side)
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Fee "), Run("due", bold: true));

        Assert.True(session.ReplaceTextAtSpan(p, 4, 0, "Y", side).Success);

        Assert.Equal("\"Fee \" ins:Pat[\"Y\"] \"due\"", Shape(session, p));
    }

    [Fact]
    public void Tracked_FormattedInsertionAfterADeletion_GoesAfterItAndIsFormatted()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", TenantPromptlyRepaint());

        var edit = session.ReplaceTextAtSpanWithFormat(p, 17, 0, "Y", new FormatOp { Bold = true }, DeletedTextSide.After);

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("\"The tenant shall \" del:A[\"promptly \"] ins:Pat[\"Y\"] \"repaint\"", Shape(session, p));
        Assert.Equal("Y", string.Concat(session.GetFormatting(p)!.Runs.Where(r => r.Effective.Bold is true).Select(r => r.Text)));
    }

    [Theory]
    [InlineData(TrackedChangeMode.Accept, DeletedTextSide.Before)]
    [InlineData(TrackedChangeMode.Accept, DeletedTextSide.After)]
    [InlineData(TrackedChangeMode.RenderInline, DeletedTextSide.Before)]
    [InlineData(TrackedChangeMode.RenderInline, DeletedTextSide.After)]
    public void DeletedTextSide_InsideALiveRun_DoesNotJumpToALaterDeletion(
        TrackedChangeMode tracking, DeletedTextSide side)
    {
        var (session, p) = Open(tracking, "Pat", TenantPromptlyRepaint());

        var edit = session.ReplaceTextAtSpan(p, 2, 0, "Y", side);

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("ThYe tenant shall repaint", Accepted(session));
        Assert.Contains("delete:A:promptly ", Revisions(session));
        Assert.Equal(tracking == TrackedChangeMode.Accept
            ? "ThYe tenant shall promptly repaint" : "The tenant shall promptly repaint", Rejected(session));
    }

    [Theory]
    [InlineData(TrackedChangeMode.Accept, DeletedTextSide.Before)]
    [InlineData(TrackedChangeMode.Accept, DeletedTextSide.After)]
    [InlineData(TrackedChangeMode.RenderInline, DeletedTextSide.Before)]
    [InlineData(TrackedChangeMode.RenderInline, DeletedTextSide.After)]
    public void DeletedTextSide_BesideAMoveSource_KeepsNewTextOutsideItsRange(
        TrackedChangeMode tracking, DeletedTextSide side)
    {
        var (session, p) = Open(tracking, "Pat", Run("Fee "),
            new XElement(W.moveFromRangeStart, Stamp("Alex", 10), new XAttribute(W.name, "move1")),
            new XElement(W.moveFrom, Stamp("Alex", 11), new XElement(W.r, new XElement(W.delText, "amount"))),
            new XElement(W.moveFromRangeEnd, new XAttribute(W.id, 10)), Run("due "),
            new XElement(W.moveToRangeStart, Stamp("Alex", 12), new XAttribute(W.name, "move1")),
            new XElement(W.moveTo, Stamp("Alex", 13), Run("amount")),
            new XElement(W.moveToRangeEnd, new XAttribute(W.id, 12)));

        var edit = session.ReplaceTextAtSpan(p, 4, 0, "XY", side);

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("Fee XYdue amount", Accepted(session));
        Assert.Equal(tracking == TrackedChangeMode.RenderInline ? "Fee amountdue "
            : side == DeletedTextSide.Before ? "Fee XYamountdue " : "Fee amountXYdue ", Rejected(session));
    }

    [Theory]
    [InlineData(TrackedChangeMode.Accept, false)]
    [InlineData(TrackedChangeMode.Accept, true)]
    [InlineData(TrackedChangeMode.RenderInline, false)]
    [InlineData(TrackedChangeMode.RenderInline, true)]
    public void DeletedTextSide_OutsideAMoveSourceParagraph_IsRefusedWithoutMutation(
        TrackedChangeMode tracking, bool after)
    {
        using var stream = new MemoryStream();
        stream.Write(Build(Run("Fee "),
            new XElement(W.moveFromRangeStart, Stamp("Alex", 10), new XAttribute(W.name, "move1")),
            new XElement(W.moveFrom, Stamp("Alex", 11), new XElement(W.r, new XElement(W.delText, "amount")))));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var xml = document.MainDocumentPart!.GetXDocument();
            xml.Root!.Element(W.body)!.Add(
                new XElement(W.p,
                    new XElement(W.moveFrom, Stamp("Alex", 14), new XElement(W.r, new XElement(W.delText, " more"))),
                    new XElement(W.moveFromRangeEnd, new XAttribute(W.id, 10)), Run("due")),
                new XElement(W.p,
                    new XElement(W.moveToRangeStart, Stamp("Alex", 12), new XAttribute(W.name, "move1")),
                    new XElement(W.moveTo, Stamp("Alex", 13), Run("amount more")),
                    new XElement(W.moveToRangeEnd, new XAttribute(W.id, 12))));
            document.MainDocumentPart.PutXDocument(xml);
        }
        using var session = new DocxSession(stream.ToArray(), new DocxSessionSettings
        {
            TrackedChanges = tracking, RevisionAuthor = "Pat",
        });
        var p = session.FindByText(after ? "Fee" : "due")!.Anchor.Id;
        var before = session.Raw.GetXml(p);
        var version = session.Version;

        var edit = session.ReplaceTextAtSpan(p, after ? 4 : 0, 0, "XY",
            after ? DeletedTextSide.After : DeletedTextSide.Before);

        Assert.False(edit.Success);
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary, edit.Error?.Code);
        Assert.Equal(before, session.Raw.GetXml(p));
        Assert.Equal(version, session.Version);
    }

    // ─── Helpers ─────────────────────────────────────────────────────────

    internal static XElement Run(string text, bool bold = false) => new(W.r,
        bold ? new XElement(W.rPr, new XElement(W.b)) : null,
        new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), text));

    internal static object[] Stamp(string author, int id) => new object[]
    {
        new XAttribute(W.id, id), new XAttribute(W.author, author), new XAttribute(W.date, Date),
    };

    private static XElement Ins(string author, int id, params object[] content) =>
        new(W.ins, Stamp(author, id), content);

    internal static XElement Del(string author, int id, string text) => new(W.del, Stamp(author, id),
        new XElement(W.r, new XElement(W.delText, new XAttribute(XNamespace.Xml + "space", "preserve"), text)));

    private static XElement FieldChar(string type) =>
        new(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, type)));

    private static XElement Bookmark(XName marker) => marker == W.bookmarkStart
        ? new XElement(marker, new XAttribute(W.id, 10), new XAttribute(W.name, "_Toc1"))
        : new XElement(marker, new XAttribute(W.id, 10));

    private static (DocxSession Session, string Anchor) Open(
        TrackedChangeMode tracking, string author, params object[] paragraphContent)
    {
        var session = new DocxSession(Build(paragraphContent), new DocxSessionSettings
        {
            TrackedChanges = tracking,
            RevisionAuthor = author,
        });
        return (session, session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id);
    }

    /// <summary>A document whose body is one paragraph holding <paramref name="paragraphContent"/>.</summary>
    internal static byte[] Build(params object[] paragraphContent)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
            main.PutXDocument(new XDocument(new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XElement(W.body, new XElement(W.p, paragraphContent)))));
        }
        return stream.ToArray();
    }

    /// <summary>The paragraph's content as runs and revision envelopes, ignoring properties,
    /// ids and dates: <c>"Fee " ins:Alex["amo"]</c>.</summary>
    private static string Shape(DocxSession session, string anchor) =>
        Shape(XElement.Parse(session.Raw.GetXml(anchor)));

    private static string Shape(XElement container) => string.Join(" ", container.Elements()
        .Where(e => e.Name != W.pPr && e.Name != W.rPr)
        .Select(e => e.Name == W.r
            ? "\"" + string.Concat(e.Elements().Select(t => t.Name == W.t || t.Name == W.delText ? t.Value
                : t.Name == W.tab ? "⇥" : t.Name == W.br ? "↵" : "")) + "\""
            : $"{e.Name.LocalName}{(e.Attribute(W.author) is { } a ? ":" + a.Value : "")}[{Shape(e)}]"));

    private static string[] Revisions(DocxSession session) =>
        session.ListRevisions().Select(r => $"{r.Type}:{r.Author}:{r.Text}").ToArray();

    private static string Accepted(DocxSession session) => ParagraphText(
        RevisionProcessor.AcceptRevisions(new WmlDocument("a.docx", session.Save())).DocumentByteArray);

    private static string Rejected(DocxSession session) => ParagraphText(
        RevisionProcessor.RejectRevisions(new WmlDocument("r.docx", session.Save())).DocumentByteArray);

    private static string ParagraphText(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return string.Concat(document.MainDocumentPart!.GetXDocument().Descendants(W.t).Select(t => t.Value));
    }
}
