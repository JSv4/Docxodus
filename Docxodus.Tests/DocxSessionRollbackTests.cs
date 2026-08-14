#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Rollback-on-internal-error tests for <see cref="DocxSession"/> (DS42x).
///
/// <para>Every mutation records a pre-op snapshot before entering its <c>try</c>, so a throw from
/// inside can leave the document HALF-MUTATED. The contract these tests pin is the one the typed
/// <see cref="EditResult"/> envelope implies but never used to guarantee: <b>a failed op leaves the
/// document byte-identical to what it was before the call, and leaves nothing on the undo ring</b>.</para>
///
/// <para><b>Why a NUL character is the trigger.</b> These are not synthetic faults. Markdown payloads
/// reach <see cref="DocxSession"/> from LLM output, clipboard paste, and scraped text — all of which
/// routinely carry a stray <c>U+0000</c> or an unpaired surrogate. XML cannot represent either, so
/// the write throws <see cref="System.ArgumentException"/> from deep inside the op, well after it has
/// started mutating. <see cref="DocxSession.InsertFootnote"/> is the sharpest case: it creates the
/// FootnotesPart, its two Word-reserved notes, the <c>w:footnotePr</c> settings declaration and the
/// FootnoteText/FootnoteReference styles, and only THEN writes the note body that throws.</para>
/// </summary>
public class DocxSessionRollbackTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>A payload XML cannot encode — the realistic "stray control character" case.</summary>
    private const string NulPayload = "note\0text";

    /// <summary>An unpaired high surrogate: the other shape of the same class of bad input.</summary>
    private const string LoneSurrogatePayload = "before\ud800after";

    private static string FirstBodyParagraph(DocxSession s) =>
        s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    /// <summary>The whole package as a comparable tree: part URI → root element, so a rollback that
    /// restores the body but leaks a newly created part still fails the comparison.</summary>
    private static (string Uri, XElement Root)[] PackageParts(byte[] docxBytes)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var parts = main.Parts.Select(p => p.OpenXmlPart).Prepend((OpenXmlPart)main);

        var snapshot = new System.Collections.Generic.List<(string, XElement)>();
        foreach (var part in parts)
        {
            // Only XML parts participate: an image/font part has no comparable tree, and its
            // presence/absence is already covered by the part-URI list comparison below.
            if (!part.ContentType.Contains("xml", System.StringComparison.OrdinalIgnoreCase)) continue;
            using var stream = part.GetStream();
            snapshot.Add((part.Uri.ToString(), XDocument.Load(stream).Root!));
        }
        return snapshot.OrderBy(t => t.Item1, System.StringComparer.Ordinal).ToArray();
    }

    private static void AssertPackageUnchanged(
        (string Uri, XElement Root)[] before, (string Uri, XElement Root)[] after)
    {
        Assert.Equal(
            before.Select(p => p.Uri).ToArray(),
            after.Select(p => p.Uri).ToArray());
        for (int i = 0; i < before.Length; i++)
        {
            Assert.True(
                XNode.DeepEquals(before[i].Root, after[i].Root),
                $"part '{before[i].Uri}' differs after a FAILED operation");
        }
    }

    /// <summary>
    /// The core contract, exercised through the op with the largest partial-mutation blast radius.
    /// Before the fix this op discarded its snapshot without applying it, so the footnotes scaffold
    /// it had already built survived the failure AND the record that could have reversed it was
    /// thrown away — a permanent, un-undoable mutation reported to the caller as a clean failure.
    /// </summary>
    [Fact]
    public void DS420_InsertFootnote_ThrowsMidOp_RollsBackEveryPart()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(s);
        var before = PackageParts(s.Save());

        var result = s.InsertFootnote(anchor, 0, NulPayload);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.InternalError, result.Error!.Code);
        Assert.NotNull(s.LastInternalError);
        Assert.Null(s.LastRollbackError);

        AssertPackageUnchanged(before, PackageParts(s.Save()));
    }

    /// <summary>The scaffold the failed op built must not survive: no FootnotesPart, and no
    /// <c>w:footnotePr</c> left behind in the settings part.</summary>
    [Fact]
    public void DS421_InsertFootnote_ThrowsMidOp_LeavesNoOrphanedScaffold()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(s);

        Assert.False(s.InsertFootnote(anchor, 0, NulPayload).Success);

        using var ms = new MemoryStream(s.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;

        Assert.Null(main.FootnotesPart);
        var settings = main.DocumentSettingsPart?.GetXDocument().Root;
        Assert.Null(settings?.Element(W + "footnotePr"));
        Assert.Empty(main.GetXDocument().Descendants(W + "footnoteReference"));
    }

    /// <summary>A failed op must leave the undo ring exactly as it found it. Otherwise the next
    /// <see cref="DocxSession.Undo"/> reverts the previous SUCCESSFUL edit while the caller believes
    /// it is reverting the failed one.</summary>
    [Fact]
    public void DS422_FailedOp_DoesNotConsumeOrPolluteTheUndoRing()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(s);

        // No history yet: a failed op must not leave a record behind.
        Assert.False(s.InsertFootnote(anchor, 0, NulPayload).Success);
        Assert.False(s.Undo());

        // One real edit, then a failure, then undo: the undo must reverse the REAL edit.
        var pristine = PackageParts(s.Save());
        Assert.True(s.ReplaceText(anchor, "changed").Success);
        var afterEdit = PackageParts(s.Save());

        Assert.False(s.InsertFootnote(anchor, 0, NulPayload).Success);
        AssertPackageUnchanged(afterEdit, PackageParts(s.Save()));

        Assert.True(s.Undo());
        AssertPackageUnchanged(pristine, PackageParts(s.Save()));
        Assert.False(s.Undo());
    }

    /// <summary><see cref="DocxSession.SetHeaderText"/> creates a HeaderPart, its relationship and a
    /// <c>w:headerReference</c> in the section before writing the story that throws. All three must
    /// be reconciled away — a leaked reference to a deleted part is a document Word refuses to open.
    /// </summary>
    [Fact]
    public void DS423_SetHeaderText_ThrowsMidOp_RollsBackPartAndReference()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(s);
        var before = PackageParts(s.Save());

        var result = s.SetHeaderText(anchor, HeaderFooterKind.Default, NulPayload);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.InternalError, result.Error!.Code);
        Assert.Null(s.LastRollbackError);
        AssertPackageUnchanged(before, PackageParts(s.Save()));

        using var ms = new MemoryStream(s.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Empty(doc.MainDocumentPart!.HeaderParts);
        Assert.Empty(doc.MainDocumentPart.GetXDocument().Descendants(W + "headerReference"));
    }

    /// <summary>Same contract for the comments part, which <see cref="DocxSession.AddComment"/>
    /// find-or-creates (along with the CommentText/CommentReference styles) before writing the body.
    /// </summary>
    [Fact]
    public void DS424_AddComment_ThrowsMidOp_RollsBackCommentsPart()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(s);
        var before = PackageParts(s.Save());

        var result = s.AddComment(anchor, new CharSpan(0, 5), "Reviewer", NulPayload);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.InternalError, result.Error!.Code);
        Assert.Null(s.LastRollbackError);
        AssertPackageUnchanged(before, PackageParts(s.Save()));

        using var ms = new MemoryStream(s.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Null(doc.MainDocumentPart!.WordprocessingCommentsPart);
        Assert.Empty(doc.MainDocumentPart.GetXDocument().Descendants(W + "commentRangeStart"));
    }

    /// <summary>The unpaired-surrogate shape of the same bad input, through a body-text op whose
    /// mutation is confined to the main part.</summary>
    [Fact]
    public void DS425_ReplaceText_ThrowsMidOp_RollsBackBody()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(s);
        var before = PackageParts(s.Save());

        var result = s.ReplaceText(anchor, LoneSurrogatePayload);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.InternalError, result.Error!.Code);
        Assert.Null(s.LastRollbackError);
        AssertPackageUnchanged(before, PackageParts(s.Save()));
        Assert.False(s.Undo());
    }

    /// <summary>
    /// Snapshot-scope regression, found while building the rollback tests above and NOT limited to
    /// the error path: <c>settings.xml</c> and <c>styles.xml</c> were outside the snapshot, so an
    /// ordinary <see cref="DocxSession.Undo"/> of a SUCCESSFUL first-footnote insert left the
    /// <c>w:footnotePr</c> settings declaration and the two generated note styles behind forever.
    /// Snapshot membership follows what ops WRITE, not what the projector reads.
    /// </summary>
    [Fact]
    public void DS427_Undo_RevertsSettingsAndStylesWrittenByAnOp()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(s);
        var before = PackageParts(s.Save());

        Assert.True(s.InsertFootnote(anchor, 0, "a real footnote").Success);
        Assert.True(s.Undo());

        AssertPackageUnchanged(before, PackageParts(s.Save()));
    }

    /// <summary>The same for a SUCCESSFUL header authoring call, which writes on both sides of the
    /// old snapshot boundary at once: a new HeaderPart (already reconciled) AND the
    /// <c>w:titlePg</c> flag in settings (previously not). Undo must reverse both, or the document
    /// keeps a first-page-header flag pointing at a story that no longer exists.</summary>
    [Fact]
    public void DS428_Undo_RevertsBothANewPartAndItsSettingsFlag()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(s);
        var before = PackageParts(s.Save());

        var set = s.SetHeaderText(anchor, HeaderFooterKind.First, "First-page header");
        Assert.True(set.Success, $"{set.Error?.Code}: {set.Error?.Message}");
        Assert.True(s.Undo());

        AssertPackageUnchanged(before, PackageParts(s.Save()));
    }

    /// <summary>A session that survived a rolled-back failure is still fully usable — the rollback
    /// restores working state, not just bytes on disk. Guards against a restore that leaves a stale
    /// projection/anchor cache pointing at replaced XElements.</summary>
    [Fact]
    public void DS426_SessionRemainsUsableAfterRollback()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(s);

        Assert.False(s.InsertFootnote(anchor, 0, NulPayload).Success);

        // The same anchor still resolves, and a well-formed footnote now succeeds.
        Assert.True(s.Exists(anchor));
        var ok = s.InsertFootnote(anchor, 0, "a real footnote");
        Assert.True(ok.Success, ok.Error?.Message);

        using var ms = new MemoryStream(s.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.NotNull(doc.MainDocumentPart!.FootnotesPart);
        Assert.Single(doc.MainDocumentPart.GetXDocument().Descendants(W + "footnoteReference"));
    }
}
