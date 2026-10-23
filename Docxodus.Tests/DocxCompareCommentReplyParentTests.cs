// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1312: a revision that replies to one of the original's comments keeps the reply under that comment in
/// the redline. The redline's comment definitions come from the original, but a reply added in the revision names
/// its parent by the revision's own <c>w14:paraId</c> for that comment, which the original may lack (a tool that
/// writes no paraIds) or number differently.
/// </summary>
public class DocxCompareCommentReplyParentTests
{
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    private static readonly XNamespace W16Cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";

    private const string Clause = "The seller delivers the goods.";
    private const string Other = "The buyer pays on delivery.";
    private const string Root = "Root comment.";
    private const string Reply = "First reply.";

    private static void Write(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(xml);
    }

    /// <summary>A document with the given body and comment definitions, and a <c>commentsExtended</c> part only
    /// when <paramref name="commentsEx"/> is given.</summary>
    private static byte[] Build(string body, string comments, string? commentsEx = null)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            var w14 = comments.Contains("w14:", StringComparison.Ordinal)
                ? $" xmlns:w14=\"{W14}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"w14\"" : "";
            Write(main.AddNewPart<WordprocessingCommentsPart>(), $"<w:comments xmlns:w=\"{W.w.NamespaceName}\"{w14}>{comments}</w:comments>");
            if (commentsEx != null)
                Write(main.AddNewPart<WordprocessingCommentsExPart>(), $"<w15:commentsEx xmlns:w15=\"{W15}\">{commentsEx}</w15:commentsEx>");
            Write(main, $"<w:document xmlns:w=\"{W.w.NamespaceName}\"><w:body>{body}</w:body></w:document>");
        }
        return ms.ToArray();
    }

    private static string Commented(string text, params int[] ids) =>
        "<w:p>" + string.Concat(ids.Select(id => $"<w:commentRangeStart w:id=\"{id}\"/>"))
        + $"<w:r><w:t>{text}</w:t></w:r>"
        + string.Concat(ids.Select(id => $"<w:commentRangeEnd w:id=\"{id}\"/><w:r><w:commentReference w:id=\"{id}\"/></w:r>"))
        + "</w:p>";

    private static string Comment(int id, string author, string text, string? paraId = null) =>
        $"<w:comment w:id=\"{id}\" w:author=\"{author}\"><w:p{(paraId == null ? "" : $" w14:paraId=\"{paraId}\"")}><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>";

    private static string Ex(string paraId, string? parent = null) =>
        $"<w15:commentEx w15:paraId=\"{paraId}\"{(parent == null ? "" : $" w15:paraIdParent=\"{parent}\"")} w15:done=\"0\"/>";

    /// <summary>The original most cases start from: one clause carrying one root comment with no paraId, and no
    /// <c>commentsExtended</c> part.</summary>
    private static byte[] IdLessRoot() => Build(Commented(Clause, 0), Comment(0, "Ann", Root));

    /// <summary>A thread as a paraId-writing editor saves it: the root numbered <paramref name="rootParaId"/> and
    /// one reply under it.</summary>
    private static byte[] Thread(string clause, string rootText, string rootParaId = "00000001", string replyParaId = "00000002") =>
        Build(Commented(clause, 0, 1), Comment(0, "Ann", rootText, rootParaId) + Comment(1, "Ben", Reply, replyParaId),
            Ex(rootParaId) + Ex(replyParaId, rootParaId));

    /// <summary><paramref name="original"/> with a session-authored reply from Ben to each named comment.</summary>
    private static byte[] WithReplies(byte[] original, params (string Parent, string Reply)[] replies)
    {
        using var session = new DocxSession(original);
        foreach (var (parent, reply) in replies)
        {
            var root = session.ListComments().Single(c => c.Text == parent).DefAnchorId;
            Assert.True(session.AddCommentReply(root, "Ben", reply).Success);
        }
        return session.Save();
    }

    private static WmlDocument Doc(byte[] bytes) => new("doc.docx", bytes);

    /// <summary>Each comment's text, mapped to its parent's text (null for a top-level comment), through
    /// <see cref="DocxSession.ListComments"/>.</summary>
    private static Dictionary<string, string?> Threads(byte[] document)
    {
        using var session = new DocxSession(document);
        var comments = session.ListComments();
        var textByAnchor = comments.ToDictionary(c => c.DefAnchorId, c => c.Text);
        return comments.ToDictionary(c => c.Text, c => c.ParentAnchorId is { } parent ? textByAnchor[parent] : null);
    }

    /// <summary>The comment parts agree with each other: every threading entry names a definition, every parent
    /// has its own <c>commentsExtended</c> entry (and <c>commentsIds</c> entry when that part exists), no two
    /// definitions share a paraId, and the package validates.</summary>
    private static void AssertSelfConsistent(byte[] document)
    {
        PackageValidation.AssertValid(document, FileFormatVersions.Microsoft365);
        using var stream = new MemoryStream(document);
        using var package = WordprocessingDocument.Open(stream, false);
        var main = package.MainDocumentPart!;
        var definitions = main.WordprocessingCommentsPart!.GetXDocument().Root!.Elements(W.comment)
            .Select(c => (string?)c.Elements(W.p).Last().Attribute(W14 + "paraId")).OfType<string>().ToList();
        Assert.Equal(definitions.Count, definitions.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        var entries = main.WordprocessingCommentsExPart?.GetXDocument().Root!.Elements(W15 + "commentEx").ToList() ?? new();
        var keyed = entries.Select(e => (string)e.Attribute(W15 + "paraId")!).ToList();
        Assert.All(keyed, key => Assert.Contains(key, definitions));
        Assert.All(entries.Select(e => (string?)e.Attribute(W15 + "paraIdParent")).OfType<string>(),
            parent => Assert.Contains(parent, keyed));
        if (main.WordprocessingCommentsIdsPart?.GetXDocument().Root is { } ids)
            Assert.All(keyed, key => Assert.Single(ids.Elements(W16Cid + "commentId"),
                e => (string?)e.Attribute(W16Cid + "paraId") == key));
    }

    /// <summary>The redline threads its comments as <paramref name="expected"/>, its comment parts are
    /// self-consistent, accepting it threads the revision's comments as the revision does, and rejecting it
    /// leaves self-consistent comment parts.</summary>
    private static void AssertRedline(byte[] redline, byte[] revised, Dictionary<string, string?> expected)
    {
        Assert.Equal(expected, Threads(redline));
        AssertSelfConsistent(redline);
        var accepted = RevisionProcessor.AcceptRevisions(Doc(redline)).DocumentByteArray;
        AssertSelfConsistent(accepted);
        // Accepting leaves a deleted comment's definition behind, so compare the revision's comments only.
        var acceptedThreads = Threads(accepted);
        Assert.All(Threads(revised), thread => Assert.Equal(thread.Value, acceptedThreads[thread.Key]));
        AssertSelfConsistent(RevisionProcessor.RejectRevisions(Doc(redline)).DocumentByteArray);
    }

    private static byte[] Compare(byte[] original, byte[] revised) =>
        DocxCompare.Compare(Doc(original), Doc(revised)).DocumentByteArray;

    private static readonly Dictionary<string, string?> RootAndReply = new() { [Root] = null, [Reply] = Root };

    [Fact]
    public void ReplyToARootWithoutParaId_ListsUnderItsRoot()
    {
        var original = IdLessRoot();
        var revised = WithReplies(original, (Root, Reply));

        AssertRedline(Compare(original, revised), revised, RootAndReply);
    }

    /// <summary>The original's root already carries the paraId the revision's reply names, so the reply listed
    /// under it before #1312 too; the root's <c>commentsIds</c> entry, which the revision's part holds, was
    /// missing.</summary>
    [Fact]
    public void ReplyToARootWithParaId_ListsUnderItsRoot()
    {
        var original = Build(Commented(Clause, 0), Comment(0, "Ann", Root, "00000001"), Ex("00000001"));
        var revised = WithReplies(original, (Root, Reply));

        AssertRedline(Compare(original, revised), revised, RootAndReply);
    }

    /// <summary>The revision numbered the root's paragraph differently from the original (an editor that rewrites
    /// paraIds on save), so the paraId the reply names is not in the original.</summary>
    [Fact]
    public void ReplyToARootTheRevisionRenumbered_ListsUnderItsRoot()
    {
        var original = Build(Commented(Clause, 0), Comment(0, "Ann", Root, "0000000A"), Ex("0000000A"));
        var revised = Thread(Clause, Root, "0000000B", "0000000C");

        AssertRedline(Compare(original, revised), revised, RootAndReply);
    }

    /// <summary>The paraId the revision gave the id-less root is, in the original, another comment's. The reply
    /// must not thread under that other comment, and that comment's own entry is left alone.</summary>
    [Fact]
    public void ReplyWhoseParentParaIdTheOriginalGivesAnotherComment_ListsUnderItsOwnRoot()
    {
        var original = Build(Commented(Clause, 0) + Commented(Other, 5),
            Comment(0, "Ann", Root) + Comment(5, "Cy", "Unrelated note.", "00000001"), Ex("00000001"));
        // The revision renumbered the note and gave the root the note's old paraId; untranslated, the reply would
        // thread under the note.
        var revised = Build(Commented(Clause, 0, 1) + Commented(Other, 5),
            Comment(0, "Ann", Root, "00000001") + Comment(1, "Ben", Reply, "00000002") + Comment(5, "Cy", "Unrelated note.", "0000000A"),
            Ex("0000000A") + Ex("00000001") + Ex("00000002", "00000001"));

        AssertRedline(Compare(original, revised), revised,
            new() { [Root] = null, [Reply] = Root, ["Unrelated note."] = null });
    }

    /// <summary>Two id-less roots each get a reply, and one reply gets a reply of its own.</summary>
    [Fact]
    public void RepliesToSeveralRootsWithoutParaIds_EachListUnderTheirOwnRoot()
    {
        var original = Build(Commented(Clause, 0) + Commented(Other, 1),
            Comment(0, "Ann", "First root.") + Comment(1, "Ann", "Second root."));
        var revised = WithReplies(original, ("First root.", "Reply to first."), ("Second root.", "Reply to second."),
            ("Reply to first.", "Nested reply."));

        AssertRedline(Compare(original, revised), revised, new()
        {
            ["First root."] = null,
            ["Second root."] = null,
            ["Reply to first."] = "First root.",
            ["Reply to second."] = "Second root.",
            ["Nested reply."] = "Reply to first.",
        });
    }

    /// <summary>The revision replies to an original comment that it also threads under another original comment
    /// (the original wrote no threading at all). The reply's parent entry is carried from the revision, and so
    /// is that entry's own link, so the chain resolves end to end as it does once accepted.</summary>
    [Fact]
    public void ReplyToAnOriginalCommentTheRevisionThreads_ListsTheWholeChain()
    {
        var original = Build(Commented(Clause, 0, 1), Comment(0, "Ann", Root) + Comment(1, "Ben", Reply));
        var revised = Build(Commented(Clause, 0, 1, 2),
            Comment(0, "Ann", Root, "00000001") + Comment(1, "Ben", Reply, "00000002") + Comment(2, "Cy", "Second reply.", "00000003"),
            Ex("00000001") + Ex("00000002", "00000001") + Ex("00000003", "00000002"));

        AssertRedline(Compare(original, revised), revised,
            new() { [Root] = null, [Reply] = Root, ["Second reply."] = Reply });
    }

    /// <summary>The revision replies to an id-less root and rewrites the commented clause, so the root rides a
    /// deleted and an inserted copy; the reply follows the copy that survives accept.</summary>
    [Fact]
    public void ReplyToARootWithoutParaIdWhoseClauseIsEdited_ListsUnderItsRoot()
    {
        var original = IdLessRoot();
        var revised = Thread("The seller ships the goods by Friday.", Root);

        var redline = Compare(original, revised);

        Assert.Equal(Root, Threads(redline)[Reply]);
        AssertSelfConsistent(redline);
        var accepted = RevisionProcessor.AcceptRevisions(Doc(redline)).DocumentByteArray;
        Assert.Equal(Root, Threads(accepted)[Reply]);
    }

    /// <summary>Control: a reworded root is a different comment (comments are matched by author, date and text),
    /// so the redline deletes the original's and inserts the revision's, and the reply lists under the inserted
    /// one.</summary>
    [Fact]
    public void ReplyToARootTheRevisionRewords_ListsUnderTheRewordedRoot()
    {
        var original = IdLessRoot();
        var revised = Thread(Clause, "Root comment, reworded.");

        AssertRedline(Compare(original, revised), revised,
            new() { [Root] = null, ["Root comment, reworded."] = null, [Reply] = "Root comment, reworded." });
    }

    /// <summary>Control: the revision deletes the original's reply; the redline keeps it threaded under its root
    /// on the deleted side.</summary>
    [Fact]
    public void ReplyDeletedInTheRevision_StillListsUnderItsRoot()
    {
        var original = Thread(Clause, Root);
        var revised = Build(Commented(Clause, 0), Comment(0, "Ann", Root, "00000001"), Ex("00000001"));

        AssertRedline(Compare(original, revised), revised, RootAndReply);
    }

    /// <summary>Control: the revision deletes the commented clause with its whole thread.</summary>
    [Fact]
    public void RootDeletedWithItsReply_KeepsTheThreadOnTheDeletedSide()
    {
        var original = Build(Commented(Clause, 0, 1) + Commented(Other),
            Comment(0, "Ann", Root, "00000001") + Comment(1, "Ben", Reply, "00000002"), Ex("00000001") + Ex("00000002", "00000001"));
        var revised = Build(Commented(Other), "");

        AssertRedline(Compare(original, revised), revised, RootAndReply);
    }

    /// <summary>Control: a comment the revision adds without any paraId (a tool that writes none) cannot be
    /// threaded; it lists top-level and the redline invents no threading for it or the root.</summary>
    [Fact]
    public void CommentAddedWithoutParaId_ListsTopLevelWithoutThreadingParts()
    {
        var original = IdLessRoot();
        var revised = Build(Commented(Clause, 0, 1), Comment(0, "Ann", Root) + Comment(1, "Ben", Reply));

        var redline = Compare(original, revised);

        Assert.Equal(new Dictionary<string, string?> { [Root] = null, [Reply] = null }, Threads(redline));
        AssertSelfConsistent(redline);
        using var stream = new MemoryStream(redline);
        using var package = WordprocessingDocument.Open(stream, false);
        Assert.Null(package.MainDocumentPart!.WordprocessingCommentsExPart);
    }

    /// <summary>The revision's reply answers a comment the redline does not hold (an unanchored definition in the
    /// revision). The reply lists top-level instead of naming a parent that is not there.</summary>
    [Fact]
    public void ReplyToACommentTheRedlineDoesNotHold_ListsTopLevel()
    {
        var original = Build(Commented(Clause), "");
        var revised = Build(Commented(Clause, 1),
            Comment(0, "Ann", Root, "00000001") + Comment(1, "Ben", Reply, "00000002"), Ex("00000001") + Ex("00000002", "00000001"));

        var redline = Compare(original, revised);

        Assert.Equal(new Dictionary<string, string?> { [Reply] = null }, Threads(redline));
        AssertSelfConsistent(redline);
    }

    /// <summary>The consolidate path merges each reviewer's comments the same way; a second reviewer's reply
    /// finds the paraId the first reviewer's merge gave the root.</summary>
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void ConsolidatedRepliesToARootWithoutParaId_ListUnderItsRoot(int reviewers)
    {
        var original = IdLessRoot();
        var reviews = Enumerable.Range(1, reviewers)
            .Select(n => new DocxDiffReviewer { Document = Doc(WithReplies(original, (Root, $"Reply {n}."))), Author = $"R{n}" })
            .ToList();

        var redline = DocxCompare.Consolidate(Doc(original), reviews).DocumentByteArray;

        var expected = new Dictionary<string, string?> { [Root] = null };
        for (int n = 1; n <= reviewers; n++)
            expected[$"Reply {n}."] = Root;
        Assert.Equal(expected, Threads(redline));
        AssertSelfConsistent(redline);
    }
}
