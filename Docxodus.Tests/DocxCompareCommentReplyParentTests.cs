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
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

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

    /// <summary>A document with the given body and comment definitions, and <c>commentsExtended</c> and
    /// <c>commentsIds</c> parts only when given. A comments part that uses <c>w14</c> declares it, listed in
    /// <c>mc:Ignorable</c> unless <paramref name="ignorable"/> is false.</summary>
    private static byte[] Build(string body, string comments, string? commentsEx = null, string? commentsIds = null,
        bool ignorable = true)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            var w14 = !comments.Contains("w14:", StringComparison.Ordinal) ? ""
                : $" xmlns:w14=\"{W14}\"" + (ignorable ? $" xmlns:mc=\"{Mc}\" mc:Ignorable=\"w14\"" : "");
            Write(main.AddNewPart<WordprocessingCommentsPart>(), $"<w:comments xmlns:w=\"{W.w.NamespaceName}\"{w14}>{comments}</w:comments>");
            if (commentsEx != null)
                Write(main.AddNewPart<WordprocessingCommentsExPart>(), $"<w15:commentsEx xmlns:w15=\"{W15}\">{commentsEx}</w15:commentsEx>");
            if (commentsIds != null)
                Write(main.AddNewPart<WordprocessingCommentsIdsPart>(), $"<w16cid:commentsIds xmlns:w16cid=\"{W16Cid}\">{commentsIds}</w16cid:commentsIds>");
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

    private static string Ex(string paraId, string? parent = null, bool done = false) =>
        $"<w15:commentEx w15:paraId=\"{paraId}\"{(parent == null ? "" : $" w15:paraIdParent=\"{parent}\"")} w15:done=\"{(done ? 1 : 0)}\"/>";

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

    /// <summary>Each comment's text, mapped to whether it is resolved (null when it has no threading entry).</summary>
    private static Dictionary<string, bool?> Resolved(byte[] document)
    {
        using var session = new DocxSession(document);
        return session.ListComments().ToDictionary(c => c.Text, c => c.Resolved);
    }

    /// <summary>The comment parts agree with each other: every threading entry names a definition, every
    /// definition keyed by a paraId has exactly one entry in each threading part the package has, every parent
    /// link names an entry, no two definitions share a paraId, and the package validates.</summary>
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
        if (main.WordprocessingCommentsExPart != null)
            Assert.All(definitions, definition => Assert.Single(keyed, key => key == definition));
        if (main.WordprocessingCommentsIdsPart?.GetXDocument().Root is { } ids)
            Assert.All(definitions, definition => Assert.Single(ids.Elements(W16Cid + "commentId"),
                e => (string?)e.Attribute(W16Cid + "paraId") == definition));
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

    /// <summary>The paraId the revision gave the id-less root is, in the original, a resolved note's. Adopting the
    /// revision's paraId for the root would thread the reply under the note, or take over the note's entry and its
    /// resolved state; the note keeps both.</summary>
    [Fact]
    public void ReplyWhoseParentParaIdTheOriginalGivesAnotherComment_ListsUnderItsOwnRoot()
    {
        var original = Build(Commented(Clause, 0) + Commented(Other, 5),
            Comment(0, "Ann", Root) + Comment(5, "Cy", "Unrelated note.", "00000001"), Ex("00000001", done: true));
        // The revision renumbered the note and gave the root the note's old paraId.
        var revised = Build(Commented(Clause, 0, 1) + Commented(Other, 5),
            Comment(0, "Ann", Root, "00000001") + Comment(1, "Ben", Reply, "00000002") + Comment(5, "Cy", "Unrelated note.", "0000000A"),
            Ex("0000000A", done: true) + Ex("00000001") + Ex("00000002", "00000001"));

        var redline = Compare(original, revised);

        AssertRedline(redline, revised, new() { [Root] = null, [Reply] = Root, ["Unrelated note."] = null });
        Assert.Equal(new Dictionary<string, bool?> { [Root] = false, [Reply] = false, ["Unrelated note."] = true }, Resolved(redline));
        using var stream = new MemoryStream(redline);
        using var package = WordprocessingDocument.Open(stream, false);
        var note = package.MainDocumentPart!.WordprocessingCommentsPart!.GetXDocument().Root!.Elements(W.comment)
            .Single(c => (string?)c.Attribute(W.id) == "5");
        Assert.Equal("00000001", (string?)note.Element(W.p)!.Attribute(W14 + "paraId"));
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

    /// <summary>The revision threads one original comment under another and replies to it. Threading is not a
    /// tracked change and the redline's definitions are the original's, so the original's two comments keep the
    /// original's (flat) threading, with or without entries of their own; only the added reply is linked.</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReplyToAnOriginalCommentTheRevisionRethreads_KeepsTheOriginalsThreading(bool originalEntries)
    {
        var original = originalEntries
            ? Build(Commented(Clause, 0, 1), Comment(0, "Ann", Root, "0000000A") + Comment(1, "Ben", Reply, "0000000B"), Ex("0000000A") + Ex("0000000B"))
            : Build(Commented(Clause, 0, 1), Comment(0, "Ann", Root) + Comment(1, "Ben", Reply));
        var revised = Build(Commented(Clause, 0, 1, 2),
            Comment(0, "Ann", Root, "00000001") + Comment(1, "Ben", Reply, "00000002") + Comment(2, "Cy", "Second reply.", "00000003"),
            Ex("00000001") + Ex("00000002", "00000001") + Ex("00000003", "00000002"));

        var redline = Compare(original, revised);

        Assert.Equal(new Dictionary<string, string?> { [Root] = null, [Reply] = null, ["Second reply."] = Reply }, Threads(redline));
        AssertSelfConsistent(redline);
    }

    /// <summary>The revision resolves the root as it replies to it. Resolution is not a tracked change and the
    /// redline's root is the original's, which is open, whether the original wrote an entry for it or not, and
    /// whether or not a reply comes with it.</summary>
    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    public void RootTheRevisionResolves_StaysOpenInTheRedline(bool originalEntry, bool replied)
    {
        var original = originalEntry
            ? Build(Commented(Clause, 0), Comment(0, "Ann", Root, "0000000A"), Ex("0000000A"))
            : IdLessRoot();
        var revised = replied
            ? Build(Commented(Clause, 0, 1), Comment(0, "Ann", Root, "00000001") + Comment(1, "Ben", Reply, "00000002"),
                Ex("00000001", done: true) + Ex("00000002", "00000001", done: true))
            : Build(Commented(Clause, 0), Comment(0, "Ann", Root, "00000001"), Ex("00000001", done: true));

        var redline = Compare(original, revised);

        Assert.NotEqual(true, Resolved(redline)[Root]);
        if (replied)
            Assert.Equal(Root, Threads(redline)[Reply]);
        AssertSelfConsistent(redline);
    }

    /// <summary>Two identical comments (same author, date and text) in the original; the revision, which also
    /// renumbers its comments, replies to the second. The redline cannot tell which of the two the reply answers,
    /// so it lists the reply top-level rather than guess.</summary>
    [Fact]
    public void ReplyToOneOfTwoIdenticalComments_ListsTopLevel()
    {
        var original = Build(Commented(Clause, 0) + Commented(Other, 1),
            Comment(0, "Ann", "Please confirm.") + Comment(1, "Ann", "Please confirm."));
        var revised = Build(Commented("A new opening clause.", 0) + Commented(Clause, 1) + Commented(Other, 2, 3),
            Comment(0, "Cy", "New note.", "0A000001") + Comment(1, "Ann", "Please confirm.", "0A000002")
            + Comment(2, "Ann", "Please confirm.", "0A000003") + Comment(3, "Ben", Reply, "0A000004"),
            Ex("0A000001") + Ex("0A000002") + Ex("0A000003") + Ex("0A000004", "0A000003"));

        var redline = Compare(original, revised);

        using (var session = new DocxSession(redline))
            Assert.Null(session.ListComments().Single(c => c.Text == Reply).ParentAnchorId);
        AssertSelfConsistent(redline);
    }

    /// <summary>Added comments whose paraIds the original already uses take fresh ones. A fresh paraId never
    /// repeats one a later added comment holds, so each reply stays under its own parent.</summary>
    [Fact]
    public void AddedCommentsRenamedForCollidingParaIds_KeepTheirThreads()
    {
        var original = Build(Commented("One.", 0) + Commented("Two.", 9) + Commented("Three."),
            Comment(0, "Ann", "Left note.", "00000003") + Comment(9, "Ann", "Left other.", "00000004"), Ex("00000003") + Ex("00000004"));
        var revised = Build(Commented("One.", 1) + Commented("Two.", 2, 9) + Commented("Three.", 3, 4),
            Comment(1, "Bob", "A.", "00000003") + Comment(2, "Bob", "B.", "00000006") + Comment(3, "Bob", "C.", "00000004")
            + Comment(4, "Dee", "D.", "00000008") + Comment(9, "Ann", "Left other.", "0000000F"),
            Ex("00000003") + Ex("00000006") + Ex("00000004", "00000006") + Ex("00000008", "00000004") + Ex("0000000F"));

        var redline = Compare(original, revised);

        var threads = Threads(redline);
        Assert.Equal("B.", threads["C."]);
        Assert.Equal("C.", threads["D."]);
        Assert.Null(threads["B."]);
        AssertSelfConsistent(redline);
    }

    /// <summary>The original has a <c>commentsIds</c> part and the revision does not: the reply it adds still gets
    /// a durable id, as every comment in a package with that part does.</summary>
    [Fact]
    public void ReplyAddedWithoutCommentsIds_GetsADurableIdInTheOriginalsPart()
    {
        var original = Build(Commented(Clause, 0), Comment(0, "Ann", Root, "00000001"), Ex("00000001"),
            "<w16cid:commentId w16cid:paraId=\"00000001\" w16cid:durableId=\"11111111\"/>");
        var revised = Thread(Clause, Root, "0E000001", "0E000002");

        AssertRedline(Compare(original, revised), revised, RootAndReply);
    }

    /// <summary>A real comment-dense document, with session replies to four of its comments: every comment of the
    /// redline gets its own entry in each threading part, and each reply lists under its root.</summary>
    [Fact]
    public void RepliesAddedToADenseCommentedDocument_ListUnderTheirRoots()
    {
        var original = File.ReadAllBytes(Path.Combine("../../../../TestFiles", "DD", "DD002-DenseComments.docx"));
        List<string> roots;
        using (var session = new DocxSession(original))
            roots = session.ListComments().Where(c => c.ParentAnchorId == null).Take(4).Select(c => c.Text).ToList();
        var revised = WithReplies(original, roots.Select(r => (r, "Re " + r)).ToArray());

        var redline = Compare(original, revised);

        var threads = Threads(redline);
        Assert.All(roots, root => Assert.Equal(root, threads["Re " + root]));
        AssertSelfConsistent(redline);
    }

    /// <summary>The paraId minted for an id-less root is a Word 2010 extension: the comments part lists
    /// <c>w14</c> in <c>mc:Ignorable</c> even when no input does.</summary>
    [Fact]
    public void ParaIdMintedForARoot_IsDeclaredIgnorable()
    {
        var original = IdLessRoot();
        var revised = Build(Commented(Clause, 0, 1), Comment(0, "Ann", Root, "00000001") + Comment(1, "Ben", Reply, "00000002"),
            Ex("00000001") + Ex("00000002", "00000001"), ignorable: false);

        var redline = Compare(original, revised);

        Assert.Equal(Root, Threads(redline)[Reply]);
        using var stream = new MemoryStream(redline);
        using var package = WordprocessingDocument.Open(stream, false);
        var root = package.MainDocumentPart!.WordprocessingCommentsPart!.GetXDocument().Root!;
        var prefix = root.GetPrefixOfNamespace(W14);
        Assert.NotNull(prefix);
        Assert.Contains(prefix, ((string?)root.Attribute(Mc + "Ignorable") ?? "").Split(' '));
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

    /// <summary>Reviewers who renumber paraIds each reply, and one also resolves the root. The redline's root stays
    /// the original's open comment whatever order the reviewers come in.</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConsolidatedReviewerWhoResolvesTheRoot_LeavesItOpenInAnyOrder(bool resolverFirst)
    {
        var original = IdLessRoot();
        byte[] Review(int n, bool done) => Build(Commented(Clause, 0, 1),
            Comment(0, "Ann", Root, $"1{n}000001") + Comment(1, $"R{n}", $"Reply {n}.", $"1{n}000002"),
            Ex($"1{n}000001", done: done) + Ex($"1{n}000002", $"1{n}000001", done));
        var order = resolverFirst ? new[] { (2, true), (1, false), (3, false) } : new[] { (1, false), (2, true), (3, false) };
        var reviews = order.Select(r => new DocxDiffReviewer { Document = Doc(Review(r.Item1, r.Item2)), Author = $"R{r.Item1}" }).ToList();

        var redline = DocxCompare.Consolidate(Doc(original), reviews).DocumentByteArray;

        Assert.NotEqual(true, Resolved(redline)[Root]);
        var threads = Threads(redline);
        Assert.All(new[] { 1, 2, 3 }, n => Assert.Equal(Root, threads[$"Reply {n}."]));
        AssertSelfConsistent(redline);
    }
}
