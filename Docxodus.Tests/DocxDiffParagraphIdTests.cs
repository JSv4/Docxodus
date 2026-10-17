// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1170: a <c>w14:paraId</c> is unique across the whole package and lies in <c>[1, 0x80000000)</c>.
/// A comparison output holds paragraphs and comments from two packages and mints ids for the comments it
/// clones, so every id it writes must avoid every id already in the body, headers, footers, notes,
/// comments, comment threading parts and glossary document.
/// </summary>
public class DocxDiffParagraphIdTests
{
    private const string W14Ns = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string W15Ns = "http://schemas.microsoft.com/office/word/2012/wordml";
    private const string W16CidNs = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
    private const string RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XName ParaId = XName.Get("paraId", W14Ns);

    private const string Alpha = "Alpha clause: the parties agree to these terms.";
    private const string Bravo = "Bravo clause: backups are verified monthly by staff.";
    private static readonly string BravoEdited = Bravo.Replace("monthly", "weekly");

    /// <summary>Two comments, the second a reply to the first, threaded through <c>commentsExtended</c> and
    /// given durable ids in <c>commentsIds</c>.</summary>
    private const string ThreadedComments =
        "<w:comment w:id=\"0\" w:author=\"Ann\" w:initials=\"A\" w:date=\"2026-01-01T00:00:00Z\"><w:p w14:paraId=\"00000001\"><w:r><w:t>Parent question.</w:t></w:r></w:p></w:comment>"
        + "<w:comment w:id=\"1\" w:author=\"Ann\" w:initials=\"A\" w:date=\"2026-01-01T00:00:00Z\"><w:p w14:paraId=\"00000002\"><w:r><w:t>Child reply.</w:t></w:r></w:p></w:comment>";

    private const string ThreadedCommentsEx =
        "<w15:commentEx w15:paraId=\"00000001\" w15:done=\"0\"/><w15:commentEx w15:paraId=\"00000002\" w15:paraIdParent=\"00000001\" w15:done=\"0\"/>";

    private const string ThreadedCommentIds =
        "<w16cid:commentId w16cid:paraId=\"00000001\" w16cid:durableId=\"11111111\"/><w16cid:commentId w16cid:paraId=\"00000002\" w16cid:durableId=\"22222222\"/>";

    private static void Write(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(xml);
    }

    /// <summary>A document with the given body, comments (and threading entries), and optionally a default header,
    /// a footnote with id 1, and a glossary document; every part declares <c>w14</c>.</summary>
    private static WmlDocument Build(string body, string comments = ThreadedComments, string commentsEx = ThreadedCommentsEx,
        string commentIds = ThreadedCommentIds, string? header = null, string? footnote = null, string? glossary = null)
    {
        string Root(string name, string inner, string extra = "") =>
            $"<w:{name} xmlns:w=\"{W.w.NamespaceName}\" xmlns:w14=\"{W14Ns}\"{extra}>{inner}</w:{name}>";
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            Write(main.AddNewPart<WordprocessingCommentsPart>(), Root("comments", comments));
            Write(main.AddNewPart<WordprocessingCommentsExPart>(), $"<w15:commentsEx xmlns:w15=\"{W15Ns}\">{commentsEx}</w15:commentsEx>");
            Write(main.AddNewPart<WordprocessingCommentsIdsPart>(), $"<w16cid:commentsIds xmlns:w16cid=\"{W16CidNs}\">{commentIds}</w16cid:commentsIds>");
            var sectPr = "";
            if (header != null)
            {
                Write(main.AddNewPart<HeaderPart>("rIdH1"), Root("hdr", header));
                sectPr = "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdH1\"/></w:sectPr>";
            }
            if (footnote != null)
                Write(main.AddNewPart<FootnotesPart>(), Root("footnotes",
                    "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
                    + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
                    + $"<w:footnote w:id=\"1\">{footnote}</w:footnote>"));
            if (glossary != null)
                Write(main.AddNewPart<GlossaryDocumentPart>(), Root("glossaryDocument",
                    $"<w:docParts><w:docPart><w:docPartPr><w:name w:val=\"Entry\"/></w:docPartPr><w:docPartBody>{glossary}</w:docPartBody></w:docPart></w:docParts>"));
            Write(main, Root("document", $"<w:body>{body}{sectPr}</w:body>", $" xmlns:r=\"{RNs}\""));
        }
        return new WmlDocument("paraids.docx", ms.ToArray());
    }

    private static string P(string text, string? paraId = null) =>
        $"<w:p{(paraId == null ? "" : $" w14:paraId=\"{paraId}\"")}><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string Sdt(string inner) =>
        $"<w:sdt><w:sdtPr><w:alias w:val=\"Article\"/><w:tag w:val=\"a1\"/><w:id w:val=\"101\"/></w:sdtPr><w:sdtContent>{inner}</w:sdtContent></w:sdt>";

    /// <summary>A content control whose blocks are bracketed by both threaded comments' ranges, so a comparison
    /// replaces it whole and gives each comment a deleted-view clone.</summary>
    private static string ThreadInWrapper(string bravo, string? alphaParaId = null, string? bravoParaId = null, string? referenceParaId = null) =>
        Sdt("<w:commentRangeStart w:id=\"0\"/><w:commentRangeStart w:id=\"1\"/>" + P(Alpha, alphaParaId) + P(bravo, bravoParaId)
            + "<w:commentRangeEnd w:id=\"0\"/><w:commentRangeEnd w:id=\"1\"/>"
            + $"<w:p{(referenceParaId == null ? "" : $" w14:paraId=\"{referenceParaId}\"")}><w:r><w:commentReference w:id=\"0\"/></w:r><w:r><w:commentReference w:id=\"1\"/></w:r></w:p>");

    private const string TailWithNote = "<w:p><w:r><w:t>Tail</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>";

    // ---- the ids every comparison writes ---------------------------------------------------------

    /// <summary>Every <c>w14:paraId</c> in the package, by part.</summary>
    private static List<(string Part, string Value)> ParaIds(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.GetAllParts()
            .Where(p => p.ContentType.EndsWith("xml", StringComparison.Ordinal))
            .SelectMany(p => (p.GetXDocument().Root?.DescendantsAndSelf().Attributes(ParaId) ?? Enumerable.Empty<XAttribute>())
                .Select(a => (p.Uri.ToString(), a.Value)))
            .ToList();
    }

    /// <summary>The package holds <paramref name="expected"/> paraIds (when given), all distinct and each in
    /// <c>[1, 0x80000000)</c>.</summary>
    private static void AssertParaIdsUniqueAndInRange(byte[] bytes, int? expected = null)
    {
        var ids = ParaIds(bytes);
        var listing = string.Join(", ", ids.Select(i => $"{i.Part}={i.Value}"));
        if (expected is { } count)
            Assert.True(ids.Count == count, $"expected {count} paraIds: {listing}");
        Assert.True(ids.Select(i => i.Value.ToUpperInvariant()).Distinct().Count() == ids.Count, $"duplicate w14:paraId: {listing}");
        Assert.All(ids, i => Assert.InRange(uint.Parse(i.Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture), 1u, 0x7FFFFFFFu));
    }

    /// <summary>Each comment definition's paraId keys exactly one <c>commentsExtended</c> and one
    /// <c>commentsIds</c> entry, and each reply's <c>paraIdParent</c> names a definition.</summary>
    private static void AssertThreadingFollowsTheDefinitions(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var definitionIds = main.WordprocessingCommentsPart!.GetXDocument().Root!.Elements(W.comment)
            .Select(c => (string)c.Elements(W.p).Last().Attribute(ParaId)!).ToList();
        var ex = main.WordprocessingCommentsExPart!.GetXDocument().Root!.Elements(XName.Get("commentEx", W15Ns)).ToList();
        var ids = main.WordprocessingCommentsIdsPart!.GetXDocument().Root!.Elements(XName.Get("commentId", W16CidNs)).ToList();
        foreach (var id in definitionIds)
        {
            Assert.Single(ex, e => (string?)e.Attribute(XName.Get("paraId", W15Ns)) == id);
            Assert.Single(ids, e => (string?)e.Attribute(XName.Get("paraId", W16CidNs)) == id);
        }
        foreach (var parent in ex.Select(e => (string?)e.Attribute(XName.Get("paraIdParent", W15Ns))).Where(p => p != null))
            Assert.Contains(parent, definitionIds);
    }

    private static void AssertNoNewValidationErrors(WmlDocument redline, params WmlDocument[] inputs)
    {
        static HashSet<string> Errors(WmlDocument document)
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            return new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package)
                .Select(e => e.Description).ToHashSet(StringComparer.Ordinal);
        }
        var known = inputs.SelectMany(Errors).ToHashSet(StringComparer.Ordinal);
        Assert.All(Errors(redline), e => Assert.Contains(e, known));
    }

    // ---- comment clones -------------------------------------------------------------------------

    public enum Holder
    {
        Header,
        Footnote,
        Glossary,
    }

    /// <summary>The ids one and two past the comments part's largest are held by a header, footnote or glossary
    /// paragraph, so the deleted-view clones of the two comments must take others — through Compare and
    /// Consolidate alike.</summary>
    [Theory]
    [InlineData(Holder.Header)]
    [InlineData(Holder.Footnote)]
    [InlineData(Holder.Glossary)]
    public void CommentClones_TakeParaIdsNoHeaderNoteOrGlossaryParagraphHolds(Holder holder)
    {
        var taken = P("Holds the next id.", "00000003") + P("Holds the one after.", "00000004");
        WmlDocument Document(string bravo) => Build(ThreadInWrapper(bravo) + (holder == Holder.Footnote ? TailWithNote : P("Tail")),
            header: holder == Holder.Header ? taken : null,
            footnote: holder == Holder.Footnote ? taken : null,
            glossary: holder == Holder.Glossary ? taken : null);
        var left = Document(Bravo);
        var right = Document(BravoEdited);

        var redlines = new[]
        {
            DocxDiff.Compare(left, right),
            DocxCompare.Consolidate(left, new[] { new DocxDiffReviewer { Document = right, Author = "Reviewer" } }),
        };

        foreach (var redline in redlines)
        {
            AssertParaIdsUniqueAndInRange(redline.DocumentByteArray, expected: 6); // 2 comments + 2 clones + 2 taken
            AssertThreadingFollowsTheDefinitions(redline.DocumentByteArray);
        }
    }

    /// <summary>The body already holds the largest valid paraId, 7FFFFFFF: one past it is out of range, so the
    /// clones take the lowest free values and the redline stays schema-valid.</summary>
    [Fact]
    public void CommentClones_AfterTheLargestValidParaId_TakeTheLowestFreeOnes()
    {
        WmlDocument Document(string bravo) => Build(ThreadInWrapper(bravo) + P("Tail", "7FFFFFFF"));
        var left = Document(Bravo);
        var right = Document(BravoEdited);

        var redline = DocxDiff.Compare(left, right);

        AssertParaIdsUniqueAndInRange(redline.DocumentByteArray, expected: 5);
        AssertThreadingFollowsTheDefinitions(redline.DocumentByteArray);
        AssertNoNewValidationErrors(redline, left, right);
    }

    // ---- the shared allocator --------------------------------------------------------------------

    private static (WordprocessingDocument Document, MemoryStream Stream) Open(WmlDocument document)
    {
        var stream = new MemoryStream();
        stream.Write(document.DocumentByteArray);
        return (WordprocessingDocument.Open(stream, true), stream);
    }

    [Fact]
    public void ParaIdAllocator_PastTheLargestValidId_TakesTheLowestFreeValuesInTurn()
    {
        var (document, stream) = Open(Build(P("One", "00000003") + P("Two", "7FFFFFFF") + P("Three", "80000005")));
        using (stream)
        using (document)
        {
            var allocator = new CommentOps.ParaIdAllocator(document.MainDocumentPart!);
            // 00000001/00000002 key the comments, 00000003 and 7FFFFFFF the body; 80000005 is out of range.
            Assert.Equal(new[] { "00000004", "00000005", "00000006" }, new[] { allocator.Next(), allocator.Next(), allocator.Next() });
        }
    }

    [Fact]
    public void ParaIdAllocator_BelowTheCeiling_TakesOnePastTheLargest()
    {
        var (document, stream) = Open(Build(P("One", "00000010") + P("Two", "0000000A")));
        using (stream)
        using (document)
        {
            var allocator = new CommentOps.ParaIdAllocator(document.MainDocumentPart!);
            Assert.Equal(new[] { "00000011", "00000012" }, new[] { allocator.Next(), allocator.Next() });
        }
    }

    [Fact]
    public void NextDurableId_PastTheLargestSignedValue_TakesTheLowestFreeOne()
    {
        var root = XElement.Parse($"<w16cid:commentsIds xmlns:w16cid=\"{W16CidNs}\">"
            + "<w16cid:commentId w16cid:paraId=\"00000001\" w16cid:durableId=\"00000001\"/>"
            + "<w16cid:commentId w16cid:paraId=\"00000002\" w16cid:durableId=\"7FFFFFFE\"/></w16cid:commentsIds>");
        Assert.Equal("00000002", CommentOps.NextDurableId(root));
    }

    // ---- paragraphs from both packages -----------------------------------------------------------

    /// <summary>A content control replaced whole keeps a deleted and an inserted copy of its paragraphs, each
    /// with the source's paraIds; the copies must not share them — through Compare and Consolidate alike.</summary>
    [Fact]
    public void WrapperReplacedWhole_GivesEachCopysParagraphsTheirOwnParaIds()
    {
        WmlDocument Document(string bravo) => Build(ThreadInWrapper(bravo, "0000000A", "0000000B", "0000000C") + P("Tail"));
        var left = Document(Bravo);
        var right = Document(BravoEdited);

        var redlines = new[]
        {
            DocxDiff.Compare(left, right),
            DocxCompare.Consolidate(left, new[] { new DocxDiffReviewer { Document = right, Author = "Reviewer" } }),
        };

        foreach (var redline in redlines)
        {
            Assert.Equal(2, Body(redline).Descendants(W.sdt).Count());
            AssertParaIdsUniqueAndInRange(redline.DocumentByteArray, expected: 10); // 2 comments + 2 clones + 2 × 3 paragraphs
            AssertThreadingFollowsTheDefinitions(redline.DocumentByteArray);
        }
    }

    /// <summary>A paragraph rewritten whole is drawn as a deleted copy and an inserted copy, both carrying the
    /// paraId the two sources share — through both compare front doors.</summary>
    [Fact]
    public void ParagraphRewrittenWhole_GivesItsTwoCopiesTheirOwnParaIds()
    {
        var left = Build(P("The quick brown fox jumps over the lazy dog.", "00000010") + P("Tail"), comments: "", commentsEx: "", commentIds: "");
        var right = Build(P("Entirely different words appear in this sentence now.", "00000010") + P("Tail"), comments: "", commentsEx: "", commentIds: "");

        foreach (var redline in new[] { DocxDiff.Compare(left, right), DocxCompare.Compare(left, right) })
        {
            Assert.Contains(Body(redline).Descendants(), e => e.Name == W.del);
            AssertParaIdsUniqueAndInRange(redline.DocumentByteArray, expected: 2);
        }
    }

    /// <summary>The revision adds a comment whose definition paraId, in its own package, is free — but the
    /// original uses it for a paragraph the revision deletes. The merged definition keeps its threading key,
    /// so the deleted paragraph is the one renamed.</summary>
    [Fact]
    public void AddedComment_ParaIdHeldByADeletedParagraph_IsNotShared()
    {
        const string Added = "<w:comment w:id=\"2\" w:author=\"Ann\" w:initials=\"A\" w:date=\"2026-01-01T00:00:00Z\"><w:p w14:paraId=\"00000009\"><w:r><w:t>New remark.</w:t></w:r></w:p></w:comment>";
        var left = Build(P(Alpha) + P("Deleted paragraph here.", "00000009") + P("Tail"));
        var right = Build(P(Alpha) + "<w:p><w:commentRangeStart w:id=\"2\"/><w:r><w:t>Tail</w:t></w:r><w:commentRangeEnd w:id=\"2\"/><w:r><w:commentReference w:id=\"2\"/></w:r></w:p>",
            comments: ThreadedComments + Added,
            commentsEx: ThreadedCommentsEx + "<w15:commentEx w15:paraId=\"00000009\" w15:done=\"0\"/>",
            commentIds: ThreadedCommentIds + "<w16cid:commentId w16cid:paraId=\"00000009\" w16cid:durableId=\"33333333\"/>");

        var redline = DocxDiff.Compare(left, right);

        AssertParaIdsUniqueAndInRange(redline.DocumentByteArray);
        Assert.Single(ParaIds(redline.DocumentByteArray), i => i.Part == "/word/comments.xml" && i.Value == "00000009");
        AssertThreadingFollowsTheDefinitions(redline.DocumentByteArray);
    }

    /// <summary>The original's reply (comment 1) survives on deleted text while the revision adds its own reply
    /// (comment 2) that its package numbered with the same paraId. The added reply takes a fresh paraId and its
    /// threading entries follow it, still answering comment 0.</summary>
    [Fact]
    public void AddedReply_ParaIdHeldByTheOriginalsReply_TakesAFreshOneAndKeepsItsThread()
    {
        const string RightReply = "<w:comment w:id=\"2\" w:author=\"Bo\" w:initials=\"B\" w:date=\"2026-01-02T00:00:00Z\"><w:p w14:paraId=\"00000002\"><w:r><w:t>Another reply.</w:t></w:r></w:p></w:comment>";
        static string Body(int reply, string tail) =>
            "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t xml:space=\"preserve\">Parent anchor. </w:t></w:r><w:commentRangeEnd w:id=\"0\"/>"
            + "<w:r><w:commentReference w:id=\"0\"/></w:r></w:p>"
            + $"<w:p><w:commentRangeStart w:id=\"{reply}\"/><w:r><w:t xml:space=\"preserve\">{tail}</w:t></w:r><w:commentRangeEnd w:id=\"{reply}\"/>"
            + $"<w:r><w:commentReference w:id=\"{reply}\"/></w:r></w:p>";
        var left = Build(Body(1, "The original reply's anchor text."));
        var right = Build(Body(2, "A completely rewritten sentence now."),
            comments: ThreadedComments.Substring(0, ThreadedComments.IndexOf("<w:comment w:id=\"1\"", StringComparison.Ordinal)) + RightReply,
            commentsEx: "<w15:commentEx w15:paraId=\"00000001\" w15:done=\"0\"/><w15:commentEx w15:paraId=\"00000002\" w15:paraIdParent=\"00000001\" w15:done=\"0\"/>",
            commentIds: "<w16cid:commentId w16cid:paraId=\"00000001\" w16cid:durableId=\"11111111\"/><w16cid:commentId w16cid:paraId=\"00000002\" w16cid:durableId=\"33333333\"/>");

        var redline = DocxDiff.Compare(left, right);

        AssertParaIdsUniqueAndInRange(redline.DocumentByteArray, expected: 3);
        AssertThreadingFollowsTheDefinitions(redline.DocumentByteArray);
        using var stream = new MemoryStream(redline.DocumentByteArray);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var addedParaId = (string)main.WordprocessingCommentsPart!.GetXDocument().Root!.Elements(W.comment)
            .Single(c => (string?)c.Attribute(W.id) == "2").Elements(W.p).Last().Attribute(ParaId)!;
        var entry = main.WordprocessingCommentsExPart!.GetXDocument().Root!.Elements(XName.Get("commentEx", W15Ns))
            .Single(e => (string?)e.Attribute(XName.Get("paraId", W15Ns)) == addedParaId);
        Assert.Equal("00000001", (string?)entry.Attribute(XName.Get("paraIdParent", W15Ns)));
    }

    /// <summary>The revision adds a comment whose durable id, in its own package, is the one the original gives
    /// comment 0. The merged entry takes a fresh durable id, so no two comments share one.</summary>
    [Fact]
    public void AddedComment_DurableIdHeldByAnOriginalComment_TakesAFreshOne()
    {
        const string Anchored = "<w:p><w:commentRangeStart w:id=\"0\"/><w:commentRangeStart w:id=\"1\"/><w:r><w:t>Alpha</w:t></w:r>"
            + "<w:commentRangeEnd w:id=\"0\"/><w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"0\"/></w:r>"
            + "<w:r><w:commentReference w:id=\"1\"/></w:r></w:p>";
        const string Added = "<w:comment w:id=\"5\" w:author=\"Ann\" w:initials=\"A\" w:date=\"2026-01-01T00:00:00Z\"><w:p w14:paraId=\"00000005\"><w:r><w:t>New remark.</w:t></w:r></w:p></w:comment>";
        var left = Build(Anchored + P("Tail"));
        var right = Build(Anchored + "<w:p><w:commentRangeStart w:id=\"5\"/><w:r><w:t>Tail</w:t></w:r><w:commentRangeEnd w:id=\"5\"/><w:r><w:commentReference w:id=\"5\"/></w:r></w:p>",
            comments: ThreadedComments + Added,
            commentsEx: ThreadedCommentsEx + "<w15:commentEx w15:paraId=\"00000005\" w15:done=\"0\"/>",
            commentIds: ThreadedCommentIds + "<w16cid:commentId w16cid:paraId=\"00000005\" w16cid:durableId=\"11111111\"/>");

        var redline = DocxCompare.Compare(left, right);

        using var stream = new MemoryStream(redline.DocumentByteArray);
        using var document = WordprocessingDocument.Open(stream, false);
        var entries = document.MainDocumentPart!.WordprocessingCommentsIdsPart!.GetXDocument().Root!
            .Elements(XName.Get("commentId", W16CidNs)).ToList();
        var durableIds = entries.Select(e => (string)e.Attribute(XName.Get("durableId", W16CidNs))!).ToList();
        Assert.True(durableIds.Count == durableIds.Distinct(StringComparer.OrdinalIgnoreCase).Count(),
            $"duplicate durable id: {string.Join(", ", durableIds)}");
        var added = Assert.Single(entries, e => (string?)e.Attribute(XName.Get("paraId", W16CidNs)) == "00000005");
        Assert.NotEqual("11111111", (string?)added.Attribute(XName.Get("durableId", W16CidNs)));
        Assert.Equal("11111111", (string?)entries.Single(e => (string?)e.Attribute(XName.Get("paraId", W16CidNs)) == "00000001")
            .Attribute(XName.Get("durableId", W16CidNs)));
        AssertThreadingFollowsTheDefinitions(redline.DocumentByteArray);
    }

    private static XElement Body(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        return new XElement(package.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!);
    }
}
