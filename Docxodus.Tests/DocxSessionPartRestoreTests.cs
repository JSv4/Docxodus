using System.IO.Packaging;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionPartRestoreTests
{
    [Fact]
    public void RestoringStoryTopologyKeepsExistingImageSnapshotBytesShared()
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes(),
            new DocxSessionSettings { EmitMarkdownPatch = false });
        var body = session.Project().AnchorIndex.Keys.First(id => id.StartsWith("p:body:", StringComparison.Ordinal));
        var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=");
        Assert.True(session.InsertImage(body, 0, png).Success);
        var bytes = Assert.Single(session.TakeSnapshot().ImageParts).Bytes;
        Assert.True(session.InsertFootnote(body, 0, "Note").Success);

        Assert.True(session.Undo());
        Assert.Same(bytes, Assert.Single(session.TakeSnapshot().ImageParts).Bytes);
        Assert.True(session.Redo());
        Assert.Same(bytes, Assert.Single(session.TakeSnapshot().ImageParts).Bytes);
    }

    [Theory]
    [InlineData("footnote")]
    [InlineData("endnote")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("comment")]
    public void RecreatedStoriesRestoreEverySavedStateIncludingLinksAndImages(string kind)
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes(),
            new DocxSessionSettings { EmitMarkdownPatch = false });
        var body = session.Project().AnchorIndex.Keys.First(id => id.StartsWith("p:body:", StringComparison.Ordinal));
        Assert.True(session.ReplaceText(body, "Body text").Success);
        var states = new List<string[]> { PackageState(session) };
        var created = kind switch
        {
            "footnote" => session.InsertFootnote(body, 0, "Story text"),
            "endnote" => session.InsertEndnote(body, 0, "Story text"),
            "header" => session.SetHeaderText(body, HeaderFooterKind.Default, "Story text"),
            "footer" => session.SetFooterText(body, HeaderFooterKind.Default, "Story text"),
            "comment" => session.AddComment(body, null, "Reviewer", "Story text"),
            _ => throw new ArgumentOutOfRangeException(nameof(kind)),
        };
        Assert.True(created.Success, created.Error?.Message);
        states.Add(PackageState(session));
        var story = session.Project().AnchorIndex.Values.First(t =>
            t.Anchor.Scope != "body" && t.Anchor.Kind == "p").Anchor.Id;
        Assert.True(session.ReplaceText(body, "Later body edit").Success);
        states.Add(PackageState(session));
        Assert.True(session.ReplaceText(story, "Story [link](https://example.com/reference)").Success);
        states.Add(PackageState(session));
        var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=");
        var image = session.InsertImage(story, 0, png);
        Assert.True(image.Success, image.Error?.Message);
        states.Add(PackageState(session));

        AssertRoundTrips(session, states, kind);
    }

    [Theory]
    [InlineData("numbering")]
    [InlineData("styles")]
    [InlineData("settings")]
    [InlineData("comment-threading")]
    [InlineData("annotations")]
    public void RecreatedSupportingPartsRestoreLaterEdits(string kind)
    {
        using var input = new MemoryStream();
        input.Write(DocxSession.CreateBlankDocxBytes());
        using (var doc = WordprocessingDocument.Open(input, true))
        {
            var main = doc.MainDocumentPart!;
            if (kind == "styles") main.DeletePart(main.StyleDefinitionsPart!);
            if (kind == "settings") main.DeletePart(main.DocumentSettingsPart!);
            main.GetXDocument().Root!.Element(W.body)!.Element(W.p)!
                .Add(new XElement(W.r, new XElement(W.t, "Body text")));
            main.PutXDocument();
        }
        using var session = new DocxSession(input.ToArray(), new DocxSessionSettings { EmitMarkdownPatch = false });
        string Body() => session.Project().AnchorIndex.Values.First(t => t.Anchor.Scope == "body"
            && t.Anchor.Kind is "p" or "li").Anchor.Id;
        string? comment = null;
        if (kind == "comment-threading")
        {
            var result = session.AddComment(Body(), null, "Reviewer", "Comment");
            Assert.True(result.Success, result.Error?.Message);
            comment = result.Created.First(a => a.Kind == "cmt").Id;
        }
        var states = new List<string[]> { PackageState(session) };
        var create = kind switch
        {
            "numbering" => session.ApplyListFormat(Body(), ListFormat.Decimal),
            "styles" => session.InsertParagraph(Body(), Position.After, "# First heading"),
            "settings" => session.InsertFootnote(Body(), 0, "Footnote"),
            "comment-threading" => session.AddCommentReply(comment!, "Reviewer", "Reply"),
            "annotations" => session.AddAnnotation(Body(), null, new DocumentAnnotation("ann", "tag", "Original", "#FFFF00")),
            _ => throw new ArgumentOutOfRangeException(nameof(kind)),
        };
        Assert.True(create.Success, create.Error?.Message);
        states.Add(PackageState(session));
        Assert.True(session.ReplaceText(Body(), "Later body edit").Success);
        states.Add(PackageState(session));
        var edit = kind switch
        {
            "numbering" => session.ApplyListFormat(Body(), ListFormat.Bullet),
            "styles" => session.InsertParagraph(Body(), Position.After, "## Second heading"),
            "settings" => session.InsertEndnote(Body(), 0, "Endnote"),
            "comment-threading" => session.SetCommentResolved(comment!, true),
            "annotations" => session.UpdateAnnotation("ann", new AnnotationUpdate { Label = "Edited" }),
            _ => throw new ArgumentOutOfRangeException(nameof(kind)),
        };
        Assert.True(edit.Success, edit.Error?.Message);
        states.Add(PackageState(session));
        AssertRoundTrips(session, states, kind);
    }

    private static void AssertRoundTrips(DocxSession session, List<string[]> states, string kind)
    {
        for (int round = 0; round < 2; round++)
        {
            for (int i = states.Count - 2; i >= 0; i--)
            {
                Assert.True(session.Undo());
                AssertState(states[i], PackageState(session), $"{kind} undo {i}");
            }
            for (int i = 1; i < states.Count; i++)
            {
                Assert.True(session.Redo());
                AssertState(states[i], PackageState(session), $"{kind} redo {i}");
            }
        }
    }

    private static void AssertState(string[] expected, string[] actual, string context)
    {
        Assert.Equal(expected.Length, actual.Length);
        for (int i = 0; i < expected.Length; i++)
            Assert.True(expected[i] == actual[i], $"{context}\nExpected: {expected[i]}\nActual: {actual[i]}");
    }

    private static string[] PackageState(DocxSession session)
    {
        using var stream = new MemoryStream(session.Save());
        using var package = Package.Open(stream, FileMode.Open, FileAccess.Read);
        var state = new List<string> { Relationships(package.GetRelationships()) };
        foreach (var part in package.GetParts().Where(p => !PackUriHelper.IsRelationshipPartUri(p.Uri))
            .OrderBy(p => p.Uri.ToString(), StringComparer.Ordinal))
        {
            using var input = part.GetStream();
            var prefix = part.Uri + ":" + part.ContentType + ":" + Relationships(part.GetRelationships()) + ":";
            if (part.ContentType.EndsWith("xml", StringComparison.Ordinal))
            {
                state.Add(prefix + CanonicalXml(XDocument.Load(input)));
            }
            else
            {
                using var bytes = new MemoryStream();
                input.CopyTo(bytes);
                state.Add(prefix + Convert.ToHexString(bytes.ToArray()));
            }
        }
        return state.ToArray();

        static string Relationships(IEnumerable<PackageRelationship> relationships) =>
            string.Join(";", relationships.OrderBy(r => r.Id, StringComparer.Ordinal)
                .Select(r => $"{r.Id}|{r.RelationshipType}|{r.TargetMode}|" +
                    (r.TargetMode == TargetMode.Internal ? PackUriHelper.ResolvePartUri(r.SourceUri, r.TargetUri) : r.TargetUri)));
    }

    // Prefix declarations and attribute order are serialization choices; expanded names and
    // values are the XML state. Content types are compared per part, ignoring unused defaults.
    private static string CanonicalXml(XDocument document) =>
        Copy(document.Root!).ToString(SaveOptions.DisableFormatting);

    private static XElement Copy(XElement element) => new(element.Name,
        element.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString(), StringComparer.Ordinal),
        element.Nodes().Select(n => n is XElement child ? Copy(child) : n));

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void RedoRestoresLaterEditsInsideARecreatedNotePart(bool endnotes, bool emitPatch)
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes(),
            new DocxSessionSettings { EmitMarkdownPatch = emitPatch });
        var body = session.Project().AnchorIndex.Keys.First(id => id.StartsWith("p:body:", StringComparison.Ordinal));
        var inserted = endnotes
            ? session.InsertEndnote(body, 0, "Note text")
            : session.InsertFootnote(body, 0, "Note text");
        Assert.True(inserted.Success, inserted.Error?.Message);
        var note = session.Project().AnchorIndex.Values.First(t =>
            t.Anchor.Scope == (endnotes ? "en" : "fn") && t.Anchor.Kind == "p").Anchor.Id;
        Assert.True(session.ReplaceText(body, "Body edit").Success);
        Assert.True(session.ApplyListFormat(note, ListFormat.Decimal).Success);
        var expected = NoteXml();

        for (int i = 0; i < 3; i++) Assert.True(session.Undo());
        for (int i = 0; i < 3; i++) Assert.True(session.Redo());

        Assert.Equal(CanonicalXml(expected), CanonicalXml(NoteXml()));
        using var reopened = new DocxSession(session.Save());
        Assert.Contains(reopened.Project().AnchorIndex.Values,
            t => t.Anchor.Scope == (endnotes ? "en" : "fn") && t.Anchor.Kind == "li");

        XDocument NoteXml()
        {
            var main = session.LiveDocument.MainDocumentPart!;
            return new XDocument((endnotes
                ? main.EndnotesPart!.GetXDocument()
                : main.FootnotesPart!.GetXDocument()));
        }
    }
}
