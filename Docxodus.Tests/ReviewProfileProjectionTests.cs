#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using static Docxodus.Tests.Ir.Diff.RevisionsInInputFixtures;
using Xunit;

namespace Docxodus.Tests;

public class ReviewProfileProjectionTests
{
    [Theory]
    [InlineData("comments", "CMTEDIT")]
    [InlineData("glossary", "GLOSSEDIT")]
    public void ExportProjection_ResolvesAuxiliaryStoriesWithoutChangingLegacyScope(
        string story, string insertedToken)
    {
        var source = story == "comments"
            ? CommentWithRevisionDoc("Body", "Reviewer")
            : GlossaryWithRevisionDoc("Body", "Reviewer");

        // The historical diff helper deliberately keeps its narrower contract.
        Assert.NotEmpty(AllRevisionElementNames(RevisionProcessor.AcceptRevisions(source)));

        var final = new WmlDocument("final.docx",
            DocxDiffOps.ProjectReviewProfile(source.DocumentByteArray, "final"));
        var original = new WmlDocument("original.docx",
            DocxDiffOps.ProjectReviewProfile(source.DocumentByteArray, "original"));

        Assert.Empty(AllRevisionElementNames(final));
        Assert.Empty(AllRevisionElementNames(original));
        Assert.Contains(insertedToken, AllContentText(final));
        Assert.DoesNotContain(insertedToken, AllContentText(original));
    }

    [Theory]
    [InlineData("comments", "/word/comments.xml", "cmt", "Reviewer")]
    [InlineData("glossary", "/word/glossary/document.xml", "glossary", "Reviewer")]
    public void ExportInventory_ReportsEveryAuxiliaryStoryOwnedByProjection(
        string story, string partUri, string scope, string author)
    {
        var source = story == "comments"
            ? CommentWithRevisionDoc("Body", author)
            : GlossaryWithRevisionDoc("Body", author);
        using var session = new DocxSession(source.DocumentByteArray);

        Assert.Empty(session.ListRevisions());
        var revision = Assert.Single(session.ListRevisionsForExportProfile());
        Assert.Equal(partUri, revision.PartUri);
        Assert.Equal(scope, revision.Scope);
        Assert.Equal(author, revision.Author);
        Assert.Equal(RevisionFamily.ContentInsert, revision.Family);
        Assert.Equal(RevisionResolutionStatus.Supported, revision.ResolutionStatus);
    }

    [Fact]
    public void MarkupProjection_IsAnExactOwnedCopy()
    {
        var source = CommentWithRevisionDoc("Body", "Reviewer").DocumentByteArray;
        var projected = DocxDiffOps.ProjectReviewProfile(source, "markup");

        Assert.NotSame(source, projected);
        Assert.True(source.SequenceEqual(projected));
    }

    [Fact]
    public void Projection_RejectsUnknownProfile()
    {
        var source = CommentWithRevisionDoc("Body", "Reviewer").DocumentByteArray;
        Assert.Throws<ArgumentException>(() =>
            DocxDiffOps.ProjectReviewProfile(source, "current"));
    }

    private static string AllContentText(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        return string.Concat(ContentParts(package).SelectMany(part =>
        {
            using var partStream = part.GetStream(FileMode.Open, FileAccess.Read);
            var xml = XDocument.Load(partStream);
            return xml.Descendants(Wn + "t").Select(text => text.Value);
        }));
    }
}
