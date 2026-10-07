using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir.Diff;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Guards for the renderer passes that used to rescan the whole body once per item: comment
/// normalization (one scan per comment id) and unpaired move-range lowering (one root scan per
/// unpaired start). Each now scans once or walks forward, and must produce the same result.
/// </summary>
public class IrMarkupRendererScanTests
{
    [Fact]
    public void ElementsAfterInDocumentOrder_matches_root_descendants_after_every_element()
    {
        var root = XElement.Parse(
            "<a><b><c/><d><e/></d></b><f/><g><h><i/></h><j/></g></a>");

        foreach (var start in root.DescendantsAndSelf())
        {
            var expected = root.DescendantsAndSelf().SkipWhile(e => e != start).Skip(1).ToList();
            var actual = IrMarkupRenderer.ElementsAfterInDocumentOrder(start).ToList();
            Assert.Equal(expected, actual);
        }
    }

    [Fact]
    public void Compare_collapses_each_of_many_edited_comments_to_one_bare_range()
    {
        const int comments = 60;
        var left = new WmlDocument("left.docx", BuildCommentedDocument(comments, edited: false));
        var right = new WmlDocument("right.docx", BuildCommentedDocument(comments, edited: true));

        var result = DocxCompare.Compare(left, right);

        using var stream = new MemoryStream(result.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        for (int id = 0; id < comments; id++)
        {
            foreach (var kind in new[] { W.commentRangeStart, W.commentRangeEnd, W.commentReference })
            {
                var marker = Assert.Single(body.Descendants(kind), m => (string?)m.Attribute(W.id) == id.ToString());
                Assert.DoesNotContain(marker.Ancestors(), a => a.Name == W.ins || a.Name == W.del);
            }
        }
    }

    /// <summary>
    /// One paragraph per comment, the comment anchored on a middle run. The edited copy changes
    /// the anchored text, so the diff renders a deleted and an inserted copy of each range — the
    /// case comment normalization collapses back to a single bare range per comment.
    /// </summary>
    private static byte[] BuildCommentedDocument(int comments, bool edited)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            var commentsPart = main.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Wp.Comments();
            var body = new Wp.Body();
            for (int i = 0; i < comments; i++)
            {
                var id = i.ToString();
                commentsPart.Comments.Append(new Wp.Comment(new Wp.Paragraph(new Wp.Run(new Wp.Text($"note {id}"))))
                {
                    Id = id,
                    Author = "Reviewer",
                    Initials = "R",
                });
                body.Append(new Wp.Paragraph(
                    new Wp.Run(new Wp.Text($"Lead text {i} ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Wp.CommentRangeStart { Id = id },
                    new Wp.Run(new Wp.Text(edited ? $"changed words {i}" : $"original words {i}")),
                    new Wp.CommentRangeEnd { Id = id },
                    new Wp.Run(new Wp.CommentReference { Id = id }),
                    new Wp.Run(new Wp.Text($" tail {i}.") { Space = SpaceProcessingModeValues.Preserve })));
            }
            main.Document = new Wp.Document(body);
        }
        return ms.ToArray();
    }
}
