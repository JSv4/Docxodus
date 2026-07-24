#nullable enable

using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DocxDiffSettings.NormalizeRevisionAuthors"/> collapses every tracked-revision author in the
/// output to the single <see cref="DocxDiffSettings.AuthorForRevisions"/>. Renderers that color tracked
/// changes by author (LibreOffice) then show ONE color, matching Word's single-author compare output —
/// instead of a second color leaking from an input revision preserved under its original author
/// (<see cref="DocxDiffSettings.PreserveInputRevisions"/>). The flag defaults off (no-op), never touches
/// comment authors, and does not alter revision structure (accept/reject unaffected).
/// </summary>
public class DocxDiffAuthorNormalizationTests
{
    // A right document whose ACCEPTED view equals the left, carrying one pre-existing tracked insertion
    // authored by a foreign user ("Online User") plus a comment authored by someone else again.
    private static WmlDocument BuildLeft() =>
        BuildDoc("<w:p><w:r><w:t xml:space=\"preserve\">Hello brave new world.</w:t></w:r></w:p>");

    private static WmlDocument BuildRight() =>
        BuildDoc(
            "<w:p><w:r><w:t xml:space=\"preserve\">Hello </w:t></w:r>" +
            "<w:ins w:id=\"7\" w:author=\"Online User\" w:date=\"2026-01-01T00:00:00Z\">" +
            "<w:r><w:t xml:space=\"preserve\">brave new </w:t></w:r></w:ins>" +
            "<w:r><w:t>world.</w:t></w:r></w:p>");

    private static WmlDocument BuildDoc(string bodyInner)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var xml =
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                bodyInner + "</w:body></w:document>";
            using var s = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(s, new UTF8Encoding(false));
            writer.Write(xml);
        }
        return new WmlDocument("d.docx", stream.ToArray());
    }

    private static string[] RevisionAuthors(WmlDocument result)
    {
        using var s = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(s, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.GetStream());
        var xml = reader.ReadToEnd();
        return Regex.Matches(xml, "w:author=\"([^\"]*)\"").Select(m => m.Groups[1].Value).ToArray();
    }

    [Fact]
    public void NormalizeRevisionAuthors_collapses_preserved_foreign_author_to_single()
    {
        var settings = new DocxDiffSettings
        {
            PreserveInputRevisions = true,
            NormalizeRevisionAuthors = true,
            AuthorForRevisions = "Comparison",
        };

        var authors = RevisionAuthors(DocxDiff.Compare(BuildLeft(), BuildRight(), settings));

        Assert.NotEmpty(authors);                    // the preserved insertion IS in the output
        Assert.DoesNotContain("Online User", authors);
        Assert.All(authors, a => Assert.Equal("Comparison", a));
    }

    [Fact]
    public void Without_NormalizeRevisionAuthors_the_foreign_author_leaks()
    {
        var settings = new DocxDiffSettings
        {
            PreserveInputRevisions = true,
            NormalizeRevisionAuthors = false,
            AuthorForRevisions = "Comparison",
        };

        var authors = RevisionAuthors(DocxDiff.Compare(BuildLeft(), BuildRight(), settings));

        Assert.Contains("Online User", authors);     // the leak this flag exists to fix
    }
}
