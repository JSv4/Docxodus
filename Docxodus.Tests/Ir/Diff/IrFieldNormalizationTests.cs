using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// The redline's field-context pass (<c>IrMarkupRenderer.NormalizeFieldsInRoot</c>) groups each paragraph's
/// fields by their characters. An <c>end</c> that closes a field opened in an earlier paragraph belongs to no
/// field of this one (issue #1166), and a field with no result whose characters all sit in one revision is
/// that revision's field.
/// </summary>
public class IrFieldNormalizationTests
{
    private static readonly XNamespace Ns = W.w;

    private static XElement Run(XElement child) => new(W.r, child);
    private static XElement FieldChar(string type) => Run(new XElement(W.fldChar, new XAttribute(W.fldCharType, type)));
    private static XElement Text(string text) => Run(new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), text));
    private static XElement Inserted(XElement run) =>
        new(W.ins, new XAttribute(W.id, "1"), new XAttribute(W.author, "A"), run);
    private static XElement Deleted(XElement run) =>
        new(W.del, new XAttribute(W.id, "2"), new XAttribute(W.author, "A"), run);

    private static XElement Document(params object[] paragraphContent) =>
        new(W.document, new XAttribute(XNamespace.Xmlns + "w", Ns.NamespaceName),
            new XElement(W.body, new XElement(W.p, paragraphContent)));

    /// <summary>A field whose result is unchanged has its plumbing lifted out of an insertion, also when the
    /// paragraph first closes a field opened before it.</summary>
    [Fact]
    public void NormalizeFieldsInRoot_FieldAfterAnEndFromAnEarlierParagraph_LiftsItsPlumbingBare()
    {
        var root = Document(
            FieldChar("end"),
            Text(" text "),
            Inserted(FieldChar("begin")),
            Inserted(Run(new XElement(W.instrText, " PAGE "))),
            Inserted(FieldChar("separate")),
            Text("3"),
            Inserted(FieldChar("end")));

        Assert.True(IrMarkupRenderer.NormalizeFieldsInRoot(root));

        Assert.Empty(root.Descendants(W.ins));
        Assert.Equal(new[] { "end", "begin", "separate", "end" },
            root.Descendants(W.fldChar).Select(f => (string?)f.Attribute(W.fldCharType)));
    }

    /// <summary>The end of a field opened before the paragraph is grouped with nothing: the field after it is
    /// re-homed while the stray end keeps its own insertion.</summary>
    [Fact]
    public void NormalizeFieldsInRoot_InsertedEndFromAnEarlierParagraph_StaysInsertedWhileTheNextFieldIsLifted()
    {
        var root = Document(
            Text("Second line"),
            Inserted(FieldChar("end")),
            Text(" text "),
            Inserted(FieldChar("begin")),
            Inserted(Run(new XElement(W.instrText, " PAGE "))),
            Inserted(FieldChar("separate")),
            Text("3"),
            Inserted(FieldChar("end")));

        Assert.True(IrMarkupRenderer.NormalizeFieldsInRoot(root));

        var stillInserted = Assert.Single(root.Descendants(W.ins).Descendants(W.fldChar));
        Assert.Same(root.Descendants(W.fldChar).First(), stillInserted);
    }

    /// <summary>A field with no result (an index entry) whose characters all sit in one deletion is a deleted
    /// field: it is not lifted out, also after an end from an earlier paragraph.</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NormalizeFieldsInRoot_ResultlessFieldWhollyDeleted_StaysDeleted(bool afterAnEndFromAnEarlierParagraph)
    {
        var root = Document(
            afterAnEndFromAnEarlierParagraph ? FieldChar("end") : Text(""),
            Text("Hello "),
            Deleted(FieldChar("begin")),
            Deleted(Run(new XElement(W.delInstrText, " XE \"foo\" "))),
            Deleted(FieldChar("end")),
            Text("world"));

        Assert.False(IrMarkupRenderer.NormalizeFieldsInRoot(root));

        Assert.Equal(2, root.Descendants(W.del).Descendants(W.fldChar).Count());
        Assert.Single(root.Descendants(W.delInstrText));
    }

    // ---- through the comparison, keeping the inputs' own revisions -----------------------------

    private static string R(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";
    private static string F(string type) => $"<w:r><w:fldChar w:fldCharType=\"{type}\"/></w:r>";
    private static string DeletedIndexEntry =>
        "<w:del w:id=\"901\" w:author=\"Other\" w:date=\"2020-01-01T00:00:00Z\">" + F("begin") +
        "<w:r><w:delInstrText xml:space=\"preserve\"> XE \"foo\" </w:delInstrText></w:r>" + F("end") + "</w:del>";
    private const string OpensField =
        "<w:p><w:r><w:t xml:space=\"preserve\">Intro </w:t></w:r><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        "<w:r><w:instrText xml:space=\"preserve\"> REF a \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        "<w:r><w:t xml:space=\"preserve\">refres </w:t></w:r></w:p>";

    /// <summary>The revised document carries an index-entry field its own tracked change deleted, and the
    /// comparison keeps the inputs' revisions: accepting the redline must not bring the deleted field back,
    /// also in a paragraph that first closes a field opened in the paragraph before it.</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Compare_PreservingAnInputDeletionOfAResultlessField_AcceptKeepsItDeleted(bool afterAnEndFromAnEarlierParagraph)
    {
        var prefix = afterAnEndFromAnEarlierParagraph ? OpensField : "";
        var lead = afterAnEndFromAnEarlierParagraph ? R("Second ") + F("end") + R(" hello ") : R("Hello ");
        var original = IrTestDocuments.FromBodyXml(prefix + "<w:p>" + lead + R("world") + "</w:p>" +
            "<w:p><w:r><w:t>Other paragraph</w:t></w:r></w:p>");
        var revised = IrTestDocuments.FromBodyXml(prefix + "<w:p>" + lead + DeletedIndexEntry + R("world") + "</w:p>" +
            "<w:p><w:r><w:t>Other paragraph changed</w:t></w:r></w:p>");

        var redline = DocxCompare.Compare(original, revised, new DocxDiffSettings { PreserveInputRevisions = true });

        Assert.Equal(Shape(DocxDiffOps.AcceptRevisions(revised.DocumentByteArray)),
            Shape(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray)));
    }

    private static string I(string instruction) =>
        $"<w:r><w:instrText xml:space=\"preserve\"> {instruction} </w:instrText></w:r>";
    private static string Para(string content) => "<w:p>" + content + "</w:p>";

    /// <summary>The result-less fields: an index entry, a table-of-contents entry, and a field whose result is
    /// empty.</summary>
    private static string ResultlessField(string kind) => kind switch
    {
        "XE" => F("begin") + I("XE \"foo\"") + F("end"),
        "TC" => F("begin") + I("TC \"Entry\" \\l 1") + F("end"),
        _ => F("begin") + I("PAGE") + F("separate") + F("end"),
    };

    /// <summary>A default comparison whose redline draws a result-less field's paragraph as a deleted and an
    /// inserted copy (the text on both sides of the field replaced, or the paragraph split right after it): each
    /// view keeps one copy of the field, the field-context pass leaving both copies in their revisions.</summary>
    [Theory]
    [InlineData("XE", "text on both sides replaced")]
    [InlineData("TC", "text on both sides replaced")]
    [InlineData("empty-result PAGE", "text on both sides replaced")]
    [InlineData("XE", "paragraph split right after the field")]
    [InlineData("TC", "paragraph split right after the field")]
    [InlineData("empty-result PAGE", "paragraph split right after the field")]
    public void Compare_ResultlessFieldDrawnDeletedAndInserted_KeepsOneCopyInEachView(string kind, string edit)
    {
        var field = ResultlessField(kind);
        var (left, right) = edit == "text on both sides replaced"
            ? (Para(R("alpha ") + field + R("beta")), Para(R("gamma ") + field + R("delta")))
            : (Para(R("First half of text ") + field + R("second half of text")),
               Para(R("First half of text ") + field) + Para(R("second half of text")));
        var original = IrTestDocuments.FromBodyXml(left);
        var revised = IrTestDocuments.FromBodyXml(right);

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;

        Assert.Equal(Shape(revised.DocumentByteArray), Shape(DocxDiffOps.AcceptRevisions(redline)));
        Assert.Equal(Shape(original.DocumentByteArray), Shape(DocxDiffOps.RejectRevisions(redline)));
    }

    private static string Shape(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var sb = new StringBuilder();
        foreach (var e in document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!.Descendants())
        {
            if (e.Name == W.t)
                sb.Append(e.Value);
            else if (e.Name == W.fldChar)
                sb.Append('{').Append((string?)e.Attribute(W.fldCharType)).Append('}');
            else if (e.Name == W.instrText)
                sb.Append('<').Append(e.Value).Append('>');
            else if (e.Name == W.p)
                sb.Append('¶');
        }
        return sb.ToString();
    }
}
