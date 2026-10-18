using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
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
    private static string PageField(string result) => F("begin") + I("PAGE") + F("separate") + R(result) + F("end");

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

    [Theory]
    [InlineData("result", true, true)]
    [InlineData("result", true, false)]
    [InlineData("delete-prefix", true, true)]
    [InlineData("delete-prefix", true, false)]
    [InlineData("result", false, true)]
    [InlineData("result", false, false)]
    [InlineData("delete-prefix", false, true)]
    [InlineData("delete-prefix", false, false)]
    public void Consolidate_ReviewersEditingBesideAField_PreserveItsCharacters(string edit, bool field, bool frontDoor)
    {
        WmlDocument Document(string prefix, string result, string suffix) => IrTestDocuments.FromBodyXml(Para(
            R(prefix) + (field ? PageField(result) : R(result)) + R(suffix)));
        var original = Document("Alder birch ", "1", " tail");
        var prefix = edit == "delete-prefix" ? "Alder " : "Alder birch ";
        var result = edit == "result" ? "7" : "1";
        var reviewers = new[]
        {
            new DocxDiffReviewer { Author = "Cedar", Document = Document(prefix, result, " tail") },
            new DocxDiffReviewer { Author = "Willow", Document = Document("Alder birch ", "1", " tail more") },
        };
        var redline = (frontDoor ? DocxCompare.Consolidate(original, reviewers)
            : DocxDiff.Consolidate(original, reviewers)).DocumentByteArray;

        AssertViews(redline, original, Document(prefix, result, " tail more"), reviewers.Select(reviewer => reviewer.Document).ToArray());
    }

    [Theory]
    [InlineData("shared", false)]
    [InlineData("shared", true)]
    [InlineData("added", false)]
    [InlineData("added", true)]
    [InlineData("removed", false)]
    [InlineData("removed", true)]
    public void Consolidate_FieldBoundaryBookmarksKeepTheirViewAndOrder(string mode, bool table)
    {
        WmlDocument Document(bool bookmark, string result, string suffix)
        {
            var paragraph = Para(R("Alder birch ")
                + (bookmark ? "<w:bookmarkStart w:id=\"5\" w:name=\"CedarRange\"/>" : "") + PageField(result)
                + (bookmark ? "<w:bookmarkEnd w:id=\"5\"/>" : "") + R(suffix));
            return IrTestDocuments.FromBodyXml(table
                ? "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc>" + paragraph + "</w:tc></w:tr></w:tbl>"
                : paragraph);
        }
        var original = Document(mode != "added", "1", " tail");
        var reviewer = Document(mode != "removed", "7", " tail");
        var other = Document(mode != "added", "1", " tail more");
        var redline = DocxCompare.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Author = "Cedar", Document = reviewer },
            new DocxDiffReviewer { Author = "Willow", Document = other },
        }).DocumentByteArray;
        AssertViews(redline, original, Document(mode != "removed", "7", " tail more"), reviewer, other);
    }

    [Theory]
    [InlineData("ordinary")]
    [InlineData("nested-result")]
    [InlineData("nested-hyperlink")]
    [InlineData("adjacent")]
    [InlineData("hyperlink")]
    public void Consolidate_ThreeReviewersPreserveFieldsOrReportCarrierConflict(string kind)
    {
        WmlDocument Document(string prefix, string result, string suffix)
        {
            var field = kind switch
            {
                "nested-result" => F("begin") + I("REF Outer") + F("separate") + R("cedar ") + PageField(result) + R(" fir") + F("end"),
                "nested-hyperlink" => F("begin") + I("HYPERLINK \"https://example.test/cedar\"") + F("separate") + R("cedar ") + PageField(result) + R(" fir") + F("end"),
                "adjacent" => PageField(result) + R(" ") + PageField("2"),
                "hyperlink" => "<w:hyperlink w:anchor=\"Target\">" + PageField(result) + "</w:hyperlink>",
                _ => PageField(result),
            };
            return IrTestDocuments.FromBodyXml(Para(R(prefix) + field + R(suffix)));
        }
        var original = Document("Alder birch ", "1", " tail");
        var reviewers = new[]
        {
            new DocxDiffReviewer { Author = "Cedar", Document = Document("Alder birch ", "7", " tail") },
            new DocxDiffReviewer { Author = "Willow", Document = Document("Alder birch ", "1", " tail more") },
            new DocxDiffReviewer { Author = "Maple", Document = Document("Root Alder birch ", "1", " tail") },
        };
        var redline = DocxCompare.Consolidate(original, reviewers).DocumentByteArray;
        // A HYPERLINK field is a structural carrier. Its changed cached result still takes the
        // existing whole-paragraph conflict path; the field planner must not bypass that policy.
        bool carrierConflict = kind == "nested-hyperlink";
        AssertViews(redline, original, carrierConflict ? original : Document("Root Alder birch ", "7", " tail more"),
            reviewers.Select(reviewer => reviewer.Document).ToArray());
        var conflicts = DocxDiff.GetConflicts(original, reviewers);
        if (carrierConflict)
        {
            Assert.Single(conflicts);
            return;
        }
        Assert.Empty(conflicts);
        using var stream = new MemoryStream(redline);
        using var package = WordprocessingDocument.Open(stream, false);
        var inserted = package.MainDocumentPart!.GetXDocument().Descendants(W.ins)
            .GroupBy(e => (string?)e.Attribute(W.author) ?? "")
            .ToDictionary(group => group.Key, group => string.Concat(group.SelectMany(e => e.Descendants(W.t)).Select(e => e.Value)));
        Assert.Equal("7", inserted["Cedar"].Trim());
        Assert.Equal("more", inserted["Willow"].Trim());
        Assert.Equal("Root", inserted["Maple"].Trim());
    }

    private static void AssertViews(byte[] redline, WmlDocument original, WmlDocument expected, params WmlDocument[] reviewers)
    {
        var accepted = DocxDiffOps.AcceptRevisions(redline);
        var rejected = DocxDiffOps.RejectRevisions(redline);
        Assert.Equal(Shape(original.DocumentByteArray), Shape(rejected));
        Assert.Equal(Shape(expected.DocumentByteArray), Shape(accepted));
        foreach (var bytes in reviewers.Select(document => document.DocumentByteArray)
            .Concat(new[] { original.DocumentByteArray, expected.DocumentByteArray, redline, accepted, rejected }))
            PackageValidation.AssertValid(bytes, FileFormatVersions.Office2019);
    }

    private static string Shape(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var sb = new StringBuilder();
        var body = document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        var names = body.Descendants(W.bookmarkStart).GroupBy(e => (string?)e.Attribute(W.id) ?? "")
            .ToDictionary(group => group.Key, group => (string?)group.First().Attribute(W.name) ?? "");
        foreach (var e in body.Descendants())
        {
            if (e.Name == W.t)
                sb.Append(e.Value);
            else if (e.Name == W.fldChar)
                sb.Append('{').Append((string?)e.Attribute(W.fldCharType)).Append('}');
            else if (e.Name == W.instrText)
                sb.Append('<').Append(e.Value).Append('>');
            else if (e.Name == W.p)
                sb.Append('¶');
            else if (e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd)
                sb.Append('[').Append(e.Name.LocalName).Append(':')
                    .Append(names.GetValueOrDefault((string?)e.Attribute(W.id) ?? "", "missing")).Append(']');
        }
        return sb.ToString();
    }
}
