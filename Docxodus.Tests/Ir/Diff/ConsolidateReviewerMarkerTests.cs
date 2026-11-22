using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// A reviewer who only adds a comment or a bookmark keeps it when another reviewer edits the same paragraph or
/// table, and a composed table keeps its non-row children where they were (issue #1812). Each case consolidates a
/// base with an editor and a marker-only reviewer, in both orders and under every conflict policy: accepting gives
/// the base with every reviewer's change, rejecting gives the base, no conflict is recorded, and every view is
/// schema-valid.
/// </summary>
public class ConsolidateReviewerMarkerTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string Run(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static string P(params string[] content) =>
        "<w:p>" + string.Concat(content.Select(c => c.StartsWith("<w:") ? c : Run(c))) + "</w:p>";

    /// <summary>A comment around <paramref name="text"/>, its reference right after the range, as Word writes it.</summary>
    private static string Commented(string text) =>
        "<w:commentRangeStart w:id=\"0\"/>" + Run(text) + "<w:commentRangeEnd w:id=\"0\"/>" +
        "<w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:commentReference w:id=\"0\"/></w:r>";

    private static string Marked(string text) =>
        "<w:bookmarkStart w:id=\"6\" w:name=\"Shelf\"/>" + Run(text) + "<w:bookmarkEnd w:id=\"6\"/>";

    private static WmlDocument Doc(string body) => IrTestDocuments.FromBodyXml(body);

    private static WmlDocument CommentDoc(string body) =>
        IrTestDocuments.WithComment("Reviewer B", "B", "2026-03-01T00:00:00Z", "Confirm the loan period", body);

    private static string Cell(string content) =>
        $"<w:tc><w:tcPr><w:tcW w:w=\"2600\" w:type=\"dxa\"/></w:tcPr>{(content.StartsWith("<w:p") ? content : P(content))}</w:tc>";

    private static string Row(params string[] cells) => "<w:tr>" + string.Concat(cells.Select(Cell)) + "</w:tr>";

    private static string Table(string tblPr, params string[] children) =>
        $"<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/>{tblPr}</w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"2600\"/><w:gridCol w:w=\"2600\"/></w:tblGrid>" +
        string.Concat(children) + "</w:tbl>" + P("The catalogue closes here.");

    private static XElement Body(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!;
    }

    /// <summary>
    /// The body in order: each paragraph as its text with <c>[</c>, <c>]</c> and <c>^</c> for a comment's range start,
    /// end and reference and <c>{name</c>, <c>}</c> for a bookmark's; each table, row, cell and content control as its
    /// children, so a marker or a control that moves, is lost or is added changes it. Comment ids are left out: the
    /// merge renumbers them.
    /// </summary>
    private static string View(WmlDocument doc) => string.Join(" ", Body(doc).Elements()
        .Where(e => e.Name != W + "sectPr").Select(Node));

    private static string Node(XElement e) => e.Name.LocalName switch
    {
        "p" => "p(" + string.Concat(e.Descendants().Select(Inline)) + ")",
        "tbl" or "tr" or "tc" => e.Name.LocalName + "(" + string.Join(",", e.Elements()
            .Where(c => !c.Name.LocalName.EndsWith("Pr") && c.Name != W + "tblGrid" && c.Name != W + "tblPrEx").Select(Node)) + ")",
        "sdt" => "sdt(" + string.Join(",", e.Element(W + "sdtContent")?.Elements().Select(Node) ?? Enumerable.Empty<string>()) + ")",
        _ => Inline(e),
    };

    private static string Inline(XElement e) => e.Name.LocalName switch
    {
        "t" => e.Value,
        "commentRangeStart" => "[",
        "commentRangeEnd" => "]",
        "commentReference" => "^",
        "bookmarkStart" => "{" + e.Attribute(W + "name")?.Value,
        "bookmarkEnd" => "}",
        _ => "",
    };

    private static string[] CommentTexts(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        var referenced = Body(doc).Descendants(W + "commentReference").Select(r => r.Attribute(W + "id")!.Value).ToHashSet();
        return package.MainDocumentPart!.WordprocessingCommentsPart?.GetXDocument().Root!.Elements(W + "comment")
            .Where(c => referenced.Contains(c.Attribute(W + "id")!.Value))
            .Select(c => c.Value).ToArray() ?? System.Array.Empty<string>();
    }

    private static void AssertValid(params WmlDocument[] docs)
    {
        foreach (var doc in docs)
        {
            using var stream = new MemoryStream(doc.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(package));
        }
    }

    private static WmlDocument PaymentTerm(bool controlled, string term, string? annotation, bool inside = false, bool partial = false)
    {
        string Annotate(string text) => annotation switch
        {
            "comment" => Commented(text),
            "bookmark" => Marked(text),
            _ => Run(text),
        };
        string value = inside && partial && annotation is not null
            ? Annotate(term[..term.IndexOf(' ')]) + Run(term[term.IndexOf(' ')..])
            : inside ? Annotate(term) : Run(term);
        if (controlled)
            value = "<w:sdt><w:sdtPr><w:alias w:val='Term'/></w:sdtPr><w:sdtContent>" + value + "</w:sdtContent></w:sdt>";
        string xml = P(inside ? Run("Payment") : Annotate("Payment"), " is due within ", value, " of the invoice date.") +
            P("Late payments accrue interest.");
        return annotation == "comment" ? CommentDoc(xml) : Doc(xml);
    }

    [Theory]
    [InlineData(true, "comment", false)]
    [InlineData(true, "bookmark", false)]
    [InlineData(true, "comment", true)]
    [InlineData(true, "bookmark", true)]
    [InlineData(true, "comment", true, true)]
    [InlineData(true, "bookmark", true, true)]
    [InlineData(false, "comment", false)]
    [InlineData(false, "bookmark", false)]
    public void An_annotation_keeps_the_paragraph_and_its_inline_control(bool controlled, string annotation, bool inside, bool partial = false)
    {
        var original = PaymentTerm(controlled, "thirty days", null, inside, partial);
        var reviewed = PaymentTerm(controlled, "thirty days", annotation, inside, partial);
        Assert.Empty(DocxDiff.GetRevisions(original, reviewed));
        var changes = DocxDiff.GetSemanticChanges(original, reviewed).Changes;
        Assert.Contains(changes, c => c.Family == (annotation == "comment" ? SemanticChangeFamily.Comment : SemanticChangeFamily.Bookmark));
        Assert.DoesNotContain(changes, c => c.Family is SemanticChangeFamily.Text or SemanticChangeFamily.BlockStructure or SemanticChangeFamily.ContentControl);
        var reviewers = new[] { new DocxDiffReviewer { Author = "Reviewer B", Document = reviewed } };
        Assert.Empty(DocxDiff.GetConsolidatedRevisions(original, reviewers));
        foreach (var (product, redline) in new[] { ("Compare", DocxCompare.Compare(original, reviewed)), ("Consolidate", DocxCompare.Consolidate(original, reviewers)) })
        {
            var accepted = RevisionProcessor.AcceptRevisions(redline);
            var rejected = RevisionProcessor.RejectRevisions(redline);
            Assert.DoesNotContain(Body(redline).Descendants(), e => e.Name == W + "delText" ||
                (e.Name == W + "ins" && e.Descendants(W + "t").Any()));
            Assert.Equal(View(reviewed), View(accepted));
            Assert.True(View(original) == View(rejected), $"Rejecting {product} gave {View(rejected)}");
            Assert.All(new[] { redline, accepted, rejected }, doc =>
                Assert.Equal(controlled ? 1 : 0, Body(doc).Descendants(W + "sdt").Count()));
            if (annotation == "comment")
                Assert.Equal(CommentTexts(reviewed), CommentTexts(accepted));
            AssertValid(redline, accepted, rejected);
        }
    }

    [Theory]
    [MemberData(nameof(Orders))]
    public void An_annotation_beside_an_edited_inline_control_composes(ConflictResolution policy, bool editorFirst)
    {
        var original = PaymentTerm(true, "thirty days", null);
        var editor = PaymentTerm(true, "forty days", null);
        foreach (var annotation in new[] { "bookmark", "comment" })
        {
            var reviewer = PaymentTerm(true, "thirty days", annotation);
            var expected = PaymentTerm(true, "forty days", annotation);
            var reviewers = InOrder(editorFirst, editor, reviewer);
            var accepted = AssertMerged(original, expected, policy, reviewers);
            var revisions = DocxDiff.GetConsolidatedRevisions(original, reviewers, new DocxDiffConsolidateSettings { ConflictResolution = policy });
            Assert.NotEmpty(revisions);
            Assert.All(revisions, revision => Assert.Equal("Reviewer A", revision.Author));
            Assert.Equal("forty days", Assert.Single(Body(accepted).Descendants(W + "sdtContent")).Value);
            Assert.Equal(CommentTexts(expected), CommentTexts(accepted));
        }
    }

    [Theory]
    [InlineData("hyphen", "comment")]
    [InlineData("hyphen", "bookmark")]
    [InlineData("textbox", "comment")]
    [InlineData("textbox", "bookmark")]
    [InlineData("mixed", "comment")]
    [InlineData("mixed", "bookmark")]
    [InlineData("plain", "collapsed")]
    [InlineData("plain", "comment")]
    public void Replacement_ranges_use_the_winning_paragraphs_source_boundaries(string prefix, string annotation)
    {
        var leading = prefix switch
        {
            "hyphen" => "<w:r><w:t>Non</w:t><w:noBreakHyphen/><w:t>cash </w:t></w:r>",
            "textbox" => "<w:r><w:pict><v:shape xmlns:v='urn:schemas-microsoft-com:vml' id='box'><v:textbox>" +
                         "<w:txbxContent>" + P("Separate text stream") + "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>",
            _ => Run("Cash "),
        };
        WmlDocument Make(string term, string? mark)
        {
            string payment = mark switch
            {
                "comment" => Commented("Payment"),
                "bookmark" => Marked("Payment"),
                "collapsed" => "<w:bookmarkStart w:id='6' w:name='Shelf'/><w:bookmarkEnd w:id='6'/>" + Run("Payment"),
                _ => Run("Payment"),
            };
            var beforeTerm = prefix == "mixed" && mark is null
                ? "<w:r><w:t>Cash </w:t><w:softHyphen/><w:t>Invoices due. Payment is due within </w:t></w:r>"
                : (prefix == "mixed" ? Run("Cash ") + "<w:r><w:softHyphen/></w:r>" : leading) +
                  Run("Invoices due. ") + payment + Run(" is due within ");
            var body = P(beforeTerm,
                "<w:sdt><w:sdtPr><w:alias w:val='Term'/></w:sdtPr><w:sdtContent>" + Run(term) + "</w:sdtContent></w:sdt>",
                " of the invoice date.") + P("Late payments accrue interest.");
            return mark == "comment" ? CommentDoc(body) : Doc(body);
        }
        var original = Make("thirty days", null);
        var expected = Make("forty days", annotation);
        foreach (var order in Orders())
        {
            var reviewers = InOrder((bool)order[1], Make("forty days", null), Make("thirty days", annotation));
            var settings = new DocxDiffConsolidateSettings { ConflictResolution = (ConflictResolution)order[0] };
            var raw = DocxCompare.Consolidate(original, reviewers, settings);
            var accepted = RevisionProcessor.AcceptRevisions(raw);
            var rejected = RevisionProcessor.RejectRevisions(raw);
            Assert.Equal(View(expected), View(accepted));
            Assert.Equal(View(original), View(rejected));
            Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
            AssertValid(raw, accepted, rejected);
            if (annotation == "comment")
            {
                AssertRangesInOrder(raw, accepted);
                Assert.Equal(CommentTexts(expected), CommentTexts(accepted));
            }
            else
            {
                foreach (var doc in new[] { raw, accepted })
                {
                    var endpoints = Body(doc).Descendants().Where(e => e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd").ToList();
                    Assert.Equal(new[] { W + "bookmarkStart", W + "bookmarkEnd" }, endpoints.Select(e => e.Name));
                    Assert.Equal(endpoints[0].Attribute(W + "id")!.Value, endpoints[1].Attribute(W + "id")!.Value);
                }
            }
            Assert.All(DocxDiff.GetConsolidatedRevisions(original, reviewers, settings),
                revision => Assert.Equal("Reviewer A", revision.Author));
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void A_replaced_note_paragraph_keeps_its_references_defined(bool footnote)
    {
        string kind = footnote ? "footnote" : "endnote";
        string body = P("The register lists every loan.", "<w:r><w:" + kind + "Reference w:id='1'/></w:r>");
        WmlDocument Make(string term, string next, bool comment)
        {
            string note = "<w:" + kind + " w:id='1'>" + P(comment ? Commented("Payment") : Run("Payment"),
                " is due within ", "<w:sdt><w:sdtPr><w:alias w:val='Term'/></w:sdtPr><w:sdtContent>" +
                Run(term) + "</w:sdtContent></w:sdt>", " of the invoice date.") + P(next) + "</w:" + kind + ">";
            return NoteDoc(body, note, comment ? "<w:comment w:id='0' w:author='Reviewer B'>" + P("Check the term.") +
                "</w:comment>" : null, footnote);
        }
        static XElement Notes(WmlDocument doc, bool footnote)
        {
            using var stream = new MemoryStream(doc.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            var main = package.MainDocumentPart!;
            return (footnote ? (OpenXmlPart)main.FootnotesPart! : main.EndnotesPart!).GetXDocument().Root!;
        }
        var original = Make("thirty days", "The deposit is optional.", false);
        var editor = Make("forty days", "The deposit is optional.", false);
        var marker = Make("thirty days", "The deposit is required.", true);
        foreach (var order in Orders())
        {
            var reviewers = InOrder((bool)order[1], editor, marker);
            var settings = new DocxDiffConsolidateSettings { ConflictResolution = (ConflictResolution)order[0] };
            var merged = DocxCompare.Consolidate(original, reviewers, settings);
            var accepted = RevisionProcessor.AcceptRevisions(merged);
            var rejected = RevisionProcessor.RejectRevisions(merged);
            foreach (var view in new[] { merged, accepted, rejected })
                Assert.Empty(Notes(view, footnote).Descendants(W + "commentReference"));
            Assert.Equal(Notes(Make("forty days", "The deposit is required.", false), footnote).Value, Notes(accepted, footnote).Value);
            Assert.Equal(Notes(original, footnote).Value, Notes(rejected, footnote).Value);
            Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
            AssertValid(merged, accepted, rejected);
        }
    }

    [Theory]
    [MemberData(nameof(Orders))]
    public void A_consensus_bookmark_keeps_its_selected_reviewer_on_a_replacement(ConflictResolution policy, bool editorFirst)
    {
        var original = PaymentTerm(true, "thirty days", null);
        var expected = PaymentTerm(true, "forty days", "bookmark");
        var reviewers = InOrder(editorFirst, expected, PaymentTerm(true, "thirty days", "bookmark"));
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
        var raw = DocxCompare.Consolidate(original, reviewers, settings);
        var accepted = RevisionProcessor.AcceptRevisions(raw);
        var rejected = RevisionProcessor.RejectRevisions(raw);
        Assert.Equal(View(expected), View(accepted));
        Assert.Equal(View(original), View(rejected));
        var start = Assert.Single(Body(raw).Descendants(W + "bookmarkStart"));
        Assert.Equal(reviewers[0].Author, start.Ancestors(W + "ins").First().Attribute(W + "author")!.Value);
        Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
        AssertValid(raw, accepted, rejected);
    }

    [Theory]
    [InlineData("tab")]
    [InlineData("picture")]
    public void A_range_around_an_unchanged_atomic_run_keeps_its_contents(string shape)
    {
        string atom = shape == "tab" ? "<w:r><w:tab/></w:r>" :
            "<w:r><w:pict><v:shape xmlns:v='urn:schemas-microsoft-com:vml' id='pic' style='width:10pt;height:10pt'>" +
            "<v:imagedata r:id='rIdPic'/></v:shape></w:pict></w:r>";
        WmlDocument Make(string term, bool marked)
        {
            string value = marked ? "<w:bookmarkStart w:id='6' w:name='Shelf'/>" + atom + "<w:bookmarkEnd w:id='6'/>" : atom;
            string body = P("Cash ", value, "Payment is due within ", "<w:sdt><w:sdtPr><w:alias w:val='Term'/></w:sdtPr><w:sdtContent>" +
                Run(term) + "</w:sdtContent></w:sdt>", " of the invoice date.") + P("Late payments accrue interest.");
            return shape == "tab" ? Doc(body) : IrTestDocuments.FromBodyXmlWithImageParts(body, ("rIdPic", IrTestDocuments.TinyPng));
        }
        var original = Make("thirty days", false);
        foreach (var order in Orders())
        {
            var reviewers = InOrder((bool)order[1], Make("forty days", false), Make("thirty days", true));
            var accepted = AssertMerged(original, Make("forty days", true), (ConflictResolution)order[0], reviewers);
            var marks = Body(accepted).Descendants().Where(e => e.Name == W + "bookmarkStart" ||
                e.Name == W + "bookmarkEnd" || e.Name == W + (shape == "tab" ? "tab" : "pict"));
            Assert.Equal(new[] { "bookmarkStart", shape == "tab" ? "tab" : "pict", "bookmarkEnd" }, marks.Select(e => e.Name.LocalName));
        }
    }

    [Theory]
    [InlineData("two-texts")]
    [InlineData("tab-between")]
    [InlineData("plain")]
    public void Annotation_only_run_divisions_keep_the_control_envelope(string shape)
    {
        string tab = shape == "tab-between" ? "<w:tab/>" : "";
        string left = shape == "plain" ? Run("Alpha beta") :
            "<w:r><w:t xml:space='preserve'>Alpha </w:t>" + tab + "<w:t>beta</w:t></w:r>";
        string right = "<w:r><w:t xml:space='preserve'>Alpha </w:t>" + tab + "</w:r>" + Marked("beta");
        WmlDocument Make(string contents) => Doc(P("Invoice ", "<w:sdt><w:sdtPr><w:alias w:val='Term'/></w:sdtPr><w:sdtContent>" +
            contents + "</w:sdtContent></w:sdt>", " ends here."));
        var original = Make(left);
        var reviewer = Make(right);
        var reviewers = new[] { new DocxDiffReviewer { Author = "B", Document = reviewer } };
        Assert.Empty(DocxDiff.GetRevisions(original, reviewer));
        Assert.Empty(DocxDiff.GetConsolidatedRevisions(original, reviewers));
        var changes = DocxDiff.GetSemanticChanges(original, reviewer).Changes;
        Assert.Contains(changes, c => c.Family == SemanticChangeFamily.Bookmark);
        Assert.DoesNotContain(changes, c => c.Family is SemanticChangeFamily.Text or SemanticChangeFamily.BlockStructure or SemanticChangeFamily.ContentControl);
        foreach (var raw in new[] { DocxCompare.Compare(original, reviewer), DocxCompare.Consolidate(original, reviewers) })
        {
            var accepted = RevisionProcessor.AcceptRevisions(raw);
            var rejected = RevisionProcessor.RejectRevisions(raw);
            Assert.Single(Body(raw).Descendants(W + "sdt"));
            Assert.Equal(View(reviewer), View(accepted));
            Assert.Equal(View(original), View(rejected));
            AssertValid(raw, accepted, rejected);
        }
    }

    public static IEnumerable<object[]> Orders() =>
        from policy in new[] { ConflictResolution.BaseWins, ConflictResolution.FirstReviewerWins, ConflictResolution.StackAll }
        from editorFirst in new[] { true, false }
        select new object[] { policy, editorFirst };

    [Theory]
    [InlineData("before", false)]
    [InlineData("around", false)]
    [InlineData("after", false)]
    [InlineData("before", true)]
    [InlineData("around", true)]
    [InlineData("after", true)]
    public void An_annotation_only_range_keeps_its_side_of_a_tab_in_a_mixed_run(string extent, bool comment)
    {
        string start = comment ? "<w:commentRangeStart w:id='0'/>" : "<w:bookmarkStart w:id='8' w:name='Shelf'/>";
        string end = comment ? "<w:commentRangeEnd w:id='0'/><w:r><w:commentReference w:id='0'/></w:r>" : "<w:bookmarkEnd w:id='8'/>";
        const string tab = "<w:r><w:tab/></w:r>";
        string Wrap(string payload) => "<w:sdt><w:sdtPr><w:alias w:val='Term'/></w:sdtPr><w:sdtContent>" +
            payload + "</w:sdtContent></w:sdt>";
        var original = Doc(P(Wrap("<w:r><w:t>Alpha</w:t><w:tab/><w:t>beta</w:t></w:r>")));
        string body = P(Wrap(extent switch
        {
            "before" => start + Run("Alpha") + end + tab + Run("beta"),
            "around" => Run("Alpha") + start + tab + end + Run("beta"),
            _ => Run("Alpha") + tab + start + Run("beta") + end,
        }));
        var marked = comment ? CommentDoc(body) : Doc(body);
        var reviewers = new[] { new DocxDiffReviewer { Author = "B", Document = marked } };
        Assert.Empty(DocxDiff.GetRevisions(original, marked));
        Assert.Empty(DocxDiff.GetConsolidatedRevisions(original, reviewers));
        Assert.Empty(DocxDiff.GetConflicts(original, reviewers));
        Assert.Contains(DocxDiff.GetSemanticChanges(original, marked).Changes, c => c.Family == (comment ? SemanticChangeFamily.Comment : SemanticChangeFamily.Bookmark));
        foreach (var raw in new[] { DocxCompare.Compare(original, marked), DocxCompare.Consolidate(original, reviewers) })
        {
            var accepted = RevisionProcessor.AcceptRevisions(raw);
            var rejected = RevisionProcessor.RejectRevisions(raw);
            Assert.Equal(MixedControlLeaves(marked), MixedControlLeaves(raw));
            Assert.Equal(MixedControlLeaves(marked), MixedControlLeaves(accepted));
            Assert.Equal(MixedControlLeaves(original), MixedControlLeaves(rejected));
            AssertValid(raw, accepted, rejected);
        }
    }

    private static string[] MixedControlLeaves(WmlDocument file) => Body(file).Descendants(W + "sdtContent").Single().Descendants()
        .Where(e => e.Name == W + "t" || e.Name == W + "tab" || e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd" ||
            e.Name == W + "commentRangeStart" || e.Name == W + "commentRangeEnd" || e.Name == W + "commentReference")
        .Select(e => e.Name == W + "t" ? e.Value : e.Name.LocalName).ToArray();

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    public void A_carried_range_keeps_its_side_of_a_tab_beside_another_reviewers_edit(bool comment, bool formatOnly, bool editBefore)
    {
        const string mixed = "<w:r><w:t>Alpha</w:t><w:tab/><w:t>beta</w:t></w:r>";
        string marked = (comment ? Commented("Alpha") : Marked("Alpha")) + "<w:r><w:tab/><w:t>beta</w:t></w:r>";
        string Paragraph(string content, bool edited) => P(edited && editBefore ? "The new catalogue now lists " : "The catalogue lists ",
            "<w:sdt><w:sdtPr><w:alias w:val='Term'/></w:sdtPr><w:sdtContent>" + content + "</w:sdtContent></w:sdt>",
            edited && formatOnly ? "<w:r><w:rPr><w:b/></w:rPr><w:t xml:space='preserve'> on the old shelf.</w:t></w:r>"
            : edited && !editBefore ? " on the new shelf." : " on the old shelf.") + P("Returns are due on Fridays.");
        var original = Doc(Paragraph(mixed, false));
        var editor = Doc(Paragraph(mixed, true));
        var reviewer = comment ? CommentDoc(Paragraph(marked, false)) : Doc(Paragraph(marked, false));
        foreach (var editorFirst in new[] { false, true })
        {
            var reviewers = InOrder(editorFirst, editor, reviewer);
            var raw = DocxCompare.Consolidate(original, reviewers);
            var accepted = RevisionProcessor.AcceptRevisions(raw);
            var rejected = RevisionProcessor.RejectRevisions(raw);
            Assert.Equal(MixedControlLeaves(reviewer), MixedControlLeaves(raw));
            Assert.Equal(MixedControlLeaves(reviewer), MixedControlLeaves(accepted));
            Assert.Equal(MixedControlLeaves(original), MixedControlLeaves(rejected));
            Assert.Equal(string.Concat(Body(editor).Descendants(W + "t").Select(e => e.Value)),
                string.Concat(Body(accepted).Descendants(W + "t").Select(e => e.Value)));
            Assert.Equal(formatOnly, Body(accepted).Descendants(W + "b").Any());
            Assert.Empty(DocxDiff.GetConflicts(original, reviewers));
            Assert.All(DocxDiff.GetConsolidatedRevisions(original, reviewers), r => Assert.Equal("Reviewer A", r.Author));
            if (comment) Assert.Equal(CommentTexts(reviewer), CommentTexts(accepted));
            AssertValid(raw, accepted, rejected);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void A_range_on_replaced_text_keeps_its_affinity_after_an_atomic_run(bool textbox, bool comment)
    {
        string prefix = textbox
            ? "<w:r><w:pict><v:shape xmlns:v='urn:schemas-microsoft-com:vml' id='box'><v:textbox><w:txbxContent>" +
              P("Separate text stream") + "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>"
            : "<w:r><w:t>Index</w:t><w:tab/></w:r>";
        WmlDocument Make(string word, bool mark) => mark && comment
            ? CommentDoc(P(prefix, "The shelf is ", Commented(word), " today."))
            : Doc(P(prefix, "The shelf is ", mark ? Marked(word) : Run(word), " today."));
        var original = Make("closed", false);
        var editor = Make("open", false);
        var reviewer = Make("closed", true);
        var expected = Make("open", true);
        foreach (var policy in new[] { ConflictResolution.BaseWins, ConflictResolution.FirstReviewerWins, ConflictResolution.StackAll })
        foreach (var editorFirst in new[] { false, true })
        {
            var reviewers = InOrder(editorFirst, editor, reviewer);
            var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
            var raw = DocxCompare.Consolidate(original, reviewers, settings);
            var accepted = RevisionProcessor.AcceptRevisions(raw);
            var rejected = RevisionProcessor.RejectRevisions(raw);
            Assert.Equal(View(expected), View(raw));
            Assert.Equal(View(expected), View(accepted));
            Assert.Equal(View(original), View(rejected));
            Assert.Single(Body(accepted).Descendants(W + (textbox ? "txbxContent" : "tab")));
            Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
            if (comment) Assert.Equal(CommentTexts(reviewer), CommentTexts(accepted));
            AssertValid(raw, accepted, rejected);
        }
    }

    /// <summary>Consolidate the base with <paramref name="reviewers"/> under <paramref name="policy"/>: accepting gives
    /// <paramref name="expected"/>, rejecting the base, nothing is recorded as a conflict and every view is valid.</summary>
    private static WmlDocument AssertMerged(WmlDocument original, WmlDocument expected, ConflictResolution policy,
        params DocxDiffReviewer[] reviewers)
    {
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
        var merged = DocxCompare.Consolidate(original, reviewers, settings);
        var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
        Assert.Equal(View(expected), View(accepted));
        Assert.Equal(View(original), View(rejected));
        Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
        AssertValid(merged, accepted, rejected);
        return accepted;
    }

    private static DocxDiffReviewer[] InOrder(bool editorFirst, WmlDocument editor, WmlDocument marker) => editorFirst
        ? new[] { new DocxDiffReviewer { Document = editor, Author = "Reviewer A" }, new DocxDiffReviewer { Document = marker, Author = "Reviewer B" } }
        : new[] { new DocxDiffReviewer { Document = marker, Author = "Reviewer B" }, new DocxDiffReviewer { Document = editor, Author = "Reviewer A" } };

    /// <summary>
    /// Reviewer B comments on "lists every" in a paragraph whose text reviewer A edits. <c>before</c> and <c>after</c>:
    /// A adds words before or after the comment; <c>inside</c>: A replaces "every", so the comment covers the
    /// replacement; <c>at-end</c> and <c>at-start</c>: A adds a word right after or right before the commented words,
    /// which the comment does not take.
    /// </summary>
    [Theory]
    [MemberData(nameof(Shapes))]
    public void A_comment_beside_another_reviewers_edit_in_the_same_paragraph_is_kept(string shape, ConflictResolution policy, bool editorFirst)
    {
        var (edited, expected) = shape switch
        {
            "before" => ("The new register lists every loan.", P("The new register ", Commented("lists every"), " loan.")),
            "after" => ("The register lists every loan, with dates.", P("The register ", Commented("lists every"), " loan, with dates.")),
            "inside" => ("The register lists each loan.", P("The register ", Commented("lists each"), " loan.")),
            "at-end" => ("The register lists every old loan.", P("The register ", Commented("lists every"), " old loan.")),
            "end-at-insertion" => ("The register lists every old loan.", P("The register ", Commented("lists every "), "old loan.")),
            _ => ("The register now lists every loan.", P("The register now ", Commented("lists every"), " loan.")),
        };
        const string next = "Returns are due on Fridays.";
        var original = Doc(P("The register lists every loan.") + P(next));
        var editor = Doc(P(edited) + P(next));
        // end-at-insertion: the comment's end, after "every ", is exactly where the editor's "old " goes in.
        var marker = CommentDoc((shape == "end-at-insertion" ? P("The register ", Commented("lists every "), "loan.")
            : P("The register ", Commented("lists every"), " loan.")) + P(next));
        var accepted = AssertMerged(original, CommentDoc(expected + P(next)), policy, InOrder(editorFirst, editor, marker));
        Assert.Equal(new[] { "Confirm the loan period" }, CommentTexts(accepted));
    }

    public static IEnumerable<object[]> Shapes() =>
        from shape in new[] { "before", "after", "inside", "at-end", "end-at-insertion", "at-start" }
        from order in Orders()
        select new object[] { shape, order[0], order[1] };

    /// <summary>
    /// Reviewer A deletes, moves or replaces what reviewer B comments on. <c>words-deleted</c>: A deletes the commented
    /// words, and B's comment stays on them in the redline; <c>replaced-whole</c>: A changes a field, which redlines the
    /// paragraph as a deleted and an inserted copy with the same text, and the comment goes on the inserted copy, once.
    /// <c>rewritten-whole</c> changes the field and adjacent text while preserving the commented words.
    /// <c>paragraph-deleted</c>, <c>covered-words-rewritten</c> and <c>moved</c>: no
    /// copy that survives accepting holds the commented text as written, so the comment is not carried, and accepting
    /// gives A's document exactly, with no stray marker or paragraph. <c>editor-comments-too</c> is a control: A comments
    /// on another paragraph and edits none, which main already merged. In every case rejecting gives the base and each
    /// comment's range starts before it ends.
    /// </summary>
    [Theory]
    [InlineData("words-deleted")]
    [InlineData("replaced-whole")]
    [InlineData("paragraph-deleted")]
    [InlineData("rewritten-whole")]
    [InlineData("covered-words-rewritten")]
    [InlineData("moved")]
    [InlineData("editor-comments-too")]
    public void A_comment_on_text_another_reviewer_deletes_moves_or_replaces(string shape)
    {
        const string field = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> XE \"loan\" </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";
        const string commented = "The register lists every loan of the winter term to the members of the society.";
        const string next = "Returns are due on Fridays.";
        const string last = "The reading room closes at six in the evening.";
        var original = Doc(P(commented) + P(next) + P(last));
        var editor = shape switch
        {
            "words-deleted" => Doc(P("The register loan of the winter term to the members of the society.") + P(next) + P(last)),
            "replaced-whole" => Doc(P(commented, field) + P(next) + P(last)),
            "paragraph-deleted" => Doc(P(next) + P(last)),
            "covered-words-rewritten" => Doc(P("The ledger counts each loan of the winter term to the members of the society.", field) + P(next) + P(last)),
            "rewritten-whole" => Doc(P("The ledger lists every loan of the winter term to the members of the society.", field) + P(next) + P(last)),
            "moved" => Doc(P(next) + P(last) + P(commented)),
            _ => IrTestDocuments.WithComment("Reviewer A", "A", "2026-03-02T00:00:00Z", "Check the due day",
                P(commented) + P(Commented("Returns"), " are due on Fridays.") + P(last)),
        };
        var marker = CommentDoc(P("The register ", Commented("lists every"), " loan of the winter term to the members of the society.") + P(next) + P(last));
        bool carried = shape is "words-deleted" or "replaced-whole" or "rewritten-whole" or "editor-comments-too";
        foreach (var policy in new[] { ConflictResolution.BaseWins, ConflictResolution.StackAll })
            foreach (var reviewers in new[] { InOrder(true, editor, marker), InOrder(false, editor, marker) })
            {
                var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
                var merged = DocxCompare.Consolidate(original, reviewers, settings);
                var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
                Assert.Equal(carried ? shape == "editor-comments-too" ? 2 : 1 : 0, Body(merged).Descendants(W + "commentReference").Count());
                if (carried)
                    Assert.Contains("Confirm the loan period", CommentTexts(merged));
                else
                    Assert.Equal(View(editor), View(accepted));
                if (shape is "replaced-whole" or "rewritten-whole")
                    Assert.Equal(View(CommentDoc(P(shape == "rewritten-whole" ? "The ledger " : "The register ", Commented("lists every"), " loan of the winter term to the members of the society.") +
                        P(next) + P(last))), View(accepted));
                Assert.Equal(View(original), View(rejected));
                Assert.Equal(string.Concat(Body(editor).Descendants(W + "t").Select(t => t.Value)),
                    string.Concat(Body(accepted).Descendants(W + "t").Select(t => t.Value)));
                Assert.Equal(Body(editor).Elements(W + "p").Count(), Body(accepted).Elements(W + "p").Count());
                AssertRangesInOrder(merged, accepted);
                Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
                AssertValid(merged, accepted, rejected);
            }
    }

    /// <summary>
    /// Shapes from the #1812 review where a carried marker meets another reviewer's revisions. <c>bold</c>: reviewer A
    /// makes a word bold and B's comment starts inside it, so the bold run is split (each half keeps a revision id of its
    /// own); <c>replaced-word</c>: A rewrites a word, drawn as the new word inserted before the old one deleted, and B's
    /// comment starts inside the old word, so it covers the new one; <c>bookmark-in-bold</c>: a bookmark in the bold word.
    /// </summary>
    [Theory]
    [InlineData("bold")]
    [InlineData("replaced-word")]
    [InlineData("bookmark-in-bold")]
    public void A_marker_inside_a_word_another_reviewer_reformats_or_replaces_keeps_its_place(string shape)
    {
        string bold(string text) => $"<w:r><w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\">{text}</w:t></w:r>";
        const string next = "Returns are due on Fridays.";
        WmlDocument original, editor, marker, expected;
        if (shape == "replaced-word")
        {
            original = Doc(P("The register lists THree loans.") + P(next));
            editor = Doc(P("The register lists Three loans.") + P(next));
            marker = CommentDoc(P("The register lists TH", Commented("ree"), " loans.") + P(next));
            expected = CommentDoc(P("The register lists ", Commented("Three"), " loans.") + P(next));
        }
        else
        {
            original = Doc(P("Alpha beta gamma.") + P(next));
            editor = Doc(P("Alpha ", bold("beta"), " gamma.") + P(next));
            bool comment = shape == "bold";
            string m(string text) => comment ? Commented(text) : ConsolidateReviewerMarkerTests.Marked(text);
            marker = (comment ? (System.Func<string, WmlDocument>)CommentDoc : Doc)(P("Alpha be", m("ta gam"), "ma.") + P(next));
            expected = (comment ? (System.Func<string, WmlDocument>)CommentDoc : Doc)(P("Alpha be", m("ta gam"), "ma.") + P(next));
        }
        foreach (var order in Orders())
            AssertMerged(original, expected, (ConflictResolution)order[0], InOrder((bool)order[1], editor, marker));
    }

    /// <summary>
    /// A base comment stays where the base has it (issue #1812 review): reviewer A replaces the paragraph whole and drops
    /// the base's comment, while reviewer B leaves it and adds a comment of its own on another word, which accepting keeps
    /// over exactly that word, and the base's comment stays removed.
    /// </summary>
    [Fact]
    public void A_base_comment_another_reviewer_drops_stays_dropped_beside_a_new_comment()
    {
        const string field = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> XE \"loan\" </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";
        static string Second(string text) => "<w:commentRangeStart w:id=\"1\"/>" + Run(text) + "<w:commentRangeEnd w:id=\"1\"/>" +
            "<w:r><w:commentReference w:id=\"1\"/></w:r>";
        WmlDocument Two(string body, bool baseComment, bool second) => TwoComments(body, baseComment, second);
        var original = Two(P("The register ", Commented("lists every"), " loan.") + P("Returns are due on Fridays."), true, false);
        var editor = Doc(P("The register lists every loan.", field) + P("Returns are due on Fridays."));
        var marker = Two(P("The register ", Commented("lists every"), " ", Second("loan"), ".") + P("Returns are due on Fridays."), true, true);
        var expected = Two(P("The register lists every ", Second("loan"), ".", field) + P("Returns are due on Fridays."), false, true);
        foreach (var order in Orders())
        {
            var reviewers = InOrder((bool)order[1], editor, marker);
            var settings = new DocxDiffConsolidateSettings { ConflictResolution = (ConflictResolution)order[0] };
            var merged = DocxCompare.Consolidate(original, reviewers, settings);
            var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
            Assert.Equal(View(expected), View(accepted));
            Assert.Equal(View(original), View(rejected));
            AssertRangesInOrder(merged, accepted, rejected);
            AssertValid(merged, accepted, rejected);
        }
    }

    /// <summary>
    /// A base comment stays where the base has it when a reviewer leaves its text unchanged and anchors it on another
    /// block instead, which another reviewer edits (#1812 review): the reviewer's change to where the comment sits is
    /// not carried, so accepting gives the editing reviewer's document and rejecting gives the base. Before, the merged
    /// comment's range ran from one block into the other. <c>paragraph</c>: the editor changes the next paragraph's
    /// text; <c>replaced-whole</c>: it only adds a field there, so that paragraph is replaced whole with its text
    /// unchanged; <c>table-cell</c>: the next block is a table and the editor edits the cell; <c>word-deleted</c>: the
    /// editor also deletes the commented word, and the comment with it; <c>paragraph-deleted</c>: the editor deletes
    /// the commented paragraph; <c>starts-between-blocks</c>: the comment's range starts before the first paragraph;
    /// <c>same-paragraph</c>: the marker reviewer moves the comment to another word of the same paragraph instead;
    /// <c>textless-anchor</c>: the comment covers a paragraph that holds only a tab, as a figure's does;
    /// <c>textless-anchor-unedited</c>: the same, with the next paragraph left unedited.
    /// </summary>
    [Theory]
    [InlineData("paragraph")]
    [InlineData("replaced-whole")]
    [InlineData("table-cell")]
    [InlineData("word-deleted")]
    [InlineData("paragraph-deleted")]
    [InlineData("starts-between-blocks")]
    [InlineData("same-paragraph")]
    [InlineData("textless-anchor")]
    [InlineData("textless-anchor-unedited")]
    public void A_base_comment_a_reviewer_anchors_on_another_block_stays_where_the_base_has_it(string shape)
    {
        const string field = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> XE \"due\" </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";
        const string end = "<w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r>";
        // The paragraph after the commented one: the marker reviewer's copy comments on "Returns" there.
        string Next(bool comment, bool edited)
        {
            string text = edited && shape is not "replaced-whole" and not "textless-anchor-unedited" ? " are due on Mondays." : " are due on Fridays.";
            string paragraph = comment && shape != "same-paragraph" ? P(Commented("Returns"), text)
                : P("Returns" + text, edited && shape == "replaced-whole" ? field : "");
            return shape == "table-cell" ? Table("", Row(paragraph, "Desk two")) : paragraph;
        }
        string Commented0(bool comment, bool editor) => shape switch
        {
            "word-deleted" when editor => P("Alpha gamma delta."),
            "paragraph-deleted" when editor => "",
            "starts-between-blocks" => (comment ? "<w:commentRangeStart w:id=\"0\"/>" : "") + P("Alpha beta gamma delta.") +
                (comment ? P("Second", end, " line.") : P("Second line.")),
            "same-paragraph" when !comment => P("Alpha beta gamma ", Commented("delta"), "."),
            "textless-anchor" or "textless-anchor-unedited" => comment
                ? "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:tab/></w:r>" + end + "</w:p>" + P("Alpha beta gamma delta.")
                    + (editor && shape == "textless-anchor-unedited" ? P("A figure stands above.") : "")
                : "<w:p><w:r><w:tab/></w:r></w:p>" + P("Alpha beta gamma delta."),
            _ => comment ? P("Alpha ", Commented("beta"), " gamma delta.") : P("Alpha beta gamma delta."),
        };
        var original = CommentDoc(Commented0(true, false) + Next(false, false));
        var editor = CommentDoc(Commented0(true, true) + Next(false, true));
        var marker = CommentDoc(Commented0(false, false) + Next(true, false));
        foreach (var order in Orders())
        {
            var reviewers = InOrder((bool)order[1], editor, marker);
            var settings = new DocxDiffConsolidateSettings { ConflictResolution = (ConflictResolution)order[0] };
            var merged = DocxCompare.Consolidate(original, reviewers, settings);
            var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
            Assert.Equal(View(editor), View(accepted));
            Assert.Equal(View(original), View(rejected));
            AssertRangesInOrder(merged, accepted, rejected);
            AssertValid(merged, accepted, rejected);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Similar_text_elsewhere_does_not_own_a_deleted_comment_anchor(bool tabOnly)
    {
        var source = tabOnly ? P(Commented("beta").Replace(Run("beta"), "<w:r><w:tab/></w:r>"))
            : P("Alpha ", Commented("beta"), " gamma delta.");
        var gap = P("The plates stay on the upper shelf.");
        var duplicate = tabOnly ? P() : P("Alpha beta gamma delta.");
        var target = P("Returns are due on Fridays.");
        var baseTail = P("Invoices go out on Thursdays.");
        var editedTail = P("Invoices go out on Mondays.");
        var movedTarget = P(Commented("Returns"), " are due on Fridays.");
        var original = CommentDoc(source + gap + duplicate + target + baseTail);
        var marker = CommentDoc(gap + duplicate + movedTarget + baseTail);
        var editor = CommentDoc(source + gap + duplicate + target + editedTail);
        var expected = CommentDoc(gap + duplicate + movedTarget + editedTail);
        foreach (var order in Orders())
        {
            var reviewers = InOrder((bool)order[1], editor, marker);
            var settings = new DocxDiffConsolidateSettings { ConflictResolution = (ConflictResolution)order[0] };
            var marked = DocxCompare.Consolidate(original, reviewers, settings);
            var accepted = RevisionProcessor.AcceptRevisions(marked);
            var rejected = RevisionProcessor.RejectRevisions(marked);
            var starts = Body(marked).Descendants(W + "commentRangeStart").ToList();
            Assert.Equal(2, starts.Count);
            foreach (var start in starts)
            {
                var revision = start.Ancestors().First(e => e.Name == W + "ins" || e.Name == W + "del");
                Assert.Equal("Reviewer B", (string?)revision.Attribute(W + "author"));
            }
            Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
            Assert.Equal(View(expected), View(accepted));
            Assert.Equal(View(original), View(rejected));
            AssertRangesInOrder(marked, accepted, rejected);
            AssertValid(marked, accepted, rejected);
        }
    }

    /// <summary>
    /// A reviewer who removes or rewrites a base comment's paragraph and anchors the comment on a block another
    /// reviewer edits does not carry it there (#1812 review): the editing reviewer's copy decides where the base's
    /// comment goes, and a second copy of it in the edit would leave the comment with two ranges, its reference before
    /// its end. Rejecting gives the base and every view's comment ranges come out in order. <c>deleted</c>: the marker
    /// reviewer deletes the commented paragraph; <c>word-added</c>: it adds a word there; <c>word-rewritten</c>: it
    /// rewrites the commented word.
    /// </summary>
    [Theory]
    [InlineData("deleted")]
    [InlineData("word-added")]
    [InlineData("word-rewritten")]
    public void A_base_comment_a_reviewer_re_anchors_on_another_reviewers_edit_is_not_carried_there(string shape)
    {
        const string next = " are due on Fridays.";
        var original = CommentDoc(P("Alpha ", Commented("beta"), " gamma delta.") + P("Returns" + next));
        var editor = CommentDoc(P("Alpha ", Commented("beta"), " gamma delta.") + P("Returns are due on Mondays."));
        string first = shape switch
        {
            "deleted" => "",
            "word-added" => P("Alpha beta gamma delta epsilon."),
            _ => P("Alpha zeta gamma delta."),
        };
        var marker = CommentDoc(first + P(Commented("Returns"), next));
        foreach (var order in Orders())
        {
            var reviewers = InOrder((bool)order[1], editor, marker);
            var settings = new DocxDiffConsolidateSettings { ConflictResolution = (ConflictResolution)order[0] };
            var merged = DocxCompare.Consolidate(original, reviewers, settings);
            var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
            Assert.Equal(View(original), View(rejected));
            Assert.DoesNotContain(Body(merged).Elements(W + "p").Skip(shape == "deleted" ? 0 : 1).Descendants(), e =>
                e.Name == W + "commentRangeStart" && e.Ancestors(W + "p").First().Descendants(W + "t").Any(t => t.Value.Contains("Returns")));
            AssertRangesInOrder(merged, accepted, rejected);
            AssertValid(merged, accepted, rejected);
        }
    }

    [Fact]
    public void Comments_starting_at_one_place_keep_the_reviewers_order()
    {
        // Reviewer B adds a comment that starts where the base's does, around more text, while reviewer A edits the
        // next paragraph; the base's marker orders B's beside it, so B's range opens after the base's, as in B's copy.
        string second = "<w:commentRangeStart w:id=\"1\"/>";
        var original = TwoComments(P("Alpha ", Commented("beta"), " gamma.") + P("Returns are due on Fridays."), true, false);
        var editor = TwoComments(P("Alpha ", Commented("beta"), " gamma.") + P("Returns are due on Mondays."), true, false);
        var marker = TwoComments(P("Alpha ", "<w:commentRangeStart w:id=\"0\"/>", second, Run("beta"), "<w:commentRangeEnd w:id=\"0\"/>",
            "<w:r><w:commentReference w:id=\"0\"/></w:r>", Run(" gamma"), "<w:commentRangeEnd w:id=\"1\"/>",
            "<w:r><w:commentReference w:id=\"1\"/></w:r>", Run(".")) + P("Returns are due on Fridays."), true, true);
        foreach (var order in Orders())
        {
            var reviewers = InOrder((bool)order[1], editor, marker);
            var merged = DocxCompare.Consolidate(original, reviewers, new DocxDiffConsolidateSettings { ConflictResolution = (ConflictResolution)order[0] });
            var accepted = RevisionProcessor.AcceptRevisions(merged);
            var texts = CommentTextsById(accepted);
            var starts = Body(accepted).Descendants(W + "commentRangeStart").Select(e => texts[(string)e.Attribute(W + "id")!]).ToList();
            Assert.Equal(new[] { "Confirm the loan period", "Which loans?" }, starts);
            AssertRangesInOrder(merged, accepted);
            AssertValid(merged, accepted);
        }
    }

    private static System.Collections.Generic.Dictionary<string, string> CommentTextsById(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        return package.MainDocumentPart!.WordprocessingCommentsPart!.GetXDocument().Root!.Elements(W + "comment")
            .ToDictionary(c => (string)c.Attribute(W + "id")!, c => string.Concat(c.Descendants(W + "t").Select(t => t.Value)));
    }

    /// <summary>A document with the base's comment (id 0, "Confirm the loan period") and, or, a second one (id 1).</summary>
    private static WmlDocument TwoComments(string body, bool first, bool second)
    {
        string comments = (first ? "<w:comment w:id=\"0\" w:author=\"Reviewer B\" w:date=\"2026-03-01T00:00:00Z\" w:initials=\"B\"><w:p><w:r><w:t>Confirm the loan period</w:t></w:r></w:p></w:comment>" : "") +
            (second ? "<w:comment w:id=\"1\" w:author=\"Reviewer B\" w:date=\"2026-03-02T00:00:00Z\" w:initials=\"B\"><w:p><w:r><w:t>Which loans?</w:t></w:r></w:p></w:comment>" : "");
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            Write(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles xmlns:w=\"{W}\"/>");
            Write(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings xmlns:w=\"{W}\"/>");
            Write(main.AddNewPart<WordprocessingCommentsPart>(), $"<w:comments xmlns:w=\"{W}\">{comments}</w:comments>");
            Write(main, $"<w:document xmlns:w=\"{W}\"><w:body>{body}</w:body></w:document>");
        }
        return new WmlDocument("comments.docx", stream.ToArray());
    }

    [Fact]
    public void A_comment_the_revision_drops_from_a_run_that_cannot_be_split_comes_back_on_reject()
    {
        // Two-way: the original's comment starts inside a run holding a tab and text, which the copied equal paragraph
        // cannot split; the restored range covers the whole run rather than being dropped (issue #1812 review).
        var original = CommentDoc(P("Alpha ", "<w:commentRangeStart w:id=\"0\"/>", "<w:r><w:tab/><w:t>beta</w:t></w:r>",
            "<w:commentRangeEnd w:id=\"0\"/>", "<w:r><w:commentReference w:id=\"0\"/></w:r>", " gamma.") + P("Returns are due on Fridays."));
        var revised = Doc(P("<w:r><w:t xml:space=\"preserve\">Alpha </w:t><w:tab/><w:t>beta</w:t></w:r>", " gamma.") +
            P("Returns are due on Mondays."));
        var redline = DocxCompare.Compare(original, revised);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Single(Body(rejected).Descendants(W + "commentReference"));
        AssertRangesInOrder(redline, rejected, RevisionProcessor.AcceptRevisions(redline));
        AssertValid(redline, rejected);
    }

    /// <summary>Each comment's range starts before it ends, and its reference follows its end.</summary>
    private static void AssertRangesInOrder(params WmlDocument[] docs)
    {
        foreach (var doc in docs)
        {
            var order = Body(doc).Descendants().Where(e => e.Name.LocalName.StartsWith("comment")).ToList();
            foreach (var start in order.Where(e => e.Name == W + "commentRangeStart"))
            {
                var id = start.Attribute(W + "id")!.Value;
                int end = order.FindIndex(e => e.Name == W + "commentRangeEnd" && e.Attribute(W + "id")!.Value == id);
                int reference = order.FindIndex(e => e.Name == W + "commentReference" && e.Attribute(W + "id")!.Value == id);
                Assert.True(end > order.IndexOf(start), $"comment {id} ends before it starts");
                Assert.True(reference < 0 || reference > end, $"comment {id} is referenced before it ends");
            }
        }
    }

    /// <summary>
    /// A comment the base already has is the base's, not the untouched reviewer's: when reviewer A removes it, as A edits
    /// the paragraph's text or changes a field (which replaces the paragraph whole), accepting gives A's paragraph without
    /// it and rejecting the base with it, whatever reviewer B, who left the paragraph alone, holds.
    /// </summary>
    [Theory]
    [InlineData("text-edited")]
    [InlineData("replaced-whole")]
    public void A_base_comment_another_reviewer_removes_stays_removed(string shape)
    {
        const string field = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> XE \"loan\" </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";
        const string next = "Returns are due on Fridays.";
        var original = CommentDoc(P("The register ", Commented("lists every"), " loan.") + P(next));
        var editor = shape == "text-edited" ? Doc(P("The register lists each loan.") + P(next))
            : Doc(P("The register lists every loan.", field) + P(next));
        var untouched = CommentDoc(P("The register ", Commented("lists every"), " loan.") + P(next) + P("A closing line."));
        var expected = shape == "text-edited" ? Doc(P("The register lists each loan.") + P(next) + P("A closing line."))
            : Doc(P("The register lists every loan.", field) + P(next) + P("A closing line."));
        foreach (var policy in Orders())
            AssertMerged(original, expected, (ConflictResolution)policy[0], InOrder((bool)policy[1], editor, untouched));
    }

    /// <summary>
    /// Reviewer A only changes a paragraph's formatting (centres it, or makes a word bold) while reviewer B adds a
    /// bookmark or a comment in it: B's marker is kept beside A's formatting.
    /// </summary>
    [Theory]
    [InlineData("centred", false)]
    [InlineData("centred", true)]
    [InlineData("bold", false)]
    [InlineData("bold", true)]
    public void A_marker_beside_another_reviewers_formatting_change_is_kept(string format, bool comment)
    {
        const string next = "Returns are due on Fridays.";
        string Formatted(params string[] content) => format == "centred"
            ? "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>" + string.Concat(content.Select(c => c.StartsWith("<w:") ? c : Run(c))) + "</w:p>"
            : P(content);
        string bold = "<w:r><w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\">register</w:t></w:r>";
        string Marked(string text) => comment ? Commented(text) : ConsolidateReviewerMarkerTests.Marked(text);
        WmlDocument Make(string body) => comment ? CommentDoc(body) : Doc(body);
        var original = Doc(P("The register lists every loan.") + P(next));
        var editor = Doc((format == "centred" ? Formatted("The register lists every loan.") : P("The ", bold, " lists every loan.")) + P(next));
        var marker = Make(P("The register ", Marked("lists every"), " loan.") + P(next));
        var expected = Make((format == "centred" ? Formatted("The register ", Marked("lists every"), " loan.")
            : P("The ", bold, " ", Marked("lists every"), " loan.")) + P(next));
        foreach (var order in Orders())
        {
            var accepted = AssertMerged(original, expected, (ConflictResolution)order[0], InOrder((bool)order[1], editor, marker));
            Assert.Equal(format == "centred" ? 1 : 0, Body(accepted).Descendants(W + "jc").Count());
            Assert.Equal(format == "bold" ? 1 : 0, Body(accepted).Descendants(W + "b").Count());
        }
    }

    /// <summary>
    /// A comment reviewer B adds inside a text box, in a paragraph reviewer A edits, is not written out of order: the
    /// merge cannot place it inside the box, so it keeps the base's view of the box (issue #1812 review).
    /// </summary>
    [Fact]
    public void A_comment_inside_a_text_box_another_reviewer_edits_around_stays_in_order()
    {
        static string Box(string inner) => "<w:r><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" style=\"width:120pt;height:40pt\"><v:textbox><w:txbxContent>" +
            inner + "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>";
        const string next = "Returns are due on Fridays.";
        var original = Doc(P("The register ", Box(P("Shelf mark A")), " lists every loan.") + P(next));
        var editor = Doc(P("The register ", Box(P("Shelf mark A")), " lists each loan.") + P(next));
        var marker = CommentDoc(P("The register ", Box(P("Shelf ", Commented("mark"), " A")), " lists every loan.") + P(next));
        foreach (var reviewers in new[] { InOrder(true, editor, marker), InOrder(false, editor, marker) })
        {
            var merged = DocxCompare.Consolidate(original, reviewers);
            var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
            AssertRangesInOrder(merged, accepted, rejected);
            Assert.Empty(Body(rejected).Descendants(W + "commentReference"));
            Assert.Equal(string.Concat(Body(editor).Descendants(W + "t").Select(t => t.Value)),
                string.Concat(Body(accepted).Descendants(W + "t").Select(t => t.Value)));
            AssertValid(merged, accepted, rejected);
        }
    }

    /// <summary>
    /// Reviewer B adds a comment in a footnote no reviewer edits, as reviewer A edits the body: every view stays valid
    /// (the merge brings in comment definitions for the body's references only, so the note's comment is not carried)
    /// and rejecting gives the base.
    /// </summary>
    [Fact]
    public void A_comment_in_a_footnote_beside_another_reviewers_body_edit_leaves_every_view_valid()
    {
        const string separators = "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
            "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";
        const string reference = "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:footnoteReference w:id=\"1\"/></w:r>";
        WmlDocument Make(string body, string note, bool comment) => NoteDoc(body,
            separators + "<w:footnote w:id=\"1\">" + note + "</w:footnote>",
            comment ? "<w:comment w:id=\"0\" w:author=\"Reviewer B\" w:date=\"2026-03-01T00:00:00Z\" w:initials=\"B\"><w:p><w:r><w:t>Cite the edition</w:t></w:r></w:p></w:comment>" : null);
        string body = P("The register lists every loan.", reference) + P("Returns are due on Fridays.");
        string notePlain = P("Source one.") + P("Source two.");
        string noteMarked = P("Source one.") + "<w:bookmarkStart w:id=\"8\" w:name=\"Note\"/><w:bookmarkEnd w:id=\"8\"/>" + P("Source ", Commented("two"), ".");
        var original = Make(body, notePlain, false);
        var editor = Make(body.Replace("every loan", "each loan"), notePlain, false);
        var marker = Make(body, noteMarked, true);
        foreach (var reviewers in new[] { InOrder(true, editor, marker), InOrder(false, editor, marker) })
        {
            var merged = DocxCompare.Consolidate(original, reviewers);
            var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
            AssertValid(merged, accepted, rejected);
            Assert.Equal(View(original), View(rejected));
            Assert.Equal(View(editor), View(accepted));
        }
    }

    private static WmlDocument NoteDoc(string body, string footnotes, string? comments, bool footnote = true)
    {
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            Write(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles xmlns:w=\"{W}\"/>");
            Write(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings xmlns:w=\"{W}\"/>");
            if (footnote)
                Write(main.AddNewPart<FootnotesPart>(), $"<w:footnotes xmlns:w='{W}'>{footnotes}</w:footnotes>");
            else
                Write(main.AddNewPart<EndnotesPart>(), $"<w:endnotes xmlns:w='{W}'>{footnotes}</w:endnotes>");
            if (comments is not null)
                Write(main.AddNewPart<WordprocessingCommentsPart>(), $"<w:comments xmlns:w=\"{W}\">{comments}</w:comments>");
            Write(main, $"<w:document xmlns:w=\"{W}\"><w:body>{body}</w:body></w:document>");
        }
        return new WmlDocument("notes.docx", stream.ToArray());
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml);
    }

    [Fact]
    public void A_comment_with_no_other_reviewer_in_its_paragraph_is_kept()
    {
        // Control: the commenter's paragraph is equal for every reviewer, which main already handled.
        var original = Doc(P("The register lists every loan.") + P("Returns are due on Fridays."));
        var editor = Doc(P("The register lists every loan.") + P("Returns are due on Mondays."));
        var marker = CommentDoc(P(Commented("The register"), " lists every loan.") + P("Returns are due on Fridays."));
        var expected = CommentDoc(P(Commented("The register"), " lists every loan.") + P("Returns are due on Mondays."));
        AssertMerged(original, expected, ConflictResolution.BaseWins, InOrder(true, editor, marker));
    }

    /// <summary>
    /// Reviewer B comments inside a table whose text reviewer A edits. <c>other-row</c>: the comment is in a row A
    /// leaves; <c>same-row</c>: in another cell of the row A edits; <c>same-cell</c>: on the words before A's edit in
    /// the same paragraph; <c>row-added</c>: in a row A leaves while A also adds a row, so the two tables' paragraphs
    /// no longer line up one to one.
    /// </summary>
    [Theory]
    [MemberData(nameof(TableShapes))]
    public void A_comment_inside_a_table_another_reviewer_edits_is_kept(string shape, ConflictResolution policy, bool editorFirst)
    {
        WmlDocument Make(string atlas, string shelf, string globe, string stand, bool comment, bool added = false)
        {
            string body = Table("", Row(atlas, shelf), Row(globe, stand), added ? Row("Chart", "Drawer three") : "");
            return comment ? CommentDoc(body) : Doc(body);
        }
        string c(string text) => P(Commented(text));
        var original = Make("Atlas", "Shelf one", "Globe", "Stand two", false);
        var (editor, marker, expected) = shape switch
        {
            "other-row" => (Make("Atlas", "Shelf one", "Globe", "Stand two, brass", false),
                Make("Atlas", c("Shelf one"), "Globe", "Stand two", true),
                Make("Atlas", c("Shelf one"), "Globe", "Stand two, brass", true)),
            "same-row" => (Make("Atlas", "Shelf one", "Globe", "Stand two, brass", false),
                Make("Atlas", "Shelf one", c("Globe"), "Stand two", true),
                Make("Atlas", "Shelf one", c("Globe"), "Stand two, brass", true)),
            "row-added" => (Make("Atlas", "Shelf one", "Globe", "Stand two, brass", false, added: true),
                Make("Atlas", c("Shelf one"), "Globe", "Stand two", true),
                Make("Atlas", c("Shelf one"), "Globe", "Stand two, brass", true, added: true)),
            _ => (Make("Atlas", "Shelf one", "Globe", "Stand two, brass", false),
                Make("Atlas", "Shelf one", "Globe", P(Commented("Stand"), " two"), true),
                Make("Atlas", "Shelf one", "Globe", P(Commented("Stand"), " two, brass"), true)),
        };
        var accepted = AssertMerged(original, expected, policy, InOrder(editorFirst, editor, marker));
        Assert.Equal(new[] { "Confirm the loan period" }, CommentTexts(accepted));
    }

    public static IEnumerable<object[]> TableShapes() =>
        from shape in new[] { "other-row", "same-row", "same-cell", "row-added" }
        from order in Orders()
        select new object[] { shape, order[0], order[1] };

    /// <summary>
    /// Reviewer S only centres a table and reviewer B only adds a bookmark in one of its cells, in either order, with
    /// or without reviewer A editing another cell (in every position). Accepting keeps the alignment, the bookmark and
    /// the edit.
    /// </summary>
    [Theory]
    [InlineData("S,B")]
    [InlineData("B,S")]
    [InlineData("S,B,A")]
    [InlineData("S,A,B")]
    [InlineData("A,S,B")]
    [InlineData("B,A,S")]
    public void A_bookmark_beside_a_reviewer_who_only_formats_the_table_is_kept(string order)
    {
        const string centred = "<w:jc w:val=\"center\"/>";
        WmlDocument Make(string jc, string shelf, string stand) => Doc(Table(jc, Row("Atlas", shelf), Row("Globe", stand)));
        var original = Make("", "Shelf one", "Stand two");
        var documents = new Dictionary<string, WmlDocument>
        {
            ["S"] = Make(centred, "Shelf one", "Stand two"),
            ["B"] = Make("", P(Marked("Shelf one")), "Stand two"),
            ["A"] = Make("", "Shelf one", "Stand two, brass"),
        };
        var names = order.Split(',');
        var reviewers = names.Select(n => new DocxDiffReviewer { Document = documents[n], Author = "Reviewer " + n }).ToArray();
        var expected = Make(centred, P(Marked("Shelf one")), names.Contains("A") ? "Stand two, brass" : "Stand two");
        foreach (var policy in new[] { ConflictResolution.BaseWins, ConflictResolution.FirstReviewerWins, ConflictResolution.StackAll })
        {
            var accepted = AssertMerged(original, expected, policy, reviewers);
            Assert.Single(Body(accepted).Descendants(W + "tblPr").Descendants(W + "jc"));
        }
    }

    /// <summary>
    /// A table-level content control that holds no row (empty, or only a bookmark) stays between the rows it was
    /// between when two reviewers' edits are composed into the table: first, in the middle, last, and beside a row
    /// one reviewer deletes. A bare bookmark between the rows (the control) already kept its place.
    /// </summary>
    [Theory]
    [InlineData("first", false)]
    [InlineData("middle", false)]
    [InlineData("last", false)]
    [InlineData("middle", true)]
    [InlineData("before-deleted-row", false)]
    [InlineData("bare-bookmark", false)]
    public void A_composed_table_keeps_a_control_that_holds_no_row_where_it_was(string place, bool bookmarked)
    {
        string control = bookmarked
            ? "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id=\"3\" w:name=\"Middle\"/><w:bookmarkEnd w:id=\"3\"/></w:sdtContent></w:sdt>"
            : place == "bare-bookmark" ? "<w:bookmarkStart w:id=\"3\" w:name=\"Middle\"/><w:bookmarkEnd w:id=\"3\"/>"
            : "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent/></w:sdt>";
        WmlDocument Make(string shelf, string stand, bool chart = true)
        {
            var rows = new List<string> { Row("Atlas", shelf), Row("Globe", stand) };
            if (chart) rows.Add(Row("Chart", "Drawer three"));
            int at = place switch { "first" => 0, "last" => rows.Count, "before-deleted-row" => 2, _ => 1 };
            if (place == "before-deleted-row" && !chart) at = rows.Count;
            rows.Insert(at, control);
            return Doc(Table("", rows.ToArray()));
        }
        var original = Make("Shelf one", "Stand two");
        bool deletes = place == "before-deleted-row";
        var editorA = Make("Shelf one, top", "Stand two");
        var editorB = Make("Shelf one", "Stand two, brass", chart: !deletes);
        var expected = Make("Shelf one, top", "Stand two, brass", chart: !deletes);
        foreach (var policy in new[] { ConflictResolution.BaseWins, ConflictResolution.StackAll })
            AssertMerged(original, expected, policy, InOrder(true, editorA, editorB));
    }

    /// <summary>
    /// A row reviewer A inserts next to a content control that holds no row goes on the side of the control A put it,
    /// while reviewer B edits another row: <c>head-after</c>, the control leads the table and A's row follows it;
    /// <c>last-after</c>, the control ends the table and A's row follows it; <c>middle-after</c> and
    /// <c>middle-before</c>, the control sits between two rows and A's row follows or precedes it.
    /// </summary>
    [Theory]
    [InlineData("head-after")]
    [InlineData("last-after")]
    [InlineData("middle-after")]
    [InlineData("middle-before")]
    [InlineData("middle-after-bare-bookmark")]
    public void A_row_inserted_beside_a_control_that_holds_no_row_keeps_its_side(string place)
    {
        const string control = "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id=\"3\" w:name=\"Middle\"/><w:bookmarkEnd w:id=\"3\"/></w:sdtContent></w:sdt>";
        const string bare = "<w:bookmarkStart w:id=\"7\" w:name=\"Gap\"/><w:bookmarkEnd w:id=\"7\"/>";
        string atlas = Row("Atlas", "Shelf one"), added = Row("Chart", "Drawer three");
        WmlDocument Make(string globeText, bool withRow) => Doc(Table("", place switch
        {
            "head-after" => new[] { control, withRow ? added : "", atlas, Row("Globe", globeText) },
            "last-after" => new[] { atlas, Row("Globe", globeText), control, withRow ? added : "" },
            "middle-after-bare-bookmark" => new[] { atlas, bare, control, withRow ? added : "", Row("Globe", globeText) },
            "middle-after" => new[] { atlas, control, withRow ? added : "", Row("Globe", globeText) },
            _ => new[] { atlas, withRow ? added : "", control, Row("Globe", globeText) },
        }));
        var original = Make("Stand two", false);
        var inserted = Make("Stand two", true);
        var pairwise = DocxCompare.Compare(original, inserted);
        var accepted = RevisionProcessor.AcceptRevisions(pairwise);
        var rejected = RevisionProcessor.RejectRevisions(pairwise);
        Assert.Equal(View(inserted), View(accepted));
        Assert.Equal(View(original), View(rejected));
        AssertValid(pairwise, accepted, rejected);

        var expected = Make("Stand two, brass", true);
        foreach (var policy in new[] { ConflictResolution.BaseWins, ConflictResolution.StackAll })
        {
            AssertMerged(original, expected, policy, InOrder(true, Make("Stand two", true), Make("Stand two, brass", false)));
            AssertMerged(original, expected, policy, InOrder(false, Make("Stand two", true), Make("Stand two, brass", false)));
        }
    }

    /// <summary>
    /// A composed table, or a row of one, that holds both a content control holding nothing and a bare bookmark among
    /// its rows or cells keeps each where it was: the bookmark's place is counted in rows (or cells), which the
    /// control is not one of. <c>table</c>: the control is between the first two rows and the bookmark between the
    /// last two; <c>row</c>: the same between a row's three cells.
    /// </summary>
    [Theory]
    [InlineData("table")]
    [InlineData("row")]
    public void A_composed_table_keeps_a_bookmark_beside_a_control_that_holds_nothing_in_place(string level)
    {
        const string empty = "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent/></w:sdt>";
        const string bookmark = "<w:bookmarkStart w:id=\"3\" w:name=\"Gap\"/><w:bookmarkEnd w:id=\"3\"/>";
        WmlDocument Make(string shelf, string stand) => Doc(level == "table"
            ? Table("", Row("Atlas", shelf), empty, Row("Globe", stand), bookmark, Row("Chart", "Drawer three"))
            : Table("", "<w:tr>" + Cell("Atlas") + empty + Cell(shelf) + bookmark + Cell("Drawer three") + "</w:tr>", Row("Globe", stand)));
        var original = Make("Shelf one", "Stand two");
        var expected = Make("Shelf one, top", "Stand two, brass");
        foreach (var policy in new[] { ConflictResolution.BaseWins, ConflictResolution.StackAll })
            AssertMerged(original, expected, policy, InOrder(true, Make("Shelf one, top", "Stand two"), Make("Shelf one", "Stand two, brass")));
    }
}
