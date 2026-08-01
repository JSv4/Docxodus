#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace OxPt;

/// <summary>
/// Drawing-object id uniqueness in <see cref="DocxDiff.Compare"/> output. A changed shape or textbox is
/// rendered as a deleted copy plus an inserted copy; both carry the source's drawing ids, and two revisions
/// of one document share them — so without re-issuing, the output is schema-invalid.
/// </summary>
public class DocxDiffDrawingIdCollisionTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";

    private const string Namespaces =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
        "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
        "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" " +
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"";

    private static WmlDocument Doc(string bodyInner)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write($"<w:document {Namespaces}><w:body>{bodyInner}" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("drawing.docx", ms.ToArray());
    }

    /// <summary>A VML textbox — shapetype + shape, the shape referencing the shapetype by <c>type="#id"</c>.</summary>
    private static string VmlTextbox(string boxText) =>
        "<w:p><w:r><w:t xml:space=\"preserve\">Body.</w:t></w:r>" +
        "<w:r><w:pict>" +
        "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" " +
        "path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/></v:shapetype>" +
        "<v:shape id=\"s1\" type=\"#_x0000_t202\" style=\"width:200pt;height:50pt\" o:spid=\"_x0000_s1026\">" +
        $"<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{boxText}</w:t></w:r></w:p>" +
        "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>";

    /// <summary>A DrawingML textbox — the shape Word actually writes, carrying <c>wp:docPr/@id</c>.</summary>
    private static string DrawingMlTextbox(string boxText) =>
        "<w:p><w:r><w:t xml:space=\"preserve\">Body.</w:t></w:r>" +
        "<w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" +
        "<wp:extent cx=\"2540000\" cy=\"635000\"/><wp:docPr id=\"7\" name=\"Text Box 7\"/>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
        "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr>" +
        "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"2540000\" cy=\"635000\"/></a:xfrm>" +
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></wps:spPr>" +
        $"<wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{boxText}</w:t></w:r></w:p></w:txbxContent></wps:txbx>" +
        "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    /// <summary>
    /// Forces a changed textbox to be deleted and re-inserted WHOLESALE. Since fine textbox-interior
    /// rendering landed, an
    /// ordinary textbox edit is tracked inside ONE box and no longer duplicates the drawing — so the
    /// id-collision cases now come from the paths that still emit a drawing twice: this option off, a surplus
    /// box, a whole-block paragraph replacement, Consolidate, or a source that already shares an id.
    /// </summary>
    private static DocxDiffSettings Wholesale => new() { CompareTextboxes = false };

    private static List<string> ValidationErrors(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return new OpenXmlValidator().Validate(wDoc).Select(e => e.Description).Distinct().ToList();
    }

    private static XElement Body(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        using var partStream = wDoc.MainDocumentPart!.GetStream();
        return XDocument.Load(partStream).Root!.Element(W + "body")!;
    }

    private static List<string> AttributeValues(XElement body, XName element, string attribute) =>
        body.Descendants(element).Select(e => (string?)e.Attribute(attribute) ?? "").ToList();

    /// <summary>Every VML element sharing the shape-id space.</summary>
    private static List<string> VmlShapeIds(XElement body) => body
        .Descendants()
        .Where(d => d.Name.Namespace == V && d.Name.LocalName is
            "shape" or "group" or "rect" or "oval" or "roundrect" or
            "line" or "polyline" or "arc" or "curve" or "image")
        .Select(d => (string?)d.Attribute("id") ?? "")
        .ToList();

    private static void AssertNoDanglingShapeTypeReference(XElement body, string stage)
    {
        var declared = AttributeValues(body, V + "shapetype", "id").ToHashSet();
        var referenced = body.Descendants(V + "shape")
            .Select(s => (string?)s.Attribute("type"))
            .Where(t => t is not null && t!.StartsWith("#"))
            .Select(t => t!.Substring(1))
            .ToList();

        var dangling = referenced.Where(r => !declared.Contains(r)).Distinct().ToList();
        Assert.True(dangling.Count == 0,
            $"{stage}: shapetype references with no declaration: {string.Join(", ", dangling)} " +
            $"(declared: {string.Join(", ", declared)})");
    }

    private static WmlDocument Accept(WmlDocument doc) =>
        new("a.docx", Docxodus.Internal.DocxDiffOps.AcceptRevisions(doc.DocumentByteArray));

    private static WmlDocument Reject(WmlDocument doc) =>
        new("r.docx", Docxodus.Internal.DocxDiffOps.RejectRevisions(doc.DocumentByteArray));

    // ---------------------------------------------------------------- DrawingML (the Word-authored case)

    [Fact]
    public void A_changed_DrawingML_textbox_sharing_a_docPr_id_compares_to_a_valid_document()
    {
        var compared = DocxDiff.Compare(
            Doc(DrawingMlTextbox("box one")), Doc(DrawingMlTextbox("box two")), Wholesale);
        Assert.Empty(ValidationErrors(compared));
    }

    [Fact]
    public void The_duplicated_docPr_ids_are_re_issued_uniquely()
    {
        var compared = DocxDiff.Compare(Doc(DrawingMlTextbox("box one")), Doc(DrawingMlTextbox("box two")), Wholesale);

        var ids = AttributeValues(Body(compared), WP + "docPr", "id");
        Assert.Equal(2, ids.Count);                       // both copies survive (del + ins)
        Assert.Equal(ids.Count, ids.Distinct().Count());  // …with distinct ids
        Assert.Contains("7", ids);                        // the first occurrence keeps the source value
    }

    // ---------------------------------------------------------------- VML

    [Fact]
    public void A_changed_VML_textbox_sharing_shape_ids_compares_to_a_valid_document()
    {
        var compared = DocxDiff.Compare(Doc(VmlTextbox("box one")), Doc(VmlTextbox("box two")), Wholesale);
        Assert.Empty(ValidationErrors(compared));
    }

    [Fact]
    public void The_duplicated_VML_shape_and_shapetype_ids_are_re_issued_uniquely()
    {
        var compared = DocxDiff.Compare(Doc(VmlTextbox("box one")), Doc(VmlTextbox("box two")), Wholesale);
        var body = Body(compared);

        var shapeIds = AttributeValues(body, V + "shape", "id");
        Assert.Equal(2, shapeIds.Count);
        Assert.Equal(shapeIds.Count, shapeIds.Distinct().Count());
        Assert.Contains("s1", shapeIds);

        var shapeTypeIds = AttributeValues(body, V + "shapetype", "id");
        Assert.Equal(2, shapeTypeIds.Count);
        Assert.Equal(shapeTypeIds.Count, shapeTypeIds.Distinct().Count());
    }

    [Fact]
    public void A_re_issued_shapetype_id_keeps_its_shape_reference_resolvable()
    {
        var compared = DocxDiff.Compare(Doc(VmlTextbox("box one")), Doc(VmlTextbox("box two")));

        // The compared document is not enough: a reference bound to the wrong copy only dangles once one
        // side's copy is dropped, and the validator never catches it.
        AssertNoDanglingShapeTypeReference(Body(compared), "compared");
        AssertNoDanglingShapeTypeReference(Body(Accept(compared)), "accepted");
        AssertNoDanglingShapeTypeReference(Body(Reject(compared)), "rejected");
    }

    [Fact]
    public void A_shapetype_declared_in_one_pict_and_referenced_from_another_stays_resolvable()
    {
        // Word's actual layout: ONE shapetype declaration, referenced by shapes in later picts.
        static string TwoPicts(string firstBox, string secondBox) =>
            "<w:p><w:r><w:pict>" +
            "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" " +
            "path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/></v:shapetype>" +
            "<v:shape id=\"sA\" type=\"#_x0000_t202\" style=\"width:200pt;height:50pt\">" +
            $"<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{firstBox}</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>" +
            "<w:r><w:pict>" +
            "<v:shape id=\"sB\" type=\"#_x0000_t202\" style=\"width:200pt;height:50pt\">" +
            $"<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{secondBox}</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>";

        var compared = DocxDiff.Compare(
            Doc(TwoPicts("first one", "second one")), Doc(TwoPicts("first two", "second two")), Wholesale);

        Assert.Empty(ValidationErrors(compared));
        AssertNoDanglingShapeTypeReference(Body(compared), "compared");
        AssertNoDanglingShapeTypeReference(Body(Accept(compared)), "accepted");
        AssertNoDanglingShapeTypeReference(Body(Reject(compared)), "rejected");
    }

    [Fact]
    public void A_bare_reference_still_resolves_in_the_accepted_document()
    {
        // The declaring pict changes (so its declaration duplicates into del+ins copies) while a SECOND pict
        // in an untouched paragraph references the shapetype BARE — belonging to no revision side, so it
        // cannot be rebound. The original id must therefore land on the copy the accepted document keeps.
        static string DeclaringPara(string boxText) =>
            "<w:p><w:r><w:pict>" +
            "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" " +
            "path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/></v:shapetype>" +
            "<v:shape id=\"decl\" type=\"#_x0000_t202\" style=\"width:200pt;height:50pt\">" +
            $"<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{boxText}</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>";
        const string untouchedReferencingPara =
            "<w:p><w:r><w:t xml:space=\"preserve\">Untouched paragraph text.</w:t></w:r>" +
            "<w:r><w:pict>" +
            "<v:shape id=\"ref\" type=\"#_x0000_t202\" style=\"width:200pt;height:50pt\">" +
            "<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">unchanged box</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>";

        var compared = DocxDiff.Compare(
            Doc(DeclaringPara("declaring box one") + untouchedReferencingPara),
            Doc(DeclaringPara("declaring box two") + untouchedReferencingPara), Wholesale);

        Assert.Empty(ValidationErrors(compared));
        AssertNoDanglingShapeTypeReference(Body(compared), "compared");
        AssertNoDanglingShapeTypeReference(Body(Accept(compared)), "accepted");
    }

    [Fact]
    public void A_minted_id_never_collides_with_a_declaration_that_already_holds_it()
    {
        // Three declarations where the source ALREADY contains the name a mint would produce, and the keeper
        // sits last: minting against only the ids visited so far would hand out "_ptdup1" twice.
        static string PictWithShapeType(string shapeTypeId, string shapeId, string boxText) =>
            "<w:p><w:r><w:pict>" +
            $"<v:shapetype id=\"{shapeTypeId}\" coordsize=\"21600,21600\" o:spt=\"202\" " +
            "path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/></v:shapetype>" +
            $"<v:shape id=\"{shapeId}\" type=\"#{shapeTypeId}\" style=\"width:200pt;height:50pt\">" +
            $"<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{boxText}</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>";

        // First pict changes (its declaration duplicates); a second pict already declares "X_ptdup1"; a third
        // declares "X" bare and unchanged.
        static string Body3(string firstBox) =>
            PictWithShapeType("X", "s1", firstBox) +
            PictWithShapeType("X_ptdup1", "s2", "second box") +
            PictWithShapeType("X", "s3", "third box");

        var compared = DocxDiff.Compare(Doc(Body3("first one")), Doc(Body3("first two")));
        var body = Body(compared);

        var declared = AttributeValues(body, V + "shapetype", "id");
        Assert.Equal(declared.Count, declared.Distinct().Count());
        Assert.Empty(ValidationErrors(compared));
        AssertNoDanglingShapeTypeReference(body, "compared");
        AssertNoDanglingShapeTypeReference(Body(Accept(compared)), "accepted");
    }

    [Fact]
    public void A_changed_VML_rect_sharing_an_id_compares_to_a_valid_document()
    {
        static string Rect(string text) =>
            "<w:p><w:r><w:pict>" +
            "<v:rect id=\"r1\" style=\"width:200pt;height:50pt\">" +
            $"<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:rect></w:pict></w:r></w:p>";

        var compared = DocxDiff.Compare(Doc(Rect("rect one")), Doc(Rect("rect two")), Wholesale);

        Assert.Empty(ValidationErrors(compared));
        var ids = VmlShapeIds(Body(compared));
        Assert.Equal(2, ids.Count);
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Fact]
    public void A_changed_VML_group_sharing_ids_compares_to_a_valid_document()
    {
        // v:group and the v:shape inside it share ONE id space, so both must de-duplicate together.
        static string Group(string text) =>
            "<w:p><w:r><w:pict>" +
            "<v:group id=\"g1\" style=\"width:200pt;height:50pt\" coordsize=\"21600,21600\">" +
            "<v:shape id=\"gs1\" style=\"position:absolute;width:21600;height:21600\">" +
            $"<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></v:group></w:pict></w:r></w:p>";

        var compared = DocxDiff.Compare(Doc(Group("group one")), Doc(Group("group two")), Wholesale);

        Assert.Empty(ValidationErrors(compared));
        var ids = VmlShapeIds(Body(compared));
        Assert.Equal(4, ids.Count);                       // 2 groups + 2 inner shapes
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Fact]
    public void A_re_issued_shape_under_an_object_keeps_its_OLE_binding()
    {
        // A w:object is an opaque inline the renderer keeps ONCE, so the diff never duplicates it — the
        // OLE binding path is robustness against MALFORMED input that already shares a shape id (this
        // fixture's own source is schema-invalid for that reason), not a shape Word authors.
        static string OleObject(string objectId) =>
            "<w:p><w:r><w:object w:dxaOrig=\"1440\" w:dyaOrig=\"1440\">" +
            "<v:shape id=\"os1\" style=\"width:100pt;height:50pt\" o:ole=\"\"/>" +
            $"<o:OLEObject Type=\"Embed\" ProgID=\"Package\" ShapeID=\"os1\" DrawAspect=\"Content\" ObjectID=\"{objectId}\"/>" +
            "</w:object></w:r></w:p>";

        var body = Body(DocxDiff.Compare(
            Doc(OleObject("_1") + OleObject("_2") + "<w:p><w:r><w:t>before</w:t></w:r></w:p>"),
            Doc(OleObject("_1") + OleObject("_2") + "<w:p><w:r><w:t>after</w:t></w:r></w:p>")));

        var shapeIds = VmlShapeIds(body);
        var bindings = body.Descendants(O + "OLEObject")
            .Select(o => (string?)o.Attribute("ShapeID") ?? "")
            .ToList();

        Assert.Equal(2, shapeIds.Count);                          // both objects survived the compare
        Assert.Equal(shapeIds.Count, shapeIds.Distinct().Count());
        Assert.Equal(2, bindings.Count);
        Assert.All(bindings, b => Assert.Contains(b, shapeIds));  // each binding resolves…
        Assert.Equal(bindings.Count, bindings.Distinct().Count()); // …to its OWN shape
    }

    [Fact]
    public void A_moved_paragraph_keeps_its_shapetype_reference_resolvable()
    {
        // Move markup is a third revision-side pair (w:moveFrom/w:moveTo): treating a relocated paragraph's
        // two copies as bare content rebinds both to one declaration and dangles the survivor on reject.
        const string boxPara =
            "<w:p><w:r><w:pict>" +
            "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" " +
            "path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/></v:shapetype>" +
            "<v:shape id=\"ms1\" type=\"#_x0000_t202\" style=\"width:200pt;height:50pt\">" +
            "<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">boxed content</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>";
        static string Para(string text) =>
            $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

        // The boxed paragraph relocates from the top to the bottom.
        var left = Doc(boxPara + Para("alpha one two three") + Para("beta four five six"));
        var right = Doc(Para("alpha one two three") + Para("beta four five six") + boxPara);

        var compared = DocxDiff.Compare(left, right);

        Assert.Empty(ValidationErrors(compared));
        AssertNoDanglingShapeTypeReference(Body(compared), "compared");
        AssertNoDanglingShapeTypeReference(Body(Accept(compared)), "accepted");
        AssertNoDanglingShapeTypeReference(Body(Reject(compared)), "rejected");
    }

    [Fact]
    public void Consolidate_also_uniquifies_drawing_ids()
    {
        var baseDoc = Doc(DrawingMlTextbox("box base"));
        var reviewers = new List<DocxDiffReviewer>
        {
            new() { Author = "reviewer", Document = Doc(DrawingMlTextbox("box revised")) },
        };

        var consolidated = DocxDiff.Consolidate(baseDoc, reviewers, new DocxDiffConsolidateSettings { Diff = Wholesale });

        Assert.Empty(ValidationErrors(consolidated));
        var ids = AttributeValues(Body(consolidated), WP + "docPr", "id");
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Fact]
    public void Consolidate_uniquifies_drawing_ids_in_a_footnote_too()
    {
        // The composite renderer must normalize every story, not just the body — a reviewer-changed drawing
        // inside a footnote duplicates its ids exactly as a body one does.
        static WmlDocument NoteWithDrawing(string boxText)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
                main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

                var notes = main.AddNewPart<FootnotesPart>();
                using (var noteStream = notes.GetStream(FileMode.Create, FileAccess.Write))
                using (var noteWriter = new StreamWriter(noteStream, new UTF8Encoding(false)))
                {
                    noteWriter.Write($"<w:footnotes {Namespaces}>" +
                        "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                        "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\">" +
                        "<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                        "<w:footnote w:id=\"1\"><w:p><w:r><w:pict>" +
                        "<v:rect id=\"nr1\" style=\"width:100pt;height:30pt\">" +
                        $"<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{boxText}</w:t></w:r></w:p>" +
                        "</w:txbxContent></v:textbox></v:rect></w:pict></w:r></w:p></w:footnote>" +
                        "</w:footnotes>");
                }

                using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
                using var writer = new StreamWriter(stream, new UTF8Encoding(false));
                writer.Write($"<w:document {Namespaces}><w:body>" +
                    "<w:p><w:r><w:t xml:space=\"preserve\">Body.</w:t></w:r>" +
                    "<w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
                    "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
            }
            return new WmlDocument("note-drawing.docx", ms.ToArray());
        }

        var reviewers = new List<DocxDiffReviewer>
        {
            new() { Author = "reviewer", Document = NoteWithDrawing("note box revised") },
        };

        var consolidated = DocxDiff.Consolidate(NoteWithDrawing("note box base"), reviewers, new DocxDiffConsolidateSettings { Diff = Wholesale });
        Assert.Empty(ValidationErrors(consolidated));
    }

    // ---------------------------------------------------------------- no-op guarantees

    [Fact]
    public void A_drawing_that_is_emitted_once_is_not_re_issued()
    {
        // Narrow by design: the guarantee is "no duplicate emitted ⇒ no re-issue". A source that repeats a
        // declaration in every pict is duplicated on its own terms and IS re-issued, correctly.
        foreach (var build in new[] { VmlTextbox("same box"), DrawingMlTextbox("same box") })
        {
            var compared = DocxDiff.Compare(Doc(build), Doc(build));
            Assert.Empty(ValidationErrors(compared));

            var body = Body(compared);
            Assert.DoesNotContain("_ptdup", body.ToString());
            var docPrIds = AttributeValues(body, WP + "docPr", "id");
            if (docPrIds.Count > 0)
                Assert.Equal(new[] { "7" }, docPrIds);
        }
    }

    [Fact]
    public void Two_distinct_drawings_with_already_unique_ids_are_left_alone()
    {
        var left = Doc(DrawingMlTextbox("box one") + DrawingMlTextbox("box two").Replace("id=\"7\"", "id=\"8\""));
        var right = Doc(DrawingMlTextbox("box one") + DrawingMlTextbox("box three").Replace("id=\"7\"", "id=\"8\""));

        var compared = DocxDiff.Compare(left, right, Wholesale);
        Assert.Empty(ValidationErrors(compared));

        var ids = AttributeValues(Body(compared), WP + "docPr", "id");
        Assert.Equal(ids.Count, ids.Distinct().Count());
        Assert.Contains("7", ids);
        Assert.Contains("8", ids);
    }

    [Fact]
    public void Round_trip_survives_the_re_issued_ids()
    {
        // Wholesale, like its siblings: ids are only re-issued when a drawing is emitted twice. It also keeps
        // the box text in one run — the fine path splits it at the change boundary ("box " + "two").
        var compared = DocxDiff.Compare(
            Doc(DrawingMlTextbox("box one")), Doc(DrawingMlTextbox("box two")), Wholesale);

        var accepted = Body(new WmlDocument("a.docx",
            Docxodus.Internal.DocxDiffOps.AcceptRevisions(compared.DocumentByteArray))).ToString();
        var rejected = Body(new WmlDocument("r.docx",
            Docxodus.Internal.DocxDiffOps.RejectRevisions(compared.DocumentByteArray))).ToString();

        Assert.Contains("box two", accepted);
        Assert.DoesNotContain("box one", accepted);
        Assert.Contains("box one", rejected);
        Assert.DoesNotContain("box two", rejected);
    }
}
