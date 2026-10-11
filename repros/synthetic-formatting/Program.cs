using System.Text.Json;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
var output = Path.GetFullPath(args.Length > 0 ? args[0] : "artifacts");
Directory.CreateDirectory(output);

XElement Element(string name, params object[] content) => new(w + name, content);
XAttribute Attr(string name, object value) => new(w + name, value);

WmlDocument Make(string variant, bool revised)
{
    bool numbered = variant == "numbered-body";
    bool numbering = numbered || variant == "unused-numbering";
    bool settings = variant == "changed-settings";
    bool table = variant == "numbered-table";
    numbering |= table;
    var paragraphs = new List<XElement>();
    foreach (var text in revised
        ? new[] { "Moss station records the morning temperature.", "A violet lantern marks the new storage shelf.", "Granite station records the evening temperature." }
        : new[] { "Moss station records the morning temperature.", "Granite station records the evening temperature." })
    {
        var properties = Element("pPr", Element("pStyle", Attr("val", "StationBody")));
        if (numbered || table)
            properties.Add(Element("numPr", Element("ilvl", Attr("val", 0)), Element("numId", Attr("val", 1))));
        paragraphs.Add(Element("p", properties, Element("r", Element("t", text))));
    }
    var body = Element("body");
    if (table)
        body.Add(Element("tbl", Element("tblPr", Element("tblW", Attr("w", 6200), Attr("type", "dxa"))),
            Element("tblGrid", Element("gridCol", Attr("w", 6200))),
            Element("tr", Element("tc", Element("tcPr", Element("tcW", Attr("w", 6200), Attr("type", "dxa"))), paragraphs))),
            Element("p"));
    else
        body.Add(paragraphs);
    body.Add(Element("sectPr", Element("pgSz", Attr("w", 12240), Attr("h", 15840)),
        Element("pgMar", Attr("top", 1440), Attr("right", 1440), Attr("bottom", 1440), Attr("left", 1440))));
    var defaults = Element("docDefaults",
        Element("rPrDefault", Element("rPr",
            Element("rFonts", Attr("ascii", revised ? "DejaVu Serif" : "DejaVu Sans"), Attr("hAnsi", revised ? "DejaVu Serif" : "DejaVu Sans")),
            Element("sz", Attr("val", revised ? 28 : 22)))),
        Element("pPrDefault", Element("pPr", Element("spacing", Attr("after", revised ? 200 : 80), Attr("line", revised ? 310 : 250), Attr("lineRule", "auto")))));
    var styles = Element("styles", defaults,
        Element("style", Attr("type", "paragraph"), Attr("default", 1), Attr("styleId", "StationBody"), Element("name", Attr("val", "Station body"))));
    using var memory = new MemoryStream();
    using (var document = WordprocessingDocument.Create(memory, WordprocessingDocumentType.Document))
    {
        var main = document.AddMainDocumentPart();
        using (var stream = main.GetStream())
            new XDocument(Element("document", new XAttribute(XNamespace.Xmlns + "w", w), body)).Save(stream);
        using (var stream = main.AddNewPart<StyleDefinitionsPart>().GetStream())
            new XDocument(styles).Save(stream);
        if (numbering)
        {
            var numberingXml = Element("numbering",
                Element("abstractNum", Attr("abstractNumId", 0),
                    Element("multiLevelType", Attr("val", "singleLevel")),
                    Element("lvl", Attr("ilvl", 0), Element("start", Attr("val", revised ? 4 : 1)),
                        Element("numFmt", Attr("val", "decimal")), Element("lvlText", Attr("val", "%1.")),
                        Element("lvlJc", Attr("val", "left")),
                        Element("pPr", Element("tabs", Element("tab", Attr("val", "num"), Attr("pos", 720))), Element("ind", Attr("left", 720), Attr("hanging", 360))))),
                Element("num", Attr("numId", 1), Element("abstractNumId", Attr("val", 0))));
            using var stream = main.AddNewPart<NumberingDefinitionsPart>().GetStream();
            new XDocument(numberingXml).Save(stream);
        }
        if (settings)
        {
            using var stream = main.AddNewPart<DocumentSettingsPart>().GetStream();
            new XDocument(Element("settings", Element("defaultTabStop", Attr("val", revised ? 480 : 720)))).Save(stream);
        }
    }
    return new WmlDocument("station.docx", memory.ToArray());
}

XDocument Body(WmlDocument document)
{
    using var memory = new MemoryStream(document.DocumentByteArray);
    using var package = WordprocessingDocument.Open(memory, false);
    using var stream = package.MainDocumentPart!.GetStream();
    return XDocument.Load(stream);
}

object[] Presentation(WmlDocument document)
{
    var assembled = FormattingAssembler.AssembleFormatting(document, new FormattingAssemblerSettings
    {
        CreateHtmlConverterAnnotationAttributes = false,
        ClearStyles = false,
        RemoveStyleNamesFromParagraphAndRunProperties = false,
    });
    return Body(assembled).Descendants(w + "p").Where(p => p.Descendants(w + "t").Any()).Select(p => (object)new
    {
        text = string.Concat(p.Descendants(w + "t").Select(t => t.Value)),
        spacing_after = (string?)p.Element(w + "pPr")?.Element(w + "spacing")?.Attribute(w + "after") ?? "0",
        spacing_line = (string?)p.Element(w + "pPr")?.Element(w + "spacing")?.Attribute(w + "line") ?? "240",
        runs = p.Descendants(w + "r").Where(r => r.Descendants(w + "t").Any()).Select(r => new
        {
            text = string.Concat(r.Descendants(w + "t").Select(t => t.Value)),
            font = (string?)r.Element(w + "rPr")?.Element(w + "rFonts")?.Attribute(w + "ascii"),
            size = (string?)r.Element(w + "rPr")?.Element(w + "sz")?.Attribute(w + "val"),
        }).ToArray(),
    }).ToArray();
}

string Fingerprint(object[] presentation)
{
    using var json = JsonDocument.Parse(JsonSerializer.Serialize(presentation));
    var canonical = json.RootElement.EnumerateArray().Select(paragraph => new
    {
        text = paragraph.GetProperty("text").GetString(),
        after = paragraph.GetProperty("spacing_after").GetString(),
        line = paragraph.GetProperty("spacing_line").GetString(),
        characters = paragraph.GetProperty("runs").EnumerateArray().SelectMany(run =>
            run.GetProperty("text").GetString()!.Select(character => new
            {
                character,
                font = run.GetProperty("font").GetString(),
                size = run.GetProperty("size").GetString(),
            })).ToArray(),
    }).ToArray();
    return JsonSerializer.Serialize(canonical);
}

var reports = new List<object>();
foreach (var name in new[] { "plain-control", "unused-numbering", "numbered-body", "changed-settings", "numbered-table" })
{
    var left = Make(name, false);
    var right = Make(name, true);
    var compared = DocxCompare.Compare(left, right);
    var accepted = RevisionProcessor.AcceptRevisions(compared);
    var rejected = RevisionProcessor.RejectRevisions(compared);
    var validation = new List<object>();
    foreach (var (document, view) in new[] { (left, "original"), (right, "revised"), (compared, "compared"), (accepted, "accepted"), (rejected, "rejected") })
    {
        document.SaveAs(Path.Combine(output, name + "-" + view + ".docx"));
        using var memory = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(memory, false);
        validation.Add(new { view, errors = new OpenXmlValidator(FileFormatVersions.Office2019) { MaxNumberOfErrors = 0 }.Validate(package).Select(e => new { e.Id, e.Description }).ToArray() });
    }
    string Text(WmlDocument d) => string.Concat(Body(d).Descendants(w + "t").Select(t => t.Value));
    var actualAccepted = Presentation(accepted);
    var expectedAccepted = Presentation(right);
    var actualRejected = Presentation(rejected);
    var expectedRejected = Presentation(left);
    reports.Add(new
    {
        name,
        accepted_text_matches = Text(accepted) == Text(right),
        rejected_text_matches = Text(rejected) == Text(left),
        accepted_presentation_matches = Fingerprint(actualAccepted) == Fingerprint(expectedAccepted),
        rejected_presentation_matches = Fingerprint(actualRejected) == Fingerprint(expectedRejected),
        expected_accepted = expectedAccepted,
        actual_accepted = actualAccepted,
        expected_rejected = expectedRejected,
        actual_rejected = actualRejected,
        validation,
    });
}
var json = JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true });
File.WriteAllText(Path.Combine(output, "results.json"), json);
Console.WriteLine(json);
