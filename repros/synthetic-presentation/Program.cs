using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var folder = Path.GetFullPath(args[0]);
var selected = args.Length > 1 ? args[1] : null;
var output = Path.Combine(folder, "generated");
Directory.CreateDirectory(output);
XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
var reports = new List<object>();
var failed = 0;
string Text(XElement element) => string.Concat(element.Descendants().Where(e => e.Name == w+"t" || e.Name == w+"delText").Select(e => e.Value));
XDocument Xml(ZipArchive archive, string name) { using var stream = archive.GetEntry(name)!.Open(); return XDocument.Load(stream); }

Dictionary<string, string> Properties(XElement? properties)
{
    var result = new Dictionary<string, string>();
    foreach (var child in properties?.Elements() ?? Enumerable.Empty<XElement>())
    {
        if (child.Name.LocalName.EndsWith("Change")) continue;
        if (!child.HasAttributes) result[child.Name.LocalName] = "1";
        foreach (var attribute in child.Attributes()) result[child.Name.LocalName+"/"+attribute.Name.LocalName] = attribute.Value;
    }
    return result;
}
void Merge(Dictionary<string,string> target, Dictionary<string,string> source) { foreach (var item in source) target[item.Key] = item.Value; }

Snapshot Inspect(string path)
{
    using var package = WordprocessingDocument.Open(path, false);
    var errors = new OpenXmlValidator(FileFormatVersions.Office2019) { MaxNumberOfErrors = 0 }.Validate(package)
        .Select(e => new { e.Id, e.Description, Part = e.Part?.Uri.ToString(), Path = e.Path?.XPath }).Cast<object>().ToArray();
    using var zip = ZipFile.OpenRead(path);
    var document = Xml(zip, "word/document.xml");
    var styleRoot = Xml(zip, "word/styles.xml").Root!;
    var defaultId = styleRoot.Elements(w+"style").FirstOrDefault(e => (string?)e.Attribute(w+"type") == "paragraph" &&
        new[] { "1", "true", "on" }.Contains((string?)e.Attribute(w+"default")))?.Attribute(w+"styleId")?.Value;
    var paragraphs = new List<ParagraphSnapshot>();
    foreach (var paragraph in document.Descendants(w+"p"))
    {
        var id = (string?)paragraph.Element(w+"pPr")?.Element(w+"pStyle")?.Attribute(w+"val") ?? defaultId;
        var rp = Properties(styleRoot.Element(w+"docDefaults")?.Element(w+"rPrDefault")?.Element(w+"rPr"));
        var pp = Properties(styleRoot.Element(w+"docDefaults")?.Element(w+"pPrDefault")?.Element(w+"pPr"));
        var chain = new List<XElement>();
        var seen = new HashSet<string>();
        for (var current = id; current is not null && seen.Add(current);)
        {
            var style = styleRoot.Elements(w+"style").FirstOrDefault(e => (string?)e.Attribute(w+"styleId") == current && (string?)e.Attribute(w+"type") == "paragraph");
            if (style is null) break;
            chain.Add(style); current = (string?)style.Element(w+"basedOn")?.Attribute(w+"val");
        }
        chain.Reverse();
        foreach (var style in chain) { Merge(pp, Properties(style.Element(w+"pPr"))); Merge(rp, Properties(style.Element(w+"rPr"))); }
        Merge(pp, Properties(paragraph.Element(w+"pPr")));
        Merge(rp, Properties(paragraph.Descendants(w+"r").FirstOrDefault()?.Element(w+"rPr")));
        pp.Remove("pStyle/val");
        paragraphs.Add(new(Text(paragraph), id, pp, rp));
    }
    var margins = new List<Dictionary<string,string>>();
    foreach (var table in document.Descendants(w+"tbl"))
    {
        var id = (string?)table.Element(w+"tblPr")?.Element(w+"tblStyle")?.Attribute(w+"val");
        var style = styleRoot.Elements(w+"style").FirstOrDefault(e => (string?)e.Attribute(w+"type") == "table" && (string?)e.Attribute(w+"styleId") == id);
        var values = Properties(style?.Element(w+"tblPr")?.Element(w+"tblCellMar"));
        Merge(values, Properties(table.Element(w+"tblPr")?.Element(w+"tblCellMar")));
        margins.Add(values);
    }
    return new(Path.GetFileName(path), errors, paragraphs, margins,
        Properties(document.Root?.Element(w+"body")?.Element(w+"sectPr")),
        document.ToString(SaveOptions.DisableFormatting), styleRoot.ToString(SaveOptions.DisableFormatting));
}

string Canonical(Dictionary<string,string> value) => JsonSerializer.Serialize(value.OrderBy(e => e.Key).ToDictionary(e => e.Key, e => e.Value));
// These fixtures specify literal fonts and spacing. Assert those declared axes only;
// unrelated synthesized metadata is not a failure and this is not a general theme resolver.
string[] runAxes = { "rFonts/ascii", "rFonts/hAnsi", "rFonts/eastAsia", "rFonts/cs", "sz/val", "szCs/val" };
string[] paragraphAxes = { "spacing/before", "spacing/after", "spacing/line", "spacing/lineRule" };
bool SameAxes(Dictionary<string,string> left, Dictionary<string,string> right, string[] axes) =>
    axes.All(key => left.GetValueOrDefault(key) == right.GetValueOrDefault(key));
bool SameFormat(Snapshot left, Snapshot right) => left.Paragraphs.Count == right.Paragraphs.Count && left.Paragraphs.Zip(right.Paragraphs)
    .All(p => p.First.Text == p.Second.Text && SameAxes(p.First.Run, p.Second.Run, runAxes) && SameAxes(p.First.Paragraph, p.Second.Paragraph, paragraphAxes));
string BlockText(IrBlock? block) => block is IrParagraph paragraph ? string.Concat(paragraph.Inlines.OfType<IrTextRun>().Select(e => e.Text)) : "";

using var cases = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "cases.json")));
foreach (var entry in cases.RootElement.EnumerateArray())
{
    var name = entry.GetProperty("name").GetString()!;
    if (selected is not null && selected != name) continue;
    var originalPath = Path.Combine(folder, entry.GetProperty("left").GetString()!);
    var revisedPath = Path.Combine(folder, entry.GetProperty("right").GetString()!);
    var originalDocument = new WmlDocument(originalPath); var revisedDocument = new WmlDocument(revisedPath);
    var compared = DocxCompare.Compare(originalDocument, revisedDocument);
    var comparedPath = Path.Combine(output, name+"-comparison.docx"); compared.SaveAs(comparedPath);
    var acceptedPath = Path.Combine(output, name+"-accepted.docx"); RevisionProcessor.AcceptRevisions(compared).SaveAs(acceptedPath);
    var rejectedPath = Path.Combine(output, name+"-rejected.docx"); RevisionProcessor.RejectRevisions(compared).SaveAs(rejectedPath);
    var original = Inspect(originalPath); var revised = Inspect(revisedPath); var comparison = Inspect(comparedPath);
    var accepted = Inspect(acceptedPath); var rejected = Inspect(rejectedPath);
    var alignment = IrBlockAligner.Align(IrReader.Read(originalDocument), IrReader.Read(revisedDocument), new IrDiffSettings());
    var plan = alignment.Entries.Select(e => new { Kind=e.Kind.ToString(), Left=BlockText(e.Left), Right=BlockText(e.Right),
        Members=e.MultiBlocks?.Select(BlockText).ToArray() }).ToArray();
    var checks = new Dictionary<string,bool> {
        ["all_packages_schema_valid"] = new[] { original, revised, comparison, accepted, rejected }.All(e => e.Errors.Length == 0),
        ["accepted_text_equals_revised"] = accepted.Paragraphs.Select(e=>e.Text).SequenceEqual(revised.Paragraphs.Select(e=>e.Text)),
        ["rejected_text_equals_original"] = rejected.Paragraphs.Select(e=>e.Text).SequenceEqual(original.Paragraphs.Select(e=>e.Text)),
    };
    switch (entry.GetProperty("check").GetString())
    {
        case "effective-format":
            checks["accepted_effective_format_equals_revised"] = SameFormat(accepted, revised);
            checks["rejected_effective_format_equals_original"] = SameFormat(rejected, original);
            break;
        case "table-margins":
            checks["accepted_table_margins_equal_revised"] = accepted.TableMargins.Select(Canonical).SequenceEqual(revised.TableMargins.Select(Canonical));
            checks["rejected_table_margins_equal_original"] = rejected.TableMargins.Select(Canonical).SequenceEqual(original.TableMargins.Select(Canonical));
            break;
        case "section-margin":
            checks["accepted_section_margin_equals_revised"] = accepted.Section.GetValueOrDefault("pgMar/top") == revised.Section.GetValueOrDefault("pgMar/top");
            checks["rejected_section_margin_equals_original"] = rejected.Section.GetValueOrDefault("pgMar/top") == original.Section.GetValueOrDefault("pgMar/top");
            break;
        case "heading-pairing":
            checks["edited_relay_body_is_paired"] = plan.Any(e => e.Left == "Quartz relay signal chamber enabled." && e.Right == "Quartz relay signal chamber disabled.");
            checks["edited_sensor_body_is_paired"] = plan.Any(e => e.Left == "Amber sensor pressure chamber active." && e.Right == "Amber sensor pressure chamber idle.");
            checks["new_heading_is_inserted"] = !name.StartsWith("inserted-heading") || plan.Any(e => e.Kind == "Inserted" && e.Right.StartsWith("Quartz relay signal chamber maintenance"));
            break;
    }
    if (checks.Values.Any(value => !value)) failed++;
    reports.Add(new { Name=name, Group=entry.GetProperty("group").GetString(), Checks=checks, Original=original,
        Revised=revised, Comparison=comparison, Accepted=accepted, Rejected=rejected, Alignment=plan });
    Console.WriteLine(name+": "+string.Join(", ", checks.Select(e => e.Key+"="+e.Value)));
}
File.WriteAllText(Path.Combine(folder, selected is null ? "results.json" : selected+"-results.json"),
    JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented=true }));
return failed == 0 ? 0 : 1;

record ParagraphSnapshot(string Text, string? Style, Dictionary<string,string> Paragraph, Dictionary<string,string> Run);
record Snapshot(string File, object[] Errors, List<ParagraphSnapshot> Paragraphs,
    List<Dictionary<string,string>> TableMargins, Dictionary<string,string> Section, string DocumentXml, string StylesXml);
