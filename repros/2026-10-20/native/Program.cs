using Docxodus;
using System.Reflection;
using System.Text.Json;
using System.Xml.Linq;

Console.Error.WriteLine(typeof(WmlDocument).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion);
foreach (var line in File.ReadLines(args[0]))
{
    var task = JsonSerializer.Deserialize<Dictionary<string, string>>(line)!;
    try
    {
        var source = new WmlDocument(task["input"]);
        switch (task["operation"])
        {
            case "anchors":
                using (var session = new DocxSession(File.ReadAllBytes(task["input"])))
                    File.WriteAllBytes(task["output"], session.Save(persistAnchorIds: true));
                break;
            case "validate":
                using (var document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(task["input"], false))
                {
                    var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2021)
                        .Validate(document).Select(e => new { e.Description, path = e.Path?.XPath }).ToArray();
                    File.WriteAllText(task["output"], JsonSerializer.Serialize(errors, new JsonSerializerOptions { WriteIndented = true }));
                }
                break;
            case "accept":
                RevisionProcessor.AcceptRevisions(source).SaveAs(task["output"]);
                break;
            case "reject":
                RevisionProcessor.RejectRevisions(source).SaveAs(task["output"]);
                break;
            case "compare":
                DocxCompare.Compare(source, new WmlDocument(task["target"])).SaveAs(task["output"]);
                break;
            case "html":
                var settings = new WmlToHtmlConverterSettings {
                    RenderTrackedChanges = true,
                    RenderFootnotesAndEndnotes = true,
                    RenderHeadersAndFooters = true,
                    RenderPagination = PaginationMode.Paginated,
                    ImageHandler = image => new XElement(XNamespace.Get("http://www.w3.org/1999/xhtml") + "img",
                        new XAttribute("data-image", "present"))
                };
                File.WriteAllText(task["output"], WmlToHtmlConverter.ConvertToHtml(source, settings).ToString());
                break;
            default: throw new ArgumentException("Unknown operation");
        }
        Console.WriteLine(JsonSerializer.Serialize(new { task, ok = true }));
    }
    catch (Exception error)
    {
        Console.WriteLine(JsonSerializer.Serialize(new { task, ok = false, error = error.ToString() }));
    }
}
