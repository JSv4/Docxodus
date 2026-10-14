// Measures what one DocxSession edit costs as a document grows, and how much undo history the
// default memory budget keeps (issues #965, #1022). Builds a synthetic document of N 40-word
// paragraphs plus K embedded images of M MiB each, then applies E ReplaceText edits to distinct
// paragraphs, with the per-op markdown patch on (the default) or off.
//
//   dotnet run -c Release -- [paragraphs=4000] [images=4] [imageMiB=2] [edits=20] [patch=on|off]

using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;

var paragraphs = args.Length > 0 ? int.Parse(args[0]) : 4000;
var images = args.Length > 1 ? int.Parse(args[1]) : 4;
var imageMiB = args.Length > 2 ? double.Parse(args[2], System.Globalization.CultureInfo.InvariantCulture) : 2;
var edits = args.Length > 3 ? int.Parse(args[3]) : 20;
var patch = args.Length <= 4 || args[4] != "off";

var docx = BuildDocument(paragraphs, images, (int)(imageMiB * 1024 * 1024));
Console.WriteLine($"document: {paragraphs} paragraphs, {images} images x {imageMiB} MiB, {docx.Length / 1024} KiB on disk, markdown patch {(patch ? "on" : "off")}");

using var session = new DocxSession(docx, new DocxSessionSettings { EmitMarkdownPatch = patch });
var anchors = session.Project().AnchorIndex.Values
    .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p")
    .Select(t => t.Anchor.Id)
    .ToList();
var step = Math.Max(1, anchors.Count / edits);

// One untimed edit and undo to warm the JIT and the projection cache.
Require(session.ReplaceText(anchors[0], "Warm-up edit."));
session.Undo();

var times = new List<double>();
var allocated = new List<long>();
long undoAfterFirst = 0;
for (var i = 0; i < edits; i++)
{
    var anchor = anchors[(i * step) % anchors.Count];
    var before = GC.GetAllocatedBytesForCurrentThread();
    var watch = Stopwatch.StartNew();
    Require(session.ReplaceText(anchor, $"Edited paragraph number {i} with replacement text."));
    watch.Stop();
    times.Add(watch.Elapsed.TotalMilliseconds);
    allocated.Add(GC.GetAllocatedBytesForCurrentThread() - before);
    if (i == 0) undoAfterFirst = session.UndoMemoryBytes;
}

times.Sort();
Console.WriteLine($"per edit: median {times[times.Count / 2]:F1} ms, max {times[^1]:F1} ms");
Console.WriteLine($"allocated per edit: median {allocated.OrderBy(a => a).ElementAt(allocated.Count / 2) / 1024.0:F1} KiB");
if (session.UndoCount == edits)
    Console.WriteLine($"undo bytes added per step: {(session.UndoMemoryBytes - undoAfterFirst) / (double)(edits - 1) / 1024.0:F1} KiB");
Console.WriteLine($"undo: {session.UndoCount} of {edits} steps kept, {session.UndoMemoryBytes / (1024.0 * 1024):F1} MiB counted against the budget, trimmed for memory: {session.UndoHistoryTrimmedForMemory}");
GC.Collect();
GC.WaitForPendingFinalizers();
GC.Collect();
Console.WriteLine($"managed heap after the edits: {GC.GetTotalMemory(forceFullCollection: true) / (1024.0 * 1024):F1} MiB");

static void Require(EditResult result)
{
    if (!result.Success) throw new InvalidOperationException(result.Error?.Message);
}

static byte[] BuildDocument(int paragraphs, int images, int imageBytes)
{
    const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var words = "lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor".Split(' ');
    using var stream = new MemoryStream();
    using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
    {
        var main = doc.AddMainDocumentPart();
        main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
        main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
        var random = new Random(965);
        var body = new StringBuilder();
        for (var i = 0; i < images; i++)
        {
            var image = main.AddImagePart(ImagePartType.Png);
            var bytes = new byte[imageBytes];
            random.NextBytes(bytes);
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }.CopyTo(bytes, 0);
            using (var input = new MemoryStream(bytes)) image.FeedData(input);
            body.Append(
                "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"914400\"/>" +
                $"<wp:docPr id=\"{i + 1}\" name=\"Picture {i + 1}\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                $"<pic:pic><pic:nvPicPr><pic:cNvPr id=\"{i + 1}\" name=\"image{i + 1}.png\"/><pic:cNvPicPr/></pic:nvPicPr>" +
                $"<pic:blipFill><a:blip r:embed=\"{main.GetIdOfPart(image)}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
                "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\"rect\"/></pic:spPr></pic:pic>" +
                "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>");
        }
        for (var p = 0; p < paragraphs; p++)
        {
            body.Append("<w:p><w:r><w:t xml:space=\"preserve\">");
            for (var w = 0; w < 40; w++) body.Append(words[(p + w) % words.Length]).Append(' ');
            body.Append("</w:t></w:r></w:p>");
        }
        using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            writer.Write(
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                $"<w:body>{body}<w:sectPr/></w:body></w:document>");
        doc.Save();
    }
    return stream.ToArray();
}
