using System.Xml.Linq;
using System.Text.Json;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;

XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
var output = Path.GetFullPath(args.Length == 0 ? "artifacts" : args[0]);
Directory.CreateDirectory(output);
string P(string text) => $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>";
string C(string text, int width, int span = 1, string? merge = null) =>
    $"<w:tc><w:tcPr><w:tcW w:w='{width}' w:type='dxa'/>" +
    (span == 1 ? "" : $"<w:gridSpan w:val='{span}'/>") +
    (merge is null ? "" : $"<w:vMerge w:val='{merge}'/>") + "</w:tcPr>" + P(text) + "</w:tc>";
string R(string cells, int before = 0, int after = 0, int beforeWidth = 0, int afterWidth = 0) =>
    "<w:tr>" + (before + after == 0 ? "" : "<w:trPr>" +
    (before == 0 ? "" : $"<w:gridBefore w:val='{before}'/>") +
    (after == 0 ? "" : $"<w:gridAfter w:val='{after}'/>") +
    (before == 0 ? "" : $"<w:wBefore w:w='{beforeWidth}' w:type='dxa'/>") +
    (after == 0 ? "" : $"<w:wAfter w:w='{afterWidth}' w:type='dxa'/>") + "</w:trPr>") + cells + "</w:tr>";
string T(int[] grid, params string[] rows) => "<w:tbl><w:tblPr><w:tblW w:w='7200' w:type='dxa'/><w:tblLayout w:type='fixed'/></w:tblPr><w:tblGrid>" +
    string.Concat(grid.Select(v => $"<w:gridCol w:w='{v}'/>")) + "</w:tblGrid>" + string.Concat(rows) + "</w:tbl>";
WmlDocument Make(string table) {
    using var ms = new MemoryStream();
    using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document)) {
        var main = doc.AddMainDocumentPart();
        using (var stream = main.GetStream()) XElement.Parse($"<w:document xmlns:w='{w}'><w:body>" + P("Canal survey register") + table + P("End of register.") +
            "<w:sectPr><w:pgSz w:w='12240' w:h='15840'/><w:pgMar w:top='1440' w:right='1440' w:bottom='1440' w:left='1440'/></w:sectPr></w:body></w:document>").Save(stream);
        var styles = main.AddNewPart<StyleDefinitionsPart>();
        using (var stream = styles.GetStream()) XElement.Parse($"<w:styles xmlns:w='{w}'><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii='Liberation Sans' w:hAnsi='Liberation Sans'/><w:sz w:val='22'/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after='0' w:line='240' w:lineRule='auto'/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type='paragraph' w:default='1' w:styleId='Body'><w:name w:val='Body'/></w:style></w:styles>").Save(stream);
    }
    return new WmlDocument("canal-register.docx", ms.ToArray());
}
XDocument Read(WmlDocument document) {
    using var ms = new MemoryStream(document.DocumentByteArray);
    using var pkg = WordprocessingDocument.Open(ms, false);
    using var stream = pkg.MainDocumentPart!.GetStream();
    return XDocument.Load(stream);
}
string? Val(XElement? node, string name) => (string?)node?.Attribute(w + name);
object Geometry(WmlDocument document) => Read(document).Descendants(w + "tbl").Select(t => new {
    grid = t.Element(w + "tblGrid")?.Elements(w + "gridCol").Select(c => Val(c,"w")).ToArray(),
    rows = t.Elements(w + "tr").Select(r => new {
        before = Val(r.Element(w+"trPr")?.Element(w+"gridBefore"),"val") ?? "0",
        after = Val(r.Element(w+"trPr")?.Element(w+"gridAfter"),"val") ?? "0",
        beforeWidth = Val(r.Element(w+"trPr")?.Element(w+"wBefore"),"w") ?? "0",
        afterWidth = Val(r.Element(w+"trPr")?.Element(w+"wAfter"),"w") ?? "0",
        cells = r.Elements(w+"tc").Select(c => new {
            width = Val(c.Element(w+"tcPr")?.Element(w+"tcW"),"w"),
            span = Val(c.Element(w+"tcPr")?.Element(w+"gridSpan"),"val") ?? "1",
            merge = c.Element(w+"tcPr")?.Element(w+"vMerge") is { } m ? Val(m,"val") ?? "continue" : "none",
            text = string.Concat(c.Descendants(w+"t").Select(x=>x.Value))
        }).ToArray()
    }).ToArray()
}).ToArray();
string Signature(WmlDocument d) => JsonSerializer.Serialize(Geometry(d));
string[] Validate(WmlDocument d) {
    using var ms = new MemoryStream(d.DocumentByteArray);
    using var pkg = WordprocessingDocument.Open(ms,false);
    return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(pkg).Select(e => e.Description + " @ " + e.Path?.XPath).ToArray();
}
var cases = new List<(string Name,string Left,string Right)>();
int[] a = new[]{1800,2300,3100}, b = new[]{2700,1600,2900};
string Plain(int[] g,string prefix) => R(C(prefix+" gauge",g[0])+C(prefix+" bridge",g[1])+C(prefix+" lock",g[2]));
cases.Add(("widths_only",T(a,Plain(a,"North"),Plain(a,"South")),T(b,Plain(b,"North"),Plain(b,"South"))));
cases.Add(("widths_row_delete",T(a,Plain(a,"North"),Plain(a,"Middle"),Plain(a,"South")),T(b,Plain(b,"North"),Plain(b,"South"))));
cases.Add(("widths_row_insert",T(a,Plain(a,"North"),Plain(a,"South")),T(b,Plain(b,"North"),Plain(b,"Middle"),Plain(b,"South"))));
cases.Add(("merge_horizontal",T(a,Plain(a,"North"),Plain(a,"South")),T(b,R(C("North gauge bridge",4300,2)+C("North lock",2900)),Plain(b,"South"))));
cases.Add(("split_horizontal",cases[^1].Right,cases[^1].Left));
cases.Add(("offset_shift",T(a,R(C("North bridge",2300)+C("North lock",3100),1,0,1800),Plain(a,"South")),T(b,R(C("North bridge",2700)+C("North lock",1600),0,1,0,2900),Plain(b,"South"))));
cases.Add(("offset_add_cell",T(a,R(C("North bridge",2300)+C("North lock",3100),1,0,1800),Plain(a,"South")),T(b,Plain(b,"North"),Plain(b,"South"))));
cases.Add(("offset_remove_cell",cases[^1].Right,cases[^1].Left));
string Vertical(int[] g,bool merged) => T(g,R(C("North gauge",g[0],1,merged?"restart":null)+C("North bridge",g[1])+C("North lock",g[2])),R(C(merged?"":"South gauge",g[0],1,merged?"continue":null)+C("South bridge",g[1])+C("South lock",g[2])));
cases.Add(("merge_vertical",Vertical(a,false),Vertical(b,true)));
cases.Add(("split_vertical",Vertical(a,true),Vertical(b,false)));
cases.Add(("merged_widths_row_insert",T(a,R(C("North combined",4100,2)+C("North lock",3100)),Plain(a,"South")),T(b,R(C("North combined",4300,2)+C("North lock",2900)),Plain(b,"Middle"),Plain(b,"South"))));
cases.Add(("merged_widths_row_delete",cases[^1].Right,cases[^1].Left));
cases.Add(("merge_horizontal_stable_widths",T(a,Plain(a,"North"),Plain(a,"South")),T(a,R(C("North gauge bridge",4100,2)+C("North lock",3100)),Plain(a,"South"))));
cases.Add(("offset_remove_stable_widths",T(a,Plain(a,"North"),Plain(a,"South")),T(a,R(C("North bridge",2300)+C("North lock",3100),1,0,1800),Plain(a,"South"))));
var results = new List<object>();
foreach(var item in cases) {
    var dir = Path.Combine(output,item.Name); Directory.CreateDirectory(dir);
    var left = Make(item.Left); var right = Make(item.Right);
    var red = DocxCompare.Compare(left,right);
    var accepted = RevisionProcessor.AcceptRevisions(red); var rejected = RevisionProcessor.RejectRevisions(red);
    var docs = new[]{("left",left),("right",right),("redline",red),("accepted",accepted),("rejected",rejected)};
    foreach(var (name,doc) in docs) File.WriteAllBytes(Path.Combine(dir,name+".docx"),doc.DocumentByteArray);
    var result = new { name=item.Name, accepted_matches=Signature(right)==Signature(accepted), rejected_matches=Signature(left)==Signature(rejected),
        validation=docs.ToDictionary(x=>x.Item1,x=>Validate(x.Item2)), geometry=docs.Where(x=>x.Item1!="redline").ToDictionary(x=>x.Item1,x=>Geometry(x.Item2)),
        histories=Read(red).Descendants().Where(x=>new[]{"tblGridChange","tblPrChange","trPrChange","tcPrChange","cellIns","cellDel","cellMerge"}.Contains(x.Name.LocalName)).GroupBy(x=>x.Name.LocalName).ToDictionary(g=>g.Key,g=>g.Count()) };
    results.Add(result); Console.WriteLine($"{item.Name}: accept={result.accepted_matches} reject={result.rejected_matches} schema={string.Join(',',result.validation.Select(x=>x.Key+":"+x.Value.Length))}");
}
File.WriteAllText(Path.Combine(output,"results.json"),JsonSerializer.Serialize(new{assembly_sha256=Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(typeof(DocxCompare).Assembly.Location))).ToLowerInvariant(),results},new JsonSerializerOptions{WriteIndented=true}));
