using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlLinkedTextBoxTests
{
    private static byte[] Document(bool linked)
    {
        string Box(int id, bool continuation) => $"""
            <w:drawing><wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
              <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>{id * 200000}</wp:posOffset></wp:positionH>
              <wp:positionV relativeFrom="page"><wp:posOffset>1524000</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="762000"/>
              <wp:wrapNone/><wp:docPr id="{id}" name="Alder story frame {id}"/>
              <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>
                {(continuation ? "<wps:linkedTxbx id=\"37\" seq=\"1\"/>" : "<wps:txbx id=\"37\"><w:txbxContent><w:p><w:r><w:t>Alder linked story register.</w:t></w:r></w:p></w:txbxContent></wps:txbx>")}
                <wps:bodyPr lIns="38100" rIns="38100" tIns="38100" bIns="38100"><a:noAutofit/></wps:bodyPr>
              </wps:wsp></a:graphicData></a:graphic>
            </wp:anchor></w:drawing>
            """;
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
                writer.Write($"""
                    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                      xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body>
                      <w:p><w:r>{Box(1, false)}{(linked ? Box(2, true) : "")}</w:r><w:r><w:t>Alder frame heading.</w:t></w:r></w:p><w:sectPr/>
                    </w:body></w:document>
                    """);
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            package.Save();
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OnlyARealLinkedStoryRetainsAContinuationAndDisablesFixedClipping(bool linked)
    {
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(Document(linked),
            new HtmlConversionOptions { FabricateCssClasses = false, PaginationMode = (int)PaginationMode.Paginated }));
        var boxes = html.Descendants().Where(e => (string?)e.Attribute("data-docx-textbox") == "true").ToArray();
        Assert.Equal(linked ? 2 : 1, boxes.Length);
        Assert.Contains("Alder linked story register.", boxes[0].Value);
        if (linked)
        {
            Assert.All(boxes, box => Assert.Equal("true", (string?)box.Attribute("data-docx-textbox-linked")));
            Assert.All(boxes, box => Assert.Null(box.Attribute("data-docx-textbox-clip")));
            Assert.Equal((string?)boxes[0].Attribute("data-docx-textbox-story"), (string?)boxes[1].Attribute("data-docx-textbox-story"));
            Assert.Equal(new[] { "0", "1" }, boxes.Select(box => (string?)box.Attribute("data-docx-textbox-seq")));
            Assert.Empty(boxes[1].Elements());
            Assert.Contains("height: 60pt", (string?)boxes[1].Attribute("style"));
        }
        else
        {
            Assert.Null(boxes[0].Attribute("data-docx-textbox-linked"));
            Assert.Equal("true", (string?)boxes[0].Attribute("data-docx-textbox-clip"));
        }
    }
}
