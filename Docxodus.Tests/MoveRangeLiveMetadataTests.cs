using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class MoveRangeLiveMetadataTests
{
    [Theory]
    [InlineData("bookmark", false)]
    [InlineData("bookmark", true)]
    [InlineData("content-control", false)]
    [InlineData("content-control", true)]
    [InlineData("field", false)]
    [InlineData("field", true)]
    public void ResolvingMovedText_PreservesUntrackedContentWithoutRuns(string kind, bool accept)
    {
        var path = Path.Combine(AppContext.BaseDirectory,
            "../../../../TestFiles/RP/MoveRangeMembership/unsplit-control.docx");
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(path));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!;
            var xml = part.GetXDocument();
            var end = xml.Descendants(accept ? W.moveFromRangeEnd : W.moveToRangeEnd).Single();
            var content = kind switch
            {
                "bookmark" => new[]
                {
                    new XElement(W.bookmarkStart, new XAttribute(W.id, "20"), new XAttribute(W.name, "LiveAnchor")),
                    new XElement(W.bookmarkEnd, new XAttribute(W.id, "20")),
                },
                "content-control" => new[]
                {
                    new XElement(W.sdt,
                        new XElement(W.sdtPr, new XElement(W.tag, new XAttribute(W.val, "LiveControl"))),
                        new XElement(W.sdtContent)),
                },
                _ => new[] { new XElement(W.fldSimple, new XAttribute(W.instr, " DATE ")) },
            };
            end.AddBeforeSelf(content);
            part.PutXDocument();
        }
        var bytes = stream.ToArray();
        PackageValidation.AssertValid(bytes, FileFormatVersions.Microsoft365);
        var input = new WmlDocument("live.docx", bytes);
        AssertKept((accept ? RevisionProcessor.AcceptRevisions(input) : RevisionProcessor.RejectRevisions(input))
            .DocumentByteArray, kind);
        using var session = new DocxSession(bytes);
        var result = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
        Assert.True(result.Success, result.Error?.Message);
        AssertKept(session.Save(), kind);
    }

    private static void AssertKept(byte[] bytes, string kind)
    {
        PackageValidation.AssertValid(bytes, FileFormatVersions.Microsoft365);
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var xml = document.MainDocumentPart!.GetXDocument();
        if (kind == "bookmark")
        {
            Assert.Equal("LiveAnchor", (string?)Assert.Single(xml.Descendants(W.bookmarkStart)).Attribute(W.name));
            Assert.Single(xml.Descendants(W.bookmarkEnd));
        }
        else if (kind == "content-control")
        {
            var control = Assert.Single(xml.Descendants(W.sdt));
            Assert.Equal("LiveControl", (string?)control.Element(W.sdtPr)!.Element(W.tag)!.Attribute(W.val));
        }
        else
        {
            Assert.Equal(" DATE ", (string?)Assert.Single(xml.Descendants(W.fldSimple)).Attribute(W.instr));
        }
    }
}
