using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffStyleParentTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    [Theory]
    [InlineData(false, false, 0)]
    [InlineData(true, false, 0)]
    [InlineData(false, true, 0)]
    [InlineData(true, true, 0)]
    [InlineData(false, false, 2)]
    [InlineData(true, true, 2)]
    [InlineData(false, false, 20)]
    public void ChangedParent_PreservesEffectiveSpacingInBothViews(bool italic, bool childOverride, int depth)
    {
        var left = Document("Compact", false, childOverride, depth);
        var right = Document("Airy", italic, childOverride, depth);
        var result = DocxCompare.Compare(left, right);
        AssertEffective(RevisionProcessor.AcceptRevisions(result), "Normal", childOverride ? "60" : "300", "420");
        AssertEffective(RevisionProcessor.RejectRevisions(result), "Normal", childOverride ? "60" : "0", "240");
        AssertEffective(RevisionProcessor.AcceptRevisions(result), "Descendant", "90", "420");
        AssertEffective(RevisionProcessor.RejectRevisions(result), "Descendant", "90", "240");
        Assert.True(XNode.DeepEquals(
            Canonical(Read(left).Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == "Unrelated")),
            Canonical(Read(result).Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == "Unrelated"))));
        var acceptedRun = Read(RevisionProcessor.AcceptRevisions(result)).Elements(W + "style")
            .Single(s => (string?)s.Attribute(W + "styleId") == "Normal").Element(W + "rPr");
        Assert.Equal(italic, acceptedRun?.Element(W + "i") is not null);
        Assert.Null(Read(RevisionProcessor.RejectRevisions(result)).Elements(W + "style")
            .Single(s => (string?)s.Attribute(W + "styleId") == "Normal").Element(W + "rPr")?.Element(W + "i"));
        AssertClean(result);
        AssertClean(RevisionProcessor.AcceptRevisions(result));
        AssertClean(RevisionProcessor.RejectRevisions(result));
    }

    private static XElement Canonical(XElement element) => new(element.Name,
        element.Attributes().Where(a => !a.IsNamespaceDeclaration), element.Elements().Select(Canonical));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParentChange_PreservesRunToggleComposition(bool childBold)
    {
        var left = Document("Compact", false, false, 0, "<w:b/>", childBold ? "<w:b/>" : "");
        var right = Document("Airy", false, false, 0, "<w:b/>", childBold ? "<w:b/>" : "");
        var result = DocxCompare.Compare(left, right);
        Assert.Equal(childBold, EffectiveBold(Read(RevisionProcessor.AcceptRevisions(result)), "Normal"));
        Assert.Equal(!childBold, EffectiveBold(Read(RevisionProcessor.RejectRevisions(result)), "Normal"));
        AssertClean(result);
    }

    [Fact]
    public void DescendantWithItsOwnRunChange_StillTracksTheChangedAncestor()
    {
        var left = Document("Compact", false, false, 0, "<w:b/>");
        var right = Document("Airy", false, false, 0, "<w:b/>", descendantRun: "<w:i/>");
        var result = DocxCompare.Compare(left, right);
        Assert.False(EffectiveBold(Read(RevisionProcessor.AcceptRevisions(result)), "Descendant"));
        Assert.True(EffectiveBold(Read(RevisionProcessor.RejectRevisions(result)), "Descendant"));
        AssertEffective(RevisionProcessor.AcceptRevisions(result), "Descendant", "90", "420");
        AssertEffective(RevisionProcessor.RejectRevisions(result), "Descendant", "90", "240");
        AssertClean(result);
    }

    private static bool EffectiveBold(XElement styles, string id)
    {
        var bold = false;
        var seen = new HashSet<string>();
        while (seen.Add(id))
        {
            var style = styles.Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == id);
            var b = style.Element(W + "rPr")?.Element(W + "b");
            bold ^= b is not null && (string?)b.Attribute(W + "val") is not ("0" or "false" or "off");
            if ((string?)style.Element(W + "basedOn")?.Attribute(W + "val") is not { } parent) break;
            id = parent;
        }
        return bold;
    }

    private static WmlDocument Document(string parent, bool italic, bool childOverride, int depth,
        string compactRun = "", string normalRun = "", string descendantRun = "")
    {
        string Style(string id, string? basedOn, string pPr, string rPr = "", bool isDefault = false) =>
            $"<w:style w:type=\"paragraph\" w:styleId=\"{id}\"{(isDefault ? " w:default=\"1\"" : "")}><w:name w:val=\"{id}\"/>" +
            (basedOn is null ? "" : $"<w:basedOn w:val=\"{basedOn}\"/>") + $"<w:pPr>{pPr}</w:pPr><w:rPr>{rPr}</w:rPr></w:style>";
        var styles = "<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>" +
            Style("Compact", null, "<w:spacing w:after=\"0\" w:line=\"240\"/>", compactRun) +
            Style("Airy", null, "<w:spacing w:after=\"300\" w:line=\"420\"/>");
        for (var i = 0; i < depth; i++)
        {
            styles += Style("Compact" + i, i == 0 ? "Compact" : "Compact" + (i - 1), "");
            styles += Style("Airy" + i, i == 0 ? "Airy" : "Airy" + (i - 1), "");
        }
        styles += Style("Normal", depth == 0 ? parent : parent + (depth - 1),
            childOverride ? "<w:spacing w:after=\"60\"/>" : "", (italic ? "<w:i/>" : "") + normalRun, true);
        styles += Style("Descendant", "Normal", "<w:spacing w:after=\"90\"/>", descendantRun);
        styles += Style("Unrelated", null, "<w:keepNext/>");
        return IrTestDocuments.FromBodyAndStylesXml(
            "<w:p><w:r><w:t>Harbor nursery record.</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:pStyle w:val=\"Descendant\"/></w:pPr><w:r><w:t>Second nursery record.</w:t></w:r></w:p>", styles);
    }

    private static void AssertEffective(WmlDocument doc, string id, string after, string line)
    {
        var root = Read(doc);
        var seen = new HashSet<string>();
        string? Attribute(string name)
        {
            var current = id;
            seen.Clear();
            while (seen.Add(current))
            {
                var style = root.Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == current);
                if ((string?)style.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + name) is { } value)
                    return value;
                if ((string?)style.Element(W + "basedOn")?.Attribute(W + "val") is not { } parent) break;
                current = parent;
            }
            return (string?)root.Element(W + "docDefaults")?.Element(W + "pPrDefault")?.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + name);
        }
        Assert.Equal(after, Attribute("after"));
        Assert.Equal(line, Attribute("line"));
    }

    private static XElement Read(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        using var part = package.MainDocumentPart!.StyleDefinitionsPart!.GetStream();
        return XDocument.Load(part).Root!;
    }

    private static void AssertClean(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package));
    }
}
