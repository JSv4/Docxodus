#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Pins the explicit, highly conservative WordRepairCompatibility path. The fixture deliberately models the
/// raw repair signature rather than a normal text edit: accepted body content is identical, old revisions are
/// identical, 48/82 para ids churn, two table shells are normalized, and eight styles gain only implicit
/// defaults while their order is regenerated. This keeps the test close to the repaired-output corpus without
/// depending on its external files.
/// </summary>
public class WordRepairCompatibilityTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

    [Fact]
    public void OptIn_ProjectsRepairChurnAsWholeBodyReplacement_AndRoundTrips()
    {
        var (left, right) = Pair();
        var plain = new DocxDiffSettings { AuthorForRevisions = "Compat" };
        var compat = new DocxDiffSettings { AuthorForRevisions = "Compat", WordRepairCompatibility = true };

        Assert.True(Detects(left, right));
        Assert.Equal(
            DocxDiff.GetEditScriptJson(left, right, plain),
            DocxDiff.GetEditScriptJson(left, right, compat));

        var ordinary = DocxDiff.Compare(left, right, plain);
        var projected = DocxDiff.Compare(left, right, compat);

        Assert.Empty(Body(ordinary).Descendants(W + "ins").Where(e => (string?)e.Attribute(W + "author") == "Compat"));
        Assert.Empty(Body(ordinary).Descendants(W + "del").Where(e => (string?)e.Attribute(W + "author") == "Compat"));
        Assert.NotEmpty(Body(projected).Descendants(W + "ins").Where(e => (string?)e.Attribute(W + "author") == "Compat"));
        Assert.NotEmpty(Body(projected).Descendants(W + "del").Where(e => (string?)e.Attribute(W + "author") == "Compat"));

        var accepted = RevisionProcessor.AcceptRevisions(projected);
        var rejected = RevisionProcessor.RejectRevisions(projected);
        Assert.Equal(NormalizedBody(RevisionProcessor.AcceptRevisions(right)), NormalizedBody(accepted));
        Assert.Equal(NormalizedBody(RevisionProcessor.AcceptRevisions(left)), NormalizedBody(rejected));

        // The table-shell cleanup is not merely text-equivalent: accept receives the right Normal style and
        // reject restores the left repair-era indentation/cell margins.
        Assert.Equal(2, Body(accepted).Descendants(W + "tbl").Count());
        Assert.Equal(2, Body(rejected).Descendants(W + "tbl").Count());
        Assert.Equal(2, Body(accepted).Descendants(W + "tblStyle")
            .Count(e => (string?)e.Attribute(W + "val") == "Normal"));
        Assert.Empty(Body(accepted).Descendants(W + "tblInd"));
        Assert.Equal(2, Body(rejected).Descendants(W + "tblInd").Count());
        Assert.Equal(4, Body(rejected).Descendants(W + "tblCellMar").Count());
    }

    [Fact]
    public void Detector_RejectsNearMisses()
    {
        // Two table-cell paragraphs also carry paraIds, so 61 body paragraphs remains below the 64-id floor.
        Assert.False(Detects(Pair(new FixtureOptions(ParagraphCount: 61, ChangedParaIds: 48))));
        Assert.False(Detects(Pair(new FixtureOptions(ChangedParaIds: 31))));
        Assert.False(Detects(Pair(new FixtureOptions(ChangeVisibleText: true))));
        Assert.False(Detects(Pair(new FixtureOptions(MismatchHistoricalRevision: true))));
        Assert.False(Detects(Pair(new FixtureOptions(NonRepairTableDelta: true))));
        Assert.False(Detects(Pair(new FixtureOptions(NonDefaultStyleDelta: true))));
    }

    [Fact]
    public void Detector_TreatsEmptyAuxiliaryScopeListsAsNoOps_ButRejectsChangedScopes()
    {
        var probe = BuildProbe(Pair());
        var emptyScopes = probe.Script with
        {
            NoteOps = IrNodeList.Empty<IrNoteDiff>(),
            HeaderFooterOps = IrNodeList.Empty<IrHeaderFooterDiff>(),
        };

        Assert.True(WordRepairCompatibilityProjection.ShouldRenderWholeBodyReplacement(
            emptyScopes, probe.Left, probe.Right, probe.RawLeft, probe.RawRight));

        var changedNote = emptyScopes with
        {
            NoteOps = IrNodeList.From(new[]
            {
                new IrNoteDiff(IrNoteKind.Footnote, "1", IrNodeList.Empty<IrEditOp>()),
            }),
        };
        Assert.False(WordRepairCompatibilityProjection.ShouldRenderWholeBodyReplacement(
            changedNote, probe.Left, probe.Right, probe.RawLeft, probe.RawRight));

        var changedHeader = emptyScopes with
        {
            HeaderFooterOps = IrNodeList.From(new[]
            {
                new IrHeaderFooterDiff(true, IrHeaderFooterKind.Default, 0, "hdr1", "hdr1",
                    new Uri("/word/header1.xml", UriKind.Relative),
                    new Uri("/word/header1.xml", UriKind.Relative), IrNodeList.Empty<IrEditOp>()),
            }),
        };
        Assert.False(WordRepairCompatibilityProjection.ShouldRenderWholeBodyReplacement(
            changedHeader, probe.Left, probe.Right, probe.RawLeft, probe.RawRight));
    }

    private static bool Detects((WmlDocument Left, WmlDocument Right) pair) => Detects(pair.Left, pair.Right);

    private static bool Detects(WmlDocument left, WmlDocument right)
    {
        var probe = BuildProbe(left, right);
        return WordRepairCompatibilityProjection.ShouldRenderWholeBodyReplacement(
            probe.Script, probe.Left, probe.Right, probe.RawLeft, probe.RawRight);
    }

    private static RepairProbe BuildProbe((WmlDocument Left, WmlDocument Right) pair) => BuildProbe(pair.Left, pair.Right);

    private static RepairProbe BuildProbe(WmlDocument left, WmlDocument right)
    {
        var settings = new DocxDiffSettings().ToIrDiffSettings();
        var options = new IrReaderOptions { RevisionView = RevisionView.Accept };
        var irLeft = IrReader.Read(left, options);
        var irRight = IrReader.Read(right, options);
        var script = IrEditScriptBuilder.Build(irLeft, irRight, settings);
        return new RepairProbe(script, irLeft, irRight, left, right);
    }

    private sealed record RepairProbe(
        IrEditScript Script,
        IrDocument Left,
        IrDocument Right,
        WmlDocument RawLeft,
        WmlDocument RawRight);

    private sealed record FixtureOptions(
        int ParagraphCount = 80,
        int ChangedParaIds = 48,
        bool ChangeVisibleText = false,
        bool MismatchHistoricalRevision = false,
        bool NonRepairTableDelta = false,
        bool NonDefaultStyleDelta = false,
        bool ReorderStyles = true);

    private static (WmlDocument Left, WmlDocument Right) Pair(FixtureOptions? options = null)
    {
        options ??= new FixtureOptions();
        return (Doc(false, options), Doc(true, options));
    }

    private static WmlDocument Doc(bool right, FixtureOptions options)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            WriteXml(main, DocumentXml(right, options));
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            WriteXml(styles, StylesXml(right, options));
            var settings = main.AddNewPart<DocumentSettingsPart>();
            WriteXml(settings, $"<w:settings xmlns:w=\"{W}\"/>");
        }
        return new WmlDocument(right ? "right.docx" : "left.docx", stream.ToArray());
    }

    private static void WriteXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(xml);
    }

    private static string DocumentXml(bool right, FixtureOptions options)
    {
        var body = new StringBuilder();
        for (int i = 0; i < options.ParagraphCount; i++)
        {
            body.Append($"<w:p w14:paraId=\"{ParaId(right, i, options)}\" w14:textId=\"77777777\">");
            body.Append($"<w:r><w:t>Paragraph {i}</w:t></w:r>");
            if (i == 0)
            {
                var author = right && options.MismatchHistoricalRevision ? "OtherPrior" : "Prior";
                body.Append($"<w:del w:id=\"41\" w:author=\"{author}\" w:date=\"2020-01-01T00:00:00Z\"><w:r><w:delText> old backend</w:delText></w:r></w:del>");
                body.Append($"<w:ins w:id=\"42\" w:author=\"{author}\" w:date=\"2020-01-01T00:00:00Z\"><w:r><w:t> no backend</w:t></w:r></w:ins>");
            }
            if (right && options.ChangeVisibleText && i == options.ParagraphCount - 1)
                body.Append("<w:r><w:t> changed</w:t></w:r>");
            body.Append("</w:p>");

            // Keep tables non-adjacent so accepting source revisions cannot coalesce them.
            if (i == 20 || i == 50)
                body.Append(TableXml(right, i == 20 ? 0 : 1, options));
        }
        body.Append("<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>");
        return $"<w:document xmlns:w=\"{W}\" xmlns:w14=\"{W14}\"><w:body>{body}</w:body></w:document>";
    }

    private static string ParaId(bool right, int index, FixtureOptions options)
    {
        int baseId = index < options.ChangedParaIds
            ? right ? 0x20000000 : 0x10000000
            : 0x30000000;
        return (baseId + index).ToString("X8");
    }

    private static string TableXml(bool right, int ordinal, FixtureOptions options)
    {
        var tblW = right && options.NonRepairTableDelta ? "5000" : "3000";
        var props = right
            ? $"<w:tblStyle w:val=\"Normal\"/><w:tblW w:w=\"{tblW}\" w:type=\"dxa\"/>"
            : $"<w:tblInd w:w=\"5\" w:type=\"dxa\"/><w:tblW w:w=\"{tblW}\" w:type=\"dxa\"/><w:tblCellMar><w:left w:w=\"10\" w:type=\"dxa\"/><w:right w:w=\"10\" w:type=\"dxa\"/></w:tblCellMar>";
        var rowEx = right ? string.Empty : "<w:tblPrEx><w:tblCellMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/></w:tblCellMar></w:tblPrEx>";
        return $"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid>" +
               $"<w:tr>{rowEx}<w:tc><w:tcPr/><w:p w14:paraId=\"{(0x40000000 + ordinal).ToString("X8")}\" w14:textId=\"77777777\"><w:r><w:t>Table {ordinal}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
    }

    private static string StylesXml(bool right, FixtureOptions options)
    {
        var styles = new StringBuilder($"<w:styles xmlns:w=\"{W}\">");
        var styleIds = Enumerable.Range(0, 8).ToList();
        if (right && options.ReorderStyles)
        {
            // Captured repair pairs move a definition while materializing the same default inheritance.
            styleIds.Remove(7);
            styleIds.Insert(0, 7);
        }
        foreach (var i in styleIds)
        {
            styles.Append($"<w:style w:type=\"paragraph\" w:styleId=\"S{i}\"><w:name w:val=\"S{i}\"/>");
            if (right)
            {
                styles.Append("<w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/>");
                if (options.NonDefaultStyleDelta && i == 0)
                    styles.Append("<w:rPr><w:b/></w:rPr>");
            }
            styles.Append("</w:style>");
        }
        styles.Append("</w:styles>");
        return styles.ToString();
    }

    private static XElement Body(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wordDoc = WordprocessingDocument.Open(stream, false);
        return new XElement(wordDoc.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!);
    }

    private static string NormalizedBody(WmlDocument doc)
    {
        var body = Body(doc);
        // Whole-table revision marking creates an empty trPr shell when the source row had none. It is
        // schema-equivalent to its absence after accept/reject, so normalize that renderer bookkeeping
        // before checking the content-level round-trip contract.
        foreach (var trPr in body.Descendants(W + "trPr").Where(e => !e.HasAttributes && !e.HasElements).ToList())
            trPr.Remove();
        foreach (var attr in body.DescendantsAndSelf().Attributes().ToList())
        {
            if ((attr.Name.Namespace == W14 && (attr.Name.LocalName == "paraId" || attr.Name.LocalName == "textId")) ||
                (attr.Name.Namespace == W && attr.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)))
            {
                attr.Remove();
            }
        }
        return body.ToString(SaveOptions.DisableFormatting);
    }
}
