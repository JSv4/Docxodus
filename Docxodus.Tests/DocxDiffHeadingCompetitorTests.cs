using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffHeadingCompetitorTests
{
    private const string OldBody = "Bronze radar channel cabinet enabled.";
    private const string NewBody = "Bronze radar channel cabinet disabled.";
    private const string NewHeading = "Bronze radar channel cabinet service calendar for the coastal crew.";
    private const string OldSecond = "Silver receiver voltage cabinet active.";
    private const string NewSecond = "Silver receiver voltage cabinet idle.";
    private static readonly XNamespace W = IrTestDocuments.W;

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InsertedHeading_DoesNotConsumeTheStrongerBodyEdit(bool headingStyle)
    {
        var left = Document((OldBody, false), (OldSecond, false));
        var right = Document((NewHeading, headingStyle), (NewBody, false), (NewSecond, false));
        var l = IrReader.Read(left, DocxDiff.ReadOpts);
        var r = IrReader.Read(right, DocxDiff.ReadOpts);
        var alignment = IrBlockAligner.Align(l, r, new IrDiffSettings());
        Assert.Contains(alignment.Entries, e => e.Kind == IrAlignmentKind.Modified && Text(e.Left) == OldBody && Text(e.Right) == NewBody);
        Assert.Contains(alignment.Entries, e => e.Kind == IrAlignmentKind.Modified && Text(e.Left) == OldSecond && Text(e.Right) == NewSecond);
        Assert.Contains(alignment.Entries, e => e.Kind == IrAlignmentKind.Inserted && Text(e.Right) == NewHeading);
        Assert.Contains(alignment.Entries, e => e.Kind == IrAlignmentKind.Unchanged && Text(e.Left) == "Teal register opens here.");
        Assert.Contains(alignment.Entries, e => e.Kind == IrAlignmentKind.Unchanged && Text(e.Left) == "Teal register closes here.");
        var comparison = AssertEndpoints(left, right);
        using var stream = new MemoryStream(comparison.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        using var xml = package.MainDocumentPart!.GetStream();
        var heading = XDocument.Load(xml).Descendants(W + "p").Single(p => string.Concat(p.Descendants(W + "t").Select(t => t.Value)) == NewHeading);
        Assert.NotNull(heading.Element(W + "pPr")?.Element(W + "rPr")?.Element(W + "ins"));
    }

    [Fact]
    public void GenuineHeadingEdit_RemainsPaired()
    {
        var left = Document(("Bronze radar channel cabinet service calendar.", true), (OldBody, false));
        var right = Document(("Bronze radar channel cabinet winter calendar.", true), (NewBody, false));
        var alignment = IrBlockAligner.Align(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        Assert.Contains(alignment.Entries, e => e.Kind == IrAlignmentKind.Modified && Text(e.Left).Contains("service calendar") && Text(e.Right).Contains("winter calendar"));
        AssertEndpoints(left, right);
    }

    [Fact]
    public void BodyEditsWithoutAnInsertion_RemainPaired()
    {
        var left = Document((OldBody, false), (OldSecond, false));
        var right = Document((NewBody, false), (NewSecond, false));
        var alignment = IrBlockAligner.Align(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        Assert.Contains(alignment.Entries, e => Text(e.Left) == OldBody && Text(e.Right) == NewBody);
        Assert.Contains(alignment.Entries, e => Text(e.Left) == OldSecond && Text(e.Right) == NewSecond);
        AssertEndpoints(left, right);
    }

    [Fact]
    public void EqualNormalizedEvidence_KeepsThePositionalPreference()
    {
        const string equalCandidate = "Bronze radar channel cabinet standby.";
        var left = Document((OldBody, false), (OldSecond, false));
        var right = Document((equalCandidate, false), (NewBody, false), (NewSecond, false));
        var alignment = IrBlockAligner.Align(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        Assert.Contains(alignment.Entries, e => e.Kind == IrAlignmentKind.Modified && Text(e.Left) == OldBody && Text(e.Right) == equalCandidate);
        AssertEndpoints(left, right);
    }

    [Fact]
    public void ParagraphBecomingAHeading_StillPairsWithoutAStrongerCompetitor()
    {
        var left = Document((OldBody, false));
        var right = Document((NewBody, true));
        var alignment = IrBlockAligner.Align(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        Assert.Contains(alignment.Entries, e => e.Kind == IrAlignmentKind.Modified && Text(e.Left) == OldBody && Text(e.Right) == NewBody);
        AssertEndpoints(left, right);
    }

    [Fact]
    public void UnrelatedRewrite_HasNoInventedBodyCorrespondence()
    {
        var left = Document((OldBody, false));
        var right = Document(("Orange kettle sings at dusk.", false));
        var alignment = IrBlockAligner.Align(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        Assert.Contains(alignment.Entries, e => e.Kind == IrAlignmentKind.Deleted && Text(e.Left) == OldBody);
        Assert.Contains(alignment.Entries, e => e.Kind == IrAlignmentKind.Inserted && Text(e.Right) == "Orange kettle sings at dusk.");
        AssertEndpoints(left, right);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SplitAndMerge_RetainTheirParagraphEndpoints(bool merge)
    {
        var one = Document((OldBody + " " + OldSecond, false));
        var two = Document((OldBody + " ", false), (OldSecond, false));
        AssertEndpoints(merge ? two : one, merge ? one : two);
    }

    [Fact]
    public void NormalizedTieIndex_AgreesWithAnExhaustiveScan()
    {
        var random = new System.Random(1052);
        string[] words = { "bronze", "silver", "radar", "channel", "cabinet", "beacon", "winter", "coastal", "calendar", "disabled" };
        string Sentence() => string.Join(" ", Enumerable.Range(0, random.Next(1, 10)).Select(_ => words[random.Next(words.Length)]));
        var blocks = IrReader.Read(IrTestDocuments.Create(Enumerable.Range(0, 30).Select(_ => Sentence()).ToArray())).Body.Blocks;
        var targets = IrReader.Read(IrTestDocuments.Create(Enumerable.Range(0, 20).Select(_ => Sentence()).ToArray())).Body.Blocks;
        var similarity = new IrBlockSimilarity(new IrDiffSettings());
        var indices = Enumerable.Range(0, blocks.Count).ToList();
        var index = new IrBlockAligner.ContentWordIndex(blocks, indices, similarity);
        int Count(IrParagraph paragraph) => similarity.PairingWordKeys(paragraph).Keys.Count(k => !IrBlockAligner.IsFunctionWordKey(k));
        int wins = 0;
        for (int trial = 0; trial < 2000; trial++)
        {
            var target = (IrParagraph)targets[random.Next(targets.Count)];
            int partner = random.Next(blocks.Count);
            int evidence = IrBlockAligner.SharedContentWordCount(target, (IrParagraph)blocks[partner], similarity);
            if (evidence == 0) continue;
            int lower = random.Next(-1, blocks.Count), upper = random.Next(lower + 1, blocks.Count + 1);
            var match = Enumerable.Range(0, blocks.Count).Select(_ => random.Next(4) == 0 ? 0 : -1).ToArray();
            double slotScore = (double)evidence / (Count(target) + Count((IrParagraph)blocks[partner]) - evidence);
            bool scanned = indices.Any(c =>
            {
                if (c == partner || match[c] != -1 || c <= lower || c >= upper) return false;
                int shared = IrBlockAligner.SharedContentWordCount(target, (IrParagraph)blocks[c], similarity);
                return shared == evidence && (double)shared / (Count(target) + Count((IrParagraph)blocks[c]) - shared) > slotScore;
            });
            Assert.Equal(scanned, index.OutbidsOnNormalizedTie(target, partner, evidence, match, similarity, lower, upper));
            if (scanned) wins++;
        }
        Assert.True(wins > 20);
    }

    private static WmlDocument Document(params (string Text, bool Heading)[] paragraphs)
    {
        string P(string text, bool heading = false) => "<w:p>" + (heading ? "<w:pPr><w:pStyle w:val=\"SignalHeading\"/></w:pPr>" : "") +
            $"<w:r><w:t>{text}</w:t></w:r></w:p>";
        var body = P("Teal register opens here.") + string.Concat(paragraphs.Select(p => P(p.Text, p.Heading))) + P("Teal register closes here.");
        var styles = "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"SignalHeading\"><w:name w:val=\"Signal Heading\"/><w:basedOn w:val=\"Normal\"/>" +
            "<w:pPr><w:keepNext/><w:outlineLvl w:val=\"1\"/></w:pPr><w:rPr><w:sz w:val=\"28\"/></w:rPr></w:style>";
        return IrTestDocuments.FromBodyAndStylesXml(body, styles);
    }

    private static string Text(IrBlock? block) => block is IrParagraph p ? string.Concat(p.Inlines.OfType<IrTextRun>().Select(r => r.Text)) : "";

    private static string[] Paragraphs(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        using var xml = package.MainDocumentPart!.GetStream();
        return XDocument.Load(xml).Descendants(W + "p").Select(p => string.Concat(p.Descendants(W + "t").Select(t => t.Value))).ToArray();
    }

    private static WmlDocument AssertEndpoints(WmlDocument left, WmlDocument right)
    {
        var comparison = DocxCompare.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(comparison);
        var rejected = RevisionProcessor.RejectRevisions(comparison);
        Assert.Equal(Paragraphs(right), Paragraphs(accepted));
        Assert.Equal(Paragraphs(left), Paragraphs(rejected));
        foreach (var document in new[] { left, right, comparison, accepted, rejected })
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package));
        }
        return comparison;
    }
}
