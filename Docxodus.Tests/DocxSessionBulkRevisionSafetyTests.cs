using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1385: bulk accept/reject refuses a document only for markup it cannot resolve safely.
/// Revision metadata (author, date) and native ids serve review and addressing, never the
/// resolution itself, so a defect there stays visible on the listed entry but no longer stops
/// AcceptAllRevisions/RejectAllRevisions. Every expected text is cross-checked against
/// <see cref="RevisionProcessor"/>, the whole-document transform the session must agree with.
/// </summary>
public class DocxSessionBulkRevisionSafetyTests
{
    private const string Stamp = "w:author=\"Reviewer A\" w:date=\"2026-03-02T10:00:00Z\"";

    private static string Run(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static string Ins(string attributes, string text) => $"<w:ins {attributes}>{Run(text)}</w:ins>";

    private static string Del(string attributes, string text) =>
        $"<w:del {attributes}><w:r><w:delText xml:space=\"preserve\">{text}</w:delText></w:r></w:del>";

    private static string Paragraph(params string[] content) => "<w:p>" + string.Concat(content) + "</w:p>";

    // ─── Metadata and addressing defects ─────────────────────────────────

    public static TheoryData<string, string, string> MetadataDefects => new()
    {
        { "no author", "w:id=\"1\" w:date=\"2026-03-02T10:00:00Z\"", "missing_revision_author" },
        { "unparseable date", "w:id=\"1\" w:author=\"Reviewer A\" w:date=\"0-00-00T00:00:00Z\"", "invalid_revision_date" },
        { "out-of-range date", "w:id=\"1\" w:author=\"Reviewer A\" w:date=\"2026-00-00T28:01:00Z\"", "invalid_revision_date" },
        {
            "UTC date without Z",
            "w:id=\"1\" w:author=\"Reviewer A\" w:date=\"2026-03-02T10:00:00Z\" "
                + "xmlns:w16du=\"http://schemas.microsoft.com/office/word/2023/wordml/word16du\" "
                + "w16du:dateUtc=\"2026-03-02T10:00:00\"",
            "invalid_revision_date_utc"
        },
    };

    [Theory]
    [MemberData(nameof(MetadataDefects))]
    public void MetadataDefect_StaysListedButDoesNotBlockResolution(string shape, string attributes, string code)
    {
        var body = Paragraph(Run("Terms "), Ins(attributes, "apply "), Run("today"))
            + Paragraph(Run("Second "), Del("w:id=\"2\" " + Stamp, "clause "), Run("ends"));
        var bytes = Build(body);

        using (var session = new DocxSession(bytes))
        {
            var listed = session.ListRevisions();
            var defective = Assert.Single(listed, entry => entry.Diagnostic is not null);
            Assert.True(defective.ResolutionStatus == RevisionResolutionStatus.Supported, shape);
            Assert.Equal(code, defective.Diagnostic!.Code);
            Assert.Null(Assert.Single(listed, entry => entry.Type == "delete").Diagnostic);

        }

        // One revision addressed by its id resolves too, in both directions: no step of the
        // resolution reads the author or the date.
        foreach (bool accept in new[] { true, false })
        {
            using var session = new DocxSession(bytes);
            var defective = Assert.Single(session.ListRevisions(), entry => entry.Diagnostic is not null);
            var result = accept ? session.AcceptRevision(defective.Id) : session.RejectRevision(defective.Id);
            Assert.True(result.Success, shape);
            Assert.Equal("delete", Assert.Single(session.ListRevisions()).Type);
            Assert.Equal(accept ? new[] { "Terms apply today", "Second ends" } : new[] { "Terms today", "Second ends" },
                ParagraphTexts(session.Save()));
        }

        AssertBulkMatchesProcessor(bytes,
            accepted: new[] { "Terms apply today", "Second ends" },
            rejected: new[] { "Terms today", "Second clause ends" });
    }

    [Fact]
    public void MetadataDefect_IsReportedByDeliverableInspectionWithoutBlocking()
    {
        var bytes = Build(Paragraph(Run("Terms "), Ins("w:id=\"1\" w:date=\"2026-03-02T10:00:00Z\"", "apply "),
            Run("today")));

        var result = Docxodus.Verification.DeliverableVerifier.VerifyDeliverable(bytes);

        var finding = Assert.Single(result.Findings, finding => finding.Code.StartsWith("structure.revision_",
            System.StringComparison.Ordinal));
        Assert.Equal("structure.revision_advisory", finding.Code);
        Assert.Equal(Docxodus.Verification.VerificationFindingSeverity.Info, finding.Severity);
    }

    [Fact]
    public void MetadataDefect_DoesNotMaskAResolutionSafetyCheck()
    {
        // The deleted cell is the first of its row: no surviving predecessor can absorb its grid
        // column. That refusal must win over the missing author, which alone would not block.
        var body = "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:cellDel w:id=\"1\" w:date=\"2026-03-02T10:00:00Z\"/></w:tcPr>"
            + Paragraph(Run("Left")) + "</w:tc><w:tc>" + Paragraph(Run("Right")) + "</w:tc></w:tr></w:tbl>"
            + Paragraph(Run("After"));

        using var session = new DocxSession(Build(body));

        var entry = Assert.Single(session.ListRevisions());
        Assert.Equal(RevisionResolutionStatus.Malformed, entry.ResolutionStatus);
        Assert.Equal("unabsorbable_cell_deletion", entry.Diagnostic?.Code);
        Assert.Equal(EditErrorCode.RevisionMalformed, session.AcceptAllRevisions().Error?.Code);
    }

    [Fact]
    public void DuplicateId_RefusesAddressingOneRevisionButNotBulkResolution()
    {
        var bytes = Build(
            Paragraph(Run("Alpha "), Ins("w:id=\"7\" " + Stamp, "first "), Run("end"))
            + Paragraph(Run("Beta "), Ins("w:id=\"7\" w:author=\"Reviewer B\" w:date=\"2026-03-03T10:00:00Z\"", "second "),
                Run("end")));

        using (var session = new DocxSession(bytes))
        {
            var listed = session.ListRevisions();
            Assert.Equal(2, listed.Count);
            Assert.All(listed, entry =>
            {
                Assert.Equal(RevisionResolutionStatus.Ambiguous, entry.ResolutionStatus);
                Assert.Equal("duplicate_revision_id", entry.Diagnostic?.Code);
            });
            // The shared w:id is what a caller (or a comment, or a legacy rev id) would address
            // one of them by, so selective resolution still refuses and offers the repair.
            Assert.Equal(EditErrorCode.RevisionAmbiguous, session.AcceptRevision(listed[0].Id).Error?.Code);
            Assert.Equal(2, session.ListRevisionRepairs().Count);
        }

        // Resolving every entry needs no address: each is resolved by its own markup.
        AssertBulkMatchesProcessor(bytes,
            accepted: new[] { "Alpha first end", "Beta second end" },
            rejected: new[] { "Alpha end", "Beta end" });
    }

    [Fact]
    public void DuplicateId_BesideARefusedEntry_StillRefusesBulkResolution()
    {
        // Control: relaxing the addressing defect must not let an unsafe neighbour through.
        var bytes = Build(
            Paragraph(Run("One "), Ins("w:id=\"7\" " + Stamp, "first"), Run(" end"))
            + Paragraph(Run("Two "), Ins("w:id=\"7\" " + Stamp, "second"), Run(" end"))
            + Paragraph("<w:r><w:delText>orphan</w:delText></w:r>"));

        using var session = new DocxSession(bytes);

        var result = session.AcceptAllRevisions();
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.RevisionUnsupported, result.Error?.Code);
        Assert.Equal(3, session.ListRevisions().Count);
    }

    private const string Utc = "xmlns:w16du=\"http://schemas.microsoft.com/office/word/2023/wordml/word16du\" "
        + "w16du:dateUtc=\"2026-03-02T09:00:00Z\"";

    private static string Move(string sourceMarks, string rangeMarks, string destinationMarks) =>
        Paragraph(Run("Keep "), $"<w:moveFromRangeStart w:id=\"1\" w:name=\"m1\" {rangeMarks}/>",
            $"<w:moveFrom w:id=\"2\" {sourceMarks}><w:r><w:delText xml:space=\"preserve\">moved </w:delText></w:r></w:moveFrom>",
            "<w:moveFromRangeEnd w:id=\"1\"/>", Run("here"))
        + Paragraph(Run("Then "), $"<w:moveToRangeStart w:id=\"3\" w:name=\"m1\" {rangeMarks}/>",
            $"<w:moveTo w:id=\"4\" {destinationMarks}>{Run("moved ")}</w:moveTo>",
            "<w:moveToRangeEnd w:id=\"3\"/>", Run("there"));

    [Fact]
    public void Move_RangeStartsWithoutUtcDate_AreCoherentWithUtcStampedContent()
    {
        // Files reported to be saved by Word stamp w16du:dateUtc on content marks, not on range starts.
        var bytes = Build(Move(Stamp + " " + Utc, Stamp, Stamp + " " + Utc));

        using var session = new DocxSession(bytes);

        var move = Assert.Single(session.ListRevisions());
        Assert.Equal("move", move.Type);
        Assert.Equal(RevisionResolutionStatus.Supported, move.ResolutionStatus);
        Assert.Null(move.Diagnostic);
    }

    [Fact]
    public void MoveName_JoiningTwoReviewersMoves_RefusesSelectiveResolutionButNotBulk()
    {
        // Two reviewers' moves under one w:name: the registry sees one move, so accepting it by id
        // would also resolve the other reviewer's move. That is an addressing defect only.
        const string other = "w:author=\"Reviewer B\" w:date=\"2026-03-04T10:00:00Z\"";
        static string Source(int range, int id, string stamp, string text) =>
            $"<w:moveFromRangeStart w:id=\"{range}\" w:name=\"m1\" {stamp}/>"
            + $"<w:moveFrom w:id=\"{id}\" {stamp}><w:r><w:delText xml:space=\"preserve\">{text}</w:delText></w:r></w:moveFrom>"
            + $"<w:moveFromRangeEnd w:id=\"{range}\"/>";
        static string Destination(int range, int id, string stamp, string text) =>
            $"<w:moveToRangeStart w:id=\"{range}\" w:name=\"m1\" {stamp}/>"
            + $"<w:moveTo w:id=\"{id}\" {stamp}>{Run(text)}</w:moveTo><w:moveToRangeEnd w:id=\"{range}\"/>";
        var bytes = Build(
            Paragraph(Run("Start "), Source(1, 2, Stamp, "alpha "), Run("middle "), Source(5, 6, other, "beta "), Run("end"))
            + Paragraph(Run("Then "), Destination(3, 4, Stamp, "alpha "), Run("and "), Destination(7, 8, other, "beta "),
                Run("done")));

        using (var session = new DocxSession(bytes))
        {
            var move = Assert.Single(session.ListRevisions());
            Assert.Equal("move", move.Type);
            Assert.Equal(RevisionResolutionStatus.Ambiguous, move.ResolutionStatus);
            Assert.Equal("incoherent_move_metadata", move.Diagnostic?.Code);
            Assert.Equal(EditErrorCode.RevisionAmbiguous, session.AcceptRevision(move.Id).Error?.Code);
            Assert.Equal(EditErrorCode.RevisionAmbiguous, session.RejectRevision(move.Id).Error?.Code);
            Assert.Single(session.ListRevisions());
        }

        AssertBulkMatchesProcessor(bytes,
            accepted: new[] { "Start middle end", "Then alpha and beta done" },
            rejected: new[] { "Start alpha middle beta end", "Then and done" });
    }

    // ─── Helpers ─────────────────────────────────────────────────────────

    private static void AssertBulkMatchesProcessor(byte[] bytes, string[] accepted, string[] rejected)
    {
        var document = new WmlDocument("bulk-safety.docx", bytes);
        foreach (bool accept in new[] { true, false })
        {
            var expected = accept ? accepted : rejected;
            var processor = accept ? RevisionProcessor.AcceptRevisions(document) : RevisionProcessor.RejectRevisions(document);
            Assert.Equal(expected, ParagraphTexts(processor.DocumentByteArray));

            using var session = new DocxSession(bytes);
            var result = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
            Assert.True(result.Success, result.Error?.Message);
            var saved = session.Save();
            Assert.Equal(expected, ParagraphTexts(saved));
            Assert.Empty(session.ListRevisions());
            Assert.Empty(RevisionMarkers(saved));
        }
    }

    private static List<string> ParagraphTexts(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.GetXDocument().Descendants(W.p)
            .Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value)))
            .Where(text => text.Length > 0).ToList();
    }

    private static List<string> RevisionMarkers(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.GetXDocument().Descendants()
            .Where(Docxodus.Internal.RevisionOps.IsRecognizedRevisionMarker)
            .Select(element => element.Name.LocalName).ToList();
    }

    private static byte[] Build(string bodyXml)
    {
        var xml = $"<w:document xmlns:w=\"{W.w.NamespaceName}\"><w:body>{bodyXml}"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.PutXDocument(XDocument.Parse(xml));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
        }
        return stream.ToArray();
    }
}
