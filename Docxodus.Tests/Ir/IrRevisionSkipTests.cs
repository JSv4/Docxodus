#nullable enable

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Guards the perf optimization in <see cref="IrReader"/>'s <c>ApplyRevisionView</c>: a revision-free
/// document skips the <see cref="RevisionProcessor"/> round-trip, but the skip's "is there any revision
/// markup?" scan MUST be sound — it must cover every element RevisionProcessor transforms (complete
/// element set) and every part the reader consumes (main + headers/footers/footnotes/endnotes/comments).
/// A scan that under-reports markup would silently bypass Accept/Reject and corrupt the output.
///
/// The behavioral tests assert that reading a document carrying ONLY a corner-case revision (one that an
/// earlier, main-part-only / incomplete-set scan would have missed) produces the same IR as reading
/// bytes already accepted by <see cref="RevisionProcessor.AcceptRevisions"/> — i.e. the skip did NOT
/// fire / the round-trip ran, so the result is identical.
/// </summary>
public class IrRevisionSkipTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Read the doc through the (skip-eligible) Accept path and project to markdown.</summary>
    private static string AcceptReadMarkdown(WmlDocument doc)
    {
        var ir = IrReader.Read(doc, new IrReaderOptions { RevisionView = RevisionView.Accept });
        return IrMarkdownEmitter.Emit(ir, new WmlToMarkdownConverterSettings()).Markdown;
    }

    /// <summary>
    /// Read through the (skip-eligible) Accept path and serialize the FULL IR as diagnostic JSON.
    /// Unlike the markdown projection, this captures every node hash — including opaque canonical
    /// hashes, which is where un-stripped revision markup shows up when the projection happens not to
    /// render the affected content (e.g. tblPrEx property exceptions).
    /// </summary>
    private static string AcceptReadJson(WmlDocument doc)
    {
        var ir = IrReader.Read(doc, new IrReaderOptions { RevisionView = RevisionView.Accept });
        return IrDiagnosticJson.Write(ir);
    }

    private static string PreAcceptedReadJson(WmlDocument doc) =>
        AcceptReadJson(RevisionProcessor.AcceptRevisions(doc));

    /// <summary>
    /// Pre-accept the revisions with the full RevisionProcessor round-trip, then Accept-read (now a
    /// genuine no-op, so the skip firing here is correct). This is the oracle the skip-eligible read
    /// must match byte-for-byte.
    /// </summary>
    private static string PreAcceptedReadMarkdown(WmlDocument doc) =>
        AcceptReadMarkdown(RevisionProcessor.AcceptRevisions(doc));

    private static string Table(string text) =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
        "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>" + text +
        "</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";

    [Fact]
    public void AcceptRead_TblPrExChangeOnly_MatchesPreAccepted()
    {
        // The ONLY revision markup is a w:tblPrExChange (table-level property-exception change). It can
        // appear with no other scanned revision element present (RevisionProcessor.cs:320-325 / :1725),
        // so a scan set lacking tblPrExChange would have wrongly skipped the round-trip.
        const string body =
            "<w:tbl>" +
            "<w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>" +
            "<w:tr><w:tc>" +
            "<w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/></w:tcPr>" +
            "<w:tblPrEx>" +
            "<w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/></w:tblBorders>" +
            "<w:tblPrExChange w:id=\"1\" w:author=\"a\" w:date=\"2017-01-01T00:00:00Z\">" +
            "<w:tblPrEx>" +
            "<w:tblBorders><w:top w:val=\"double\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/></w:tblBorders>" +
            "</w:tblPrEx>" +
            "</w:tblPrExChange>" +
            "</w:tblPrEx>" +
            "<w:p><w:r><w:t>cell</w:t></w:r></w:p>" +
            "</w:tc></w:tr></w:tbl>" +
            "<w:p><w:r><w:t>after</w:t></w:r></w:p>";
        var doc = IrTestDocuments.FromBodyXml(body);

        // Diagnostic JSON, not markdown: the projection never renders tblPrEx styling, so a markdown
        // comparison passes vacuously even when the skip wrongly fires and the IR reads the
        // tblPrExChange markup un-stripped. The JSON's opaque canonical hashes DO differ in that case
        // (the tainted tblPrEx hashes differently), making this assertion fail on an unsound scan set.
        Assert.Equal(PreAcceptedReadJson(doc), AcceptReadJson(doc));
        Assert.Equal(PreAcceptedReadMarkdown(doc), AcceptReadMarkdown(doc));
    }

    [Fact]
    public void AcceptRead_HeaderOnlyInsertion_MatchesPreAccepted()
    {
        // The body carries NO revision markup; the only w:ins lives in a header part. A scan that only
        // walked MainDocumentPart would skip the round-trip and leave the header insertion un-accepted.
        const string body = "<w:p><w:r><w:t>body text</w:t></w:r></w:p>";
        const string header =
            "<w:p><w:r><w:t xml:space=\"preserve\">kept </w:t></w:r>" +
            "<w:ins w:id=\"1\" w:author=\"a\" w:date=\"2017-01-01T00:00:00Z\">" +
            "<w:r><w:t>inserted-in-header</w:t></w:r></w:ins></w:p>";
        var doc = IrTestDocuments.FromBodyAndHeaderXml(body, header);

        // Sanity: the header insertion actually reaches the IR (headers are an in-scope read target),
        // so the comparison below is meaningful rather than vacuously equal on empty header output.
        var ir = IrReader.Read(doc, new IrReaderOptions { RevisionView = RevisionView.Accept });
        var headerText = string.Concat(
            ir.Headers.SelectMany(h => h.Scope.Blocks).OfType<IrParagraph>()
              .SelectMany(p => p.Inlines.OfType<IrTextRun>()).Select(t => t.Text));
        Assert.Contains("inserted-in-header", headerText);

        Assert.Equal(PreAcceptedReadMarkdown(doc), AcceptReadMarkdown(doc));
    }

    [Fact]
    public void AcceptRead_preserves_preexisting_clean_adjacent_table_boundaries()
    {
        // The insertion makes this document take the accept path.  The two following tables were
        // already adjacent and carry no revisions, so their original two-table topology is part of
        // the accepted VIEW.  RevisionProcessor's final-document normalizer normally merges them,
        // which made a clean counterpart look like a table delete+insert to the IR diff.
        const string dirtyLead =
            "<w:p><w:ins w:id=\"1\" w:author=\"A\" w:date=\"2020-01-01T00:00:00Z\">" +
            "<w:r><w:t>accepted lead</w:t></w:r></w:ins></w:p>";
        var doc = IrTestDocuments.FromBodyXml(dirtyLead + Table("first") + Table("second"));

        var accepted = IrReader.Read(doc,
            new IrReaderOptions { RevisionView = RevisionView.Accept });

        Assert.Equal(2, accepted.Body.Blocks.OfType<IrTable>().Count());
    }

    [Fact]
    public void AcceptRead_keeps_table_coalescing_when_acceptance_created_the_adjacency()
    {
        // Unlike the previous case, these tables are separated in the source.  Accepting the fully
        // deleted paragraph creates the adjacency, so it must retain RevisionProcessor's normal
        // coalescing behavior rather than being re-split as though it were an original table run.
        const string deletedParagraph =
            "<w:p><w:pPr><w:rPr><w:del w:id=\"2\" w:author=\"A\" w:date=\"2020-01-01T00:00:00Z\"/>" +
            "</w:rPr></w:pPr><w:del w:id=\"3\" w:author=\"A\" w:date=\"2020-01-01T00:00:00Z\">" +
            "<w:r><w:delText>deleted separator</w:delText></w:r></w:del></w:p>";
        var doc = IrTestDocuments.FromBodyXml(Table("first") + deletedParagraph + Table("second"));

        var accepted = IrReader.Read(doc,
            new IrReaderOptions { RevisionView = RevisionView.Accept });

        Assert.Single(accepted.Body.Blocks.OfType<IrTable>());
    }

    /// <summary>
    /// Set-drift guard. The skip scan's element set (<see cref="IrReader.ProcessorActsOnNamesForTest"/>)
    /// must list every element name <see cref="RevisionProcessor"/>'s transforms dispatch on by name.
    /// This hardcoded list is the independent witness: if RevisionProcessor.cs adds or removes a
    /// revision element, this test fails until BOTH this list AND ProcessorActsOnNameSet in IrReader.cs
    /// are updated together.
    ///
    /// Deliberately EXCLUDED (provably masked by a listed ancestor — see the masking analysis on
    /// IrReader.ProcessorActsOnNameSet): w:instrText and w:t, both only transformed inside a w:ins
    /// subtree (RevisionProcessor.cs:1223 / :1232, gated on rri.InInsert). w:ins is in the set, so the
    /// scan still catches those documents.
    /// </summary>
    [Fact]
    public void ProcessorActsOnNameSet_MatchesProcessor()
    {
        // Source of truth: the `element.Name == W.*` dispatch in RevisionProcessor.cs. When that set of
        // names changes, update this list AND IrReader.ProcessorActsOnNameSet in lock-step.
        var expected = new HashSet<XName>
        {
            // run / paragraph / move insert-delete markup
            W + "ins", W + "del", W + "moveFrom", W + "moveTo",
            W + "moveFromRangeStart", W + "moveFromRangeEnd",
            W + "moveToRangeStart", W + "moveToRangeEnd",
            // property-revision markers
            W + "rPrChange", W + "pPrChange", W + "sectPrChange",
            W + "tblPrChange", W + "tblGridChange", W + "trPrChange", W + "tcPrChange",
            W + "tblPrExChange", W + "numberingChange",
            // table cell revisions
            W + "cellIns", W + "cellDel", W + "cellMerge",
            // deleted text / field markers (transformed unconditionally by name in RevisionProcessor)
            W + "delText", W + "delInstrText",
            // custom-XML range markers (Start and End — End is removed unconditionally by Accept,
            // RevisionProcessor.cs:1693-1698, so producer-validity pairing must not be assumed)
            W + "customXmlInsRangeStart", W + "customXmlDelRangeStart",
            W + "customXmlMoveFromRangeStart", W + "customXmlMoveToRangeStart",
            W + "customXmlInsRangeEnd", W + "customXmlDelRangeEnd",
            W + "customXmlMoveFromRangeEnd", W + "customXmlMoveToRangeEnd",
        };

        var actual = new HashSet<XName>(IrReader.ProcessorActsOnNamesForTest);
        Assert.True(expected.SetEquals(actual),
            "Skip-scan element set drifted from RevisionProcessor's revision elements.\n" +
            "Missing from scan set: " + string.Join(", ", expected.Except(actual)) + "\n" +
            "Extra in scan set: " + string.Join(", ", actual.Except(expected)) + "\n" +
            "Update BOTH this list and IrReader.ProcessorActsOnNameSet (see RevisionProcessor.cs).");
    }
}
