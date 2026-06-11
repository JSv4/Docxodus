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
    /// Pre-accept the revisions with the full RevisionProcessor round-trip, then Accept-read (now a
    /// genuine no-op, so the skip firing here is correct). This is the oracle the skip-eligible read
    /// must match byte-for-byte.
    /// </summary>
    private static string PreAcceptedReadMarkdown(WmlDocument doc) =>
        AcceptReadMarkdown(RevisionProcessor.AcceptRevisions(doc));

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
            // custom-XML range-start markers
            W + "customXmlInsRangeStart", W + "customXmlDelRangeStart",
            W + "customXmlMoveFromRangeStart", W + "customXmlMoveToRangeStart",
        };

        var actual = new HashSet<XName>(IrReader.ProcessorActsOnNamesForTest);
        Assert.True(expected.SetEquals(actual),
            "Skip-scan element set drifted from RevisionProcessor's revision elements.\n" +
            "Missing from scan set: " + string.Join(", ", expected.Except(actual)) + "\n" +
            "Extra in scan set: " + string.Join(", ", actual.Except(expected)) + "\n" +
            "Update BOTH this list and IrReader.ProcessorActsOnNameSet (see RevisionProcessor.cs).");
    }
}
