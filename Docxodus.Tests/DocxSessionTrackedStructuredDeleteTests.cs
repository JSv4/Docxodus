#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using WLock = DocumentFormat.OpenXml.Wordprocessing.Lock;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace Docxodus.Tests;

public class DocxSessionTrackedStructuredDeleteTests
{
    [Fact]
    public void DS473_DeleteRange_TracksBlockContentControlInsteadOfHardRemovingIt()
    {
        using var session = OpenTrackedSession(BuildDocument(
            ParagraphWithText("before"),
            ParagraphWithText("delete start"),
            BlockControl("controlled", ParagraphWithText("controlled paragraph")),
            ParagraphWithText("after")));
        var projection = session.Project();
        var from = FindByText(session, projection, "delete start");
        var controlled = FindByText(session, projection, "controlled paragraph");
        var to = FindByText(session, projection, "after");

        var result = session.DeleteRange(from, to);

        Assert.True(result.Success, result.Error?.Message);
        AssertAnchorAccounting(result, new[] { from, controlled }, Array.Empty<string>());

        var tracked = session.Save();
        var body = Body(tracked);
        var control = Assert.Single(body.Elements(W.sdt));
        AssertEnvelopeRangeTopology(control, control.Element(W.sdtContent)!);
        AssertSchemaValid(tracked);
    }

    [Fact]
    public void DS474_NestedLockedDataBoundControls_TrackAndRoundTrip()
    {
        var outer = LockedBoundControl(
            "outer",
            ParagraphWithText("outer paragraph"),
            BlockControl("inner", ParagraphWithText("inner paragraph")));
        using var session = OpenTrackedSession(BuildDocument(
            ParagraphWithText("before"),
            ParagraphWithText("delete start"),
            outer,
            ParagraphWithText("after")));
        var projection = session.Project();
        var from = FindByText(session, projection, "delete start");
        var outerParagraph = FindByText(session, projection, "outer paragraph");
        var innerParagraph = FindByText(session, projection, "inner paragraph");
        var to = FindByText(session, projection, "after");

        var result = session.DeleteRange(from, to);

        Assert.True(result.Success, result.Error?.Message);
        AssertAnchorAccounting(
            result,
            new[] { from, outerParagraph, innerParagraph },
            Array.Empty<string>());

        var tracked = session.Save();
        var trackedBody = Body(tracked);
        var trackedOuter = Assert.Single(trackedBody.Elements(W.sdt));
        var trackedInner = Assert.Single(trackedOuter.Descendants(W.sdt));
        AssertEnvelopeRangeTopology(trackedOuter, trackedOuter.Element(W.sdtContent)!);
        AssertEnvelopeRangeTopology(trackedInner, trackedInner.Element(W.sdtContent)!);
        Assert.Equal("sdtLocked", (string?)trackedOuter.Element(W.sdtPr)?.Element(W._lock)?.Attribute(W.val));
        Assert.Equal("/root/value", (string?)trackedOuter.Element(W.sdtPr)?.Element(W.dataBinding)?.Attribute(W.xpath));
        AssertSchemaValid(tracked);

        var accepted = Resolve(tracked, accept: true);
        var acceptedBody = Body(accepted);
        Assert.Empty(acceptedBody.Descendants(W.sdt));
        Assert.DoesNotContain("outer paragraph", acceptedBody.Value);
        Assert.DoesNotContain("inner paragraph", acceptedBody.Value);
        AssertSchemaValid(accepted);

        var rejected = Resolve(tracked, accept: false);
        var rejectedBody = Body(rejected);
        Assert.Equal(2, rejectedBody.Descendants(W.sdt).Count());
        var rejectedOuter = Assert.Single(rejectedBody.Elements(W.sdt));
        Assert.Equal("sdtLocked", (string?)rejectedOuter.Element(W.sdtPr)?.Element(W._lock)?.Attribute(W.val));
        Assert.Equal("/root/value", (string?)rejectedOuter.Element(W.sdtPr)?.Element(W.dataBinding)?.Attribute(W.xpath));
        Assert.Contains("outer paragraph", rejectedBody.Value);
        Assert.Contains("inner paragraph", rejectedBody.Value);
        AssertSchemaValid(rejected);
    }

    [Fact]
    public void DS475_ControlContainingTable_TracksEveryDescendantAnchorAndRoundTrips()
    {
        using var session = OpenTrackedSession(BuildDocument(
            ParagraphWithText("before"),
            ParagraphWithText("delete start"),
            BlockControl("table-control", TwoCellTable()),
            ParagraphWithText("after")));
        var projection = session.Project();
        var from = FindByText(session, projection, "delete start");
        var to = FindByText(session, projection, "after");
        var table = projection.AnchorIndex.Values.Single(target => target.Anchor.Kind == "tbl");
        var tableXml = XElement.Parse(session.Raw.GetXml(table.Anchor.Id));
        var tableUnids = tableXml.DescendantsAndSelf()
            .Select(e => (string?)e.Attribute(PtOpenXml.Unid))
            .Where(id => id is not null)
            .ToHashSet(StringComparer.Ordinal);
        var expectedModified = projection.AnchorIndex.Values
            .Where(target => tableUnids.Contains(target.Unid))
            .Select(target => target.Anchor.Id)
            .Append(from)
            .Distinct(StringComparer.Ordinal)
            .ToList();

        var result = session.DeleteRange(from, to);

        Assert.True(result.Success, result.Error?.Message);
        AssertAnchorAccounting(result, expectedModified, Array.Empty<string>());

        var tracked = session.Save();
        var trackedBody = Body(tracked);
        var control = Assert.Single(trackedBody.Elements(W.sdt));
        AssertEnvelopeRangeTopology(control, control.Element(W.sdtContent)!);
        Assert.Single(control.Descendants(W.tr));
        Assert.Single(control.Descendants(W.trPr).Elements(W.del));
        Assert.Equal(2, control.Descendants(W.p)
            .Count(p => p.Element(W.pPr)?.Element(W.rPr)?.Element(W.del) is not null));
        AssertSchemaValid(tracked);

        var accepted = Resolve(tracked, accept: true);
        Assert.Empty(Body(accepted).Descendants(W.sdt));
        Assert.Empty(Body(accepted).Descendants(W.tbl));
        Assert.DoesNotContain("Cell A", Body(accepted).Value);
        Assert.DoesNotContain("Cell B", Body(accepted).Value);
        AssertSchemaValid(accepted);

        var rejected = Resolve(tracked, accept: false);
        Assert.Single(Body(rejected).Descendants(W.sdt));
        Assert.Single(Body(rejected).Descendants(W.tbl));
        Assert.Contains("Cell A", Body(rejected).Value);
        Assert.Contains("Cell B", Body(rejected).Value);
        AssertSchemaValid(rejected);
    }

    [Fact]
    public void DS476_CustomXmlBlock_FailsBeforeMutationWithStructuredError()
    {
        using var session = OpenTrackedSession(BuildDocument(
            ParagraphWithText("before"),
            ParagraphWithText("delete start"),
            CustomXmlBlock("clause", ParagraphWithText("custom payload")),
            ParagraphWithText("after")));
        var projection = session.Project();
        var from = FindByText(session, projection, "delete start");
        var customParagraph = FindByText(session, projection, "custom payload");
        var to = FindByText(session, projection, "after");

        var before = session.Save();
        var result = session.DeleteRange(from, to);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.IncompatibleElementType, result.Error?.Code);
        Assert.Contains("w:customXml", result.Error?.Message);
        AssertAnchorAccounting(result, Array.Empty<string>(), Array.Empty<string>());
        Assert.Equal(0, session.UndoCount);

        var after = session.Save();
        Assert.True(XNode.DeepEquals(Body(before), Body(after)));
        var preserved = Assert.Single(Body(after).Elements(W.customXml));
        Assert.Equal("clause", (string?)preserved.Attribute(W.element));
        Assert.Contains("custom payload", preserved.Value);
        Assert.Equal("custom payload", session.GetAnchorInfo(customParagraph)?.TextPreview);
        AssertSchemaValid(after);
    }

    [Fact]
    public void DS477_DeleteSection_TracksControlAndReportsSectionPropertyFallThrough()
    {
        using var session = OpenTrackedSession(BuildDocument(
            Heading("Delete section"),
            BlockControl("section-control", ParagraphWithText("controlled section payload")),
            new SectionProperties(new PageSize { Width = 12240, Height = 15840 })));
        var projection = session.Project();
        var heading = FindByText(session, projection, "Delete section");
        var controlled = FindByText(session, projection, "controlled section payload");
        var section = projection.AnchorIndex.Values.Single(target => target.Anchor.Kind == "sec").Anchor.Id;

        var result = session.DeleteSection(heading);

        Assert.True(result.Success, result.Error?.Message);
        AssertAnchorAccounting(result, new[] { heading, controlled }, new[] { section });
        var tracked = session.Save();
        Assert.Single(Body(tracked).Elements(W.sdt));
        Assert.Empty(Body(tracked).Elements(W.sectPr));
        AssertSchemaValid(tracked);
    }

    private static DocxSession OpenTrackedSession(byte[] bytes) =>
        new(bytes, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "issue-473",
        });

    private static string FindByText(
        DocxSession session,
        MarkdownProjection projection,
        string text) =>
        projection.AnchorIndex.Values
            .Single(target => session.GetAnchorInfo(target.Anchor.Id)?.TextPreview == text)
            .Anchor.Id;

    private static Paragraph ParagraphWithText(string text) =>
        new(new Run(new Text(text)));

    private static SdtBlock BlockControl(string tag, params OpenXmlElement[] content) =>
        new(
            new SdtProperties(new Tag { Val = tag }),
            new SdtContentBlock(content));

    private static SdtBlock LockedBoundControl(string tag, params OpenXmlElement[] content) =>
        new(
            new SdtProperties(
                new Tag { Val = tag },
                new WLock { Val = LockingValues.SdtLocked },
                new DataBinding
                {
                    StoreItemId = "{11111111-1111-1111-1111-111111111111}",
                    XPath = "/root/value",
                    PrefixMappings = "xmlns:x='urn:docxodus:test'",
                }),
            new SdtContentBlock(content));

    private static CustomXmlBlock CustomXmlBlock(string element, params OpenXmlElement[] content)
    {
        var customXml = new CustomXmlBlock(new CustomXmlProperties())
        {
            Uri = "urn:docxodus:test",
            Element = element,
        };
        customXml.Append(content);
        return customXml;
    }

    private static Paragraph Heading(string text) =>
        new(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text(text)));

    private static WTable TwoCellTable() =>
        new(
            new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Dxa }),
            new TableGrid(
                new GridColumn { Width = "2500" },
                new GridColumn { Width = "2500" }),
            new WTableRow(
                TableCell("Cell A"),
                TableCell("Cell B")));

    private static WTableCell TableCell(string text) =>
        new(
            new TableCellProperties(
                new TableCellWidth { Width = "2500", Type = TableWidthUnitValues.Dxa }),
            ParagraphWithText(text));

    private static byte[] BuildDocument(params OpenXmlElement[] blocks)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream,
                   WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(blocks));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new DocDefaults(),
                new Style(new StyleName { Val = "heading 1" })
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Heading1",
                });
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            document.Save();
        }

        return stream.ToArray();
    }

    private static byte[] Resolve(byte[] tracked, bool accept)
    {
        var document = new WmlDocument("tracked.docx", tracked);
        return (accept
            ? RevisionProcessor.AcceptRevisions(document)
            : RevisionProcessor.RejectRevisions(document)).DocumentByteArray;
    }

    private static XElement Body(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return new XElement(document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!);
    }

    private static void AssertEnvelopeRangeTopology(
        XElement wrapper,
        XElement contentContainer)
    {
        var parent = Assert.IsType<XElement>(wrapper.Parent);
        var siblings = parent.Elements().ToList();
        var wrapperIndex = siblings.IndexOf(wrapper);
        Assert.InRange(wrapperIndex, 1, siblings.Count - 2);

        var before = siblings[wrapperIndex - 1];
        var after = siblings[wrapperIndex + 1];
        Assert.Equal(W.customXmlDelRangeStart, before.Name);
        Assert.Equal(W.customXmlDelRangeEnd, after.Name);
        Assert.Equal("issue-473", (string?)before.Attribute(W.author));
        Assert.NotNull(before.Attribute(W.date));

        var payload = contentContainer.Elements().ToList();
        var openingEnd = payload[0];
        var closingStart = payload[^1];
        Assert.Equal(W.customXmlDelRangeEnd, openingEnd.Name);
        Assert.Equal(W.customXmlDelRangeStart, closingStart.Name);
        Assert.Equal((string?)before.Attribute(W.id), (string?)openingEnd.Attribute(W.id));
        Assert.Equal((string?)closingStart.Attribute(W.id), (string?)after.Attribute(W.id));
        Assert.NotEqual((string?)before.Attribute(W.id), (string?)closingStart.Attribute(W.id));
    }

    private static void AssertAnchorAccounting(
        EditResult result,
        IEnumerable<string> modified,
        IEnumerable<string> removed)
    {
        var expectedModified = modified.ToHashSet(StringComparer.Ordinal);
        var expectedRemoved = removed.ToHashSet(StringComparer.Ordinal);
        var actualModified = result.Modified.Select(anchor => anchor.Id).ToHashSet(StringComparer.Ordinal);
        var actualRemoved = result.Removed.Select(anchor => anchor.Id).ToHashSet(StringComparer.Ordinal);

        Assert.Equal(expectedModified.OrderBy(id => id), actualModified.OrderBy(id => id));
        Assert.Equal(expectedRemoved.OrderBy(id => id), actualRemoved.OrderBy(id => id));
        Assert.Empty(actualModified.Intersect(actualRemoved));
        Assert.Equal(actualModified.Count, result.Modified.Count);
        Assert.Equal(actualRemoved.Count, result.Removed.Count);
    }

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(
            errors.Count == 0,
            "Unexpected schema errors:\n" + string.Join("\n", errors.Select(error => error.Description)));
    }
}
