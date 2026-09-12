#nullable enable

// Copyright (c) John Scrudato IV. All rights reserved.
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
        var controlAnchor = Assert.Single(projection.AnchorIndex.Values,
            target => target.Anchor.Kind == "sdt").Anchor.Id;

        var result = session.DeleteRange(from, to);

        Assert.True(result.Success, result.Error?.Message);
        AssertAnchorAccounting(result, new[] { from, controlAnchor, controlled }, Array.Empty<string>());

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
        var controls = projection.AnchorIndex.Values
            .Where(target => target.Anchor.Kind == "sdt")
            .Select(target => target.Anchor.Id);

        var result = session.DeleteRange(from, to);

        Assert.True(result.Success, result.Error?.Message);
        AssertAnchorAccounting(
            result,
            new[] { from, outerParagraph, innerParagraph }.Concat(controls),
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
            .Append(Assert.Single(projection.AnchorIndex.Values,
                target => target.Anchor.Kind == "sdt").Anchor.Id)
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

    // Issue #764: a block w:customXml wrapper gets the same reversible envelope as a block
    // w:sdt. Its w:customXmlPr stays put (the schema orders it ahead of the range markup),
    // accept removes wrapper and payload, reject restores wrapper, attributes, properties
    // and text — through the session registry, individually and in bulk.
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS479_CustomXmlBlock_TracksWrapperAndRoundTripsThroughTheRegistry(bool accept)
    {
        byte[] tracked;
        string customParagraph;
        using (var session = OpenTrackedSession(BuildDocument(
            ParagraphWithText("before"),
            ParagraphWithText("delete start"),
            CustomXmlBlock("clause", ParagraphWithText("custom payload")),
            ParagraphWithText("after"))))
        {
            var projection = session.Project();
            var from = FindByText(session, projection, "delete start");
            customParagraph = FindByText(session, projection, "custom payload");
            var to = FindByText(session, projection, "after");

            var result = session.DeleteRange(from, to);

            Assert.True(result.Success, result.Error?.Message);
            AssertAnchorAccounting(result, new[] { from, customParagraph }, Array.Empty<string>());
            tracked = session.Save();
        }

        var body = Body(tracked);
        var wrapper = Assert.Single(body.Elements(W.customXml));
        Assert.Equal(W.customXmlPr, wrapper.Elements().First().Name);
        Assert.Equal(W.customXmlDelRangeEnd, wrapper.Elements().Skip(1).First().Name);
        Assert.Equal(W.customXmlDelRangeStart, wrapper.Elements().Last().Name);
        Assert.Equal(W.customXmlDelRangeStart, wrapper.ElementsBeforeSelf().Last().Name);
        Assert.Equal(W.customXmlDelRangeEnd, wrapper.ElementsAfterSelf().First().Name);
        Assert.NotNull(wrapper.Descendants(W.p).Single().Element(W.pPr)?.Element(W.rPr)?.Element(W.del));
        AssertSchemaValid(tracked);

        using var review = new DocxSession(tracked);
        var envelope = Assert.Single(review.ListRevisions(), revision =>
            revision.Family == RevisionFamily.ContentControlDelete);
        Assert.Equal(RevisionResolutionStatus.Supported, envelope.ResolutionStatus);
        // Anchor ids are session-minted, so the payload paragraph is identified by kind and
        // by the entry's text; a wrapper with no anchor of its own must not surface as an
        // empty anchor.
        Assert.Single(envelope.AffectedAnchors, anchor => anchor.Kind == "p");
        Assert.All(envelope.AffectedAnchors, anchor => Assert.NotEmpty(anchor.Id));
        Assert.Contains("custom payload", envelope.Text, StringComparison.Ordinal);

        var resolved = accept
            ? review.AcceptRevision(envelope.Id)
            : review.RejectRevision(envelope.Id);
        Assert.True(resolved.Success, resolved.Error?.Message);
        var reviewed = Body(review.Save());
        if (accept)
        {
            Assert.Empty(reviewed.Descendants(W.customXml));
            Assert.DoesNotContain("custom payload", reviewed.Value);
        }
        else
        {
            var restored = Assert.Single(reviewed.Elements(W.customXml));
            Assert.Equal("clause", (string?)restored.Attribute(W.element));
            Assert.Equal(W.customXmlPr, restored.Elements().First().Name);
            Assert.Equal("custom payload", restored.Value);
        }
        Assert.DoesNotContain(reviewed.Descendants(), element =>
            element.Name == W.customXmlDelRangeStart || element.Name == W.customXmlDelRangeEnd);
        AssertSchemaValid(review.Save());

        using var bulk = new DocxSession(tracked);
        Assert.True((accept ? bulk.AcceptAllRevisions() : bulk.RejectAllRevisions()).Success);
        Assert.Empty(bulk.ListRevisions());
        Assert.Equal(accept ? new[] { "before", "after" }
            : new[] { "before", "delete start", "custom payload", "after" },
            Body(bulk.Save()).Descendants(W.p).Select(p => p.Value).ToArray());
    }

    // Mixed nesting: a custom-XML wrapper holding a content control holding a table, and a
    // content control holding a custom-XML wrapper. Every wrapper gets its own envelope, a
    // user bookmark inside survives, and both directions round-trip through the registry.
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS480_MixedCustomXmlAndControlNesting_RoundTrips(bool accept)
    {
        var bookmarked = ParagraphWithText("inner paragraph");
        bookmarked.PrependChild(new BookmarkStart { Id = "5", Name = "keep" });
        bookmarked.AppendChild(new BookmarkEnd { Id = "5" });
        byte[] tracked;
        using (var session = OpenTrackedSession(BuildDocument(
            ParagraphWithText("before"),
            ParagraphWithText("delete start"),
            CustomXmlBlock("outer", BlockControl("table-control", TwoCellTable())),
            BlockControl("control", CustomXmlBlock("inner", bookmarked)),
            ParagraphWithText("after"))))
        {
            var projection = session.Project();
            var from = FindByText(session, projection, "delete start");
            var to = FindByText(session, projection, "after");

            var result = session.DeleteRange(from, to);

            Assert.True(result.Success, result.Error?.Message);
            Assert.Empty(result.Removed);
            tracked = session.Save();
        }

        var body = Body(tracked);
        Assert.Equal(2, body.Descendants(W.customXml).Count());
        Assert.Equal(2, body.Descendants(W.sdt).Count());
        Assert.Equal(8, body.Descendants(W.customXmlDelRangeStart).Count());
        Assert.Equal(8, body.Descendants(W.customXmlDelRangeEnd).Count());
        Assert.Single(body.Descendants(W.trPr).Elements(W.del));
        AssertSchemaValid(tracked);

        using var review = new DocxSession(tracked);
        var envelopes = review.ListRevisions()
            .Where(r => r.Family == RevisionFamily.ContentControlDelete).ToList();
        Assert.Equal(2, envelopes.Count);
        Assert.All(envelopes, r => Assert.Equal(RevisionResolutionStatus.Supported, r.ResolutionStatus));

        Assert.True((accept ? review.AcceptAllRevisions() : review.RejectAllRevisions()).Success);
        Assert.Empty(review.ListRevisions());
        var reviewed = Body(review.Save());
        Assert.DoesNotContain(reviewed.Descendants(), element =>
            element.Name.LocalName.StartsWith("customXmlDel", StringComparison.Ordinal));
        if (accept)
        {
            Assert.Empty(reviewed.Descendants(W.customXml));
            Assert.Empty(reviewed.Descendants(W.sdt));
            Assert.Empty(reviewed.Descendants(W.tbl));
            Assert.Equal(new[] { "before", "after" },
                reviewed.Descendants(W.p).Select(p => p.Value).ToArray());
        }
        else
        {
            Assert.Equal(2, reviewed.Descendants(W.customXml).Count());
            Assert.Equal(2, reviewed.Descendants(W.sdt).Count());
            Assert.Single(reviewed.Descendants(W.tbl));
            Assert.Single(reviewed.Descendants(W.bookmarkStart));
            Assert.Contains("inner paragraph", reviewed.Value);
            Assert.Contains("Cell A", reviewed.Value);
        }
        AssertSchemaValid(review.Save());
    }

    // Run-level custom XML inside a paragraph is the one shape the paragraph deleter cannot
    // represent (it marks direct-child runs only), so the pre-mutation refusal survives for it.
    [Fact]
    public void DS481_InlineCustomXml_FailsBeforeMutationWithStructuredError()
    {
        var inline = new CustomXmlRun(new CustomXmlProperties())
        {
            Uri = "urn:docxodus:test",
            Element = "inline",
        };
        inline.Append(new Run(new Text("inline payload")));
        using var session = OpenTrackedSession(BuildDocument(
            ParagraphWithText("before"),
            ParagraphWithText("delete start"),
            new Paragraph(new Run(new Text("host ")), inline),
            ParagraphWithText("after")));
        var projection = session.Project();
        var from = FindByText(session, projection, "delete start");
        var to = FindByText(session, projection, "after");

        var before = session.Save();
        var result = session.DeleteRange(from, to);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.IncompatibleElementType, result.Error?.Code);
        Assert.Contains("run-level w:customXml", result.Error?.Message);
        Assert.Equal(0, session.UndoCount);
        Assert.True(XNode.DeepEquals(Body(before), Body(session.Save())));
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
        var control = Assert.Single(projection.AnchorIndex.Values,
            target => target.Anchor.Kind == "sdt").Anchor.Id;

        var result = session.DeleteSection(heading);

        Assert.True(result.Success, result.Error?.Message);
        AssertAnchorAccounting(result, new[] { heading, control, controlled }, new[] { section });
        var tracked = session.Save();
        Assert.Single(Body(tracked).Elements(W.sdt));
        Assert.Empty(Body(tracked).Elements(W.sectPr));
        AssertSchemaValid(tracked);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS478_SessionRegistry_ResolvesAuthoredControlDeletionAtomically(bool accept)
    {
        byte[] tracked;
        using (var authoring = OpenTrackedSession(BuildDocument(
            ParagraphWithText("before"),
            ParagraphWithText("delete start"),
            BlockControl("controlled", ParagraphWithText("controlled paragraph")),
            ParagraphWithText("after"))))
        {
            var projection = authoring.Project();
            var from = FindByText(authoring, projection, "delete start");
            var to = FindByText(authoring, projection, "after");
            Assert.True(authoring.DeleteRange(from, to).Success);
            tracked = authoring.Save();
        }

        using var review = new DocxSession(tracked);
        var structured = Assert.Single(review.ListRevisions(), revision =>
            revision.Family == RevisionFamily.ContentControlDelete);
        Assert.Equal(RevisionResolutionStatus.Supported, structured.ResolutionStatus);
        Assert.Contains(structured.AffectedAnchors, anchor => anchor.Kind == "p");

        var result = accept
            ? review.AcceptRevision(structured.Id)
            : review.RejectRevision(structured.Id);

        Assert.True(result.Success, result.Error?.Message);
        var body = Body(review.Save());
        if (accept)
        {
            Assert.Empty(body.Elements(W.sdt));
            Assert.DoesNotContain("controlled paragraph", body.Value);
        }
        else
        {
            Assert.Single(body.Elements(W.sdt));
            Assert.Contains("controlled paragraph", body.Value);
        }
        Assert.DoesNotContain(body.Descendants(), element =>
            element.Name == W.customXmlDelRangeStart || element.Name == W.customXmlDelRangeEnd);
        AssertSchemaValid(review.Save());
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
            .Single(target => target.Anchor.Kind is "p" or "h" or "li"
                && session.GetAnchorInfo(target.Anchor.Id)?.TextPreview == text)
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
