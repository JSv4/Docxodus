// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests;

public sealed class DocxSessionContentControlTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

    [Fact]
    public void CC001_Registry_IsOuterBeforeInner_AndReportsNativeMetadataPlacementAndFailures()
    {
        using var session = new DocxSession(BuildFixture());
        var controls = session.ListContentControls();
        Assert.Equal(new[] { "100", "101", "102", "103", "104", "105", "106",
            "107", "107", null, "108", "109", "110", "111", "112" },
            controls.Select(control => control.NativeId).ToArray());

        var outer = controls[0];
        var inner = controls[1];
        Assert.Equal(ContentControlType.RichText, outer.Type);
        Assert.Equal(ContentControlPlacement.Block, outer.Placement);
        Assert.Equal("outer-tag", outer.Tag);
        Assert.Equal("Outer alias", outer.Alias);
        Assert.False(outer.CanMutate);
        Assert.Contains("nested controls", outer.UnsupportedReason, StringComparison.Ordinal);
        Assert.Equal(outer.AnchorId, inner.ParentAnchorId);
        Assert.Equal(1, inner.Depth);
        Assert.Equal(ContentControlPlacement.Inline, inner.Placement);
        Assert.Contains(inner.AnchorId,
            session.ListInlineSpans(ParagraphAnchors(session).First(anchor =>
                session.Project().AnchorIndex[anchor].TextPreview.Contains("inner", StringComparison.Ordinal)))
                .SelectMany(span => span.ContentControlAnchorIds));

        Assert.True(controls.Single(control => control.NativeId == "106").IsBound);
        Assert.True(controls.Single(control => control.NativeId == "106").CanDetachTargetBinding);
        Assert.All(controls.Where(control => control.NativeId == "107"), control =>
        {
            Assert.True(control.HasDuplicateNativeId);
            Assert.False(control.CanMutate);
        });
        Assert.False(controls.Single(control => control.NativeId is null).HasValidNativeId);
        Assert.Equal(ContentControlType.Unsupported,
            controls.Single(control => control.NativeId == "110").Type);
    }

    [Fact]
    public void CC002_TextFamilies_PreserveWrapperProperties_AndUndoRedo()
    {
        using var session = new DocxSession(BuildFixture());
        var controls = session.ListContentControls();
        string Id(string native) => controls.Single(control => control.NativeId == native).AnchorId;

        var nestedParent = session.FillContentControlRichText(Id("100"), "**replacement**");
        Assert.Equal(EditErrorCode.ContentControlNestedFillUnsupported, nestedParent.Error!.Code);
        int before = session.UndoCount;
        Assert.True(session.FillContentControlText(Id("101"), "new inner").Success);
        Assert.True(session.SetContentControlChecked(Id("102"), true).Success);
        Assert.True(session.SetContentControlDate(Id("103"),
            DateTimeOffset.Parse("2031-05-06T00:00:00Z"), "May 6, 2031").Success);
        Assert.True(session.SelectContentControlItem(Id("104"), "b").Success);
        Assert.True(session.SelectContentControlItem(Id("105"), "Alpha").Success);
        Assert.Equal(before + 5, session.UndoCount);

        var live = session.ListContentControls();
        Assert.Equal("new inner", live.Single(control => control.NativeId == "101").Text);
        Assert.Equal("☒", live.Single(control => control.NativeId == "102").Text);
        Assert.Equal("May 6, 2031", live.Single(control => control.NativeId == "103").Text);
        Assert.Equal("Beta", live.Single(control => control.NativeId == "104").Text);
        Assert.Equal("Alpha", live.Single(control => control.NativeId == "105").Text);
        Assert.Equal("outer-tag", live.Single(control => control.NativeId == "100").Tag);
        Assert.True(session.Undo());
        Assert.Equal("pick", session.ListContentControls().Single(control => control.NativeId == "105").Text);
        Assert.True(session.Redo());
        Assert.Equal("Alpha", session.ListContentControls().Single(control => control.NativeId == "105").Text);

        var saved = session.Save();
        using var reopened = new DocxSession(saved);
        Assert.Equal("☒", reopened.ListContentControls()
            .Single(control => control.NativeId == "102").Text);
        Assert.Equal("May 6, 2031", reopened.ListContentControls()
            .Single(control => control.NativeId == "103").Text);
        Assert.Equal("Beta", reopened.ListContentControls()
            .Single(control => control.NativeId == "104").Text);
        using var doc = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var validationErrors = new OpenXmlValidator(FileFormatVersions.Office2013).Validate(doc)
            .Where(IsMaterialValidationError).ToList();
        Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine,
            validationErrors.Select(validation =>
                $"{validation.Description} Node: {validation.Node?.OuterXml}")));
    }

    [Fact]
    public void CC003_BindingFailsClosed_DetachIsTargetOnly_AndCustomXmlBytesStayExact()
    {
        var fixture = BuildFixture();
        var customBefore = CustomXmlBytes(fixture);
        using var session = new DocxSession(fixture);
        var bound = session.ListContentControls().Single(control => control.NativeId == "106");
        var refused = session.FillContentControlText(bound.AnchorId, "bound replacement");
        Assert.Equal(EditErrorCode.ContentControlBound, refused.Error!.Code);
        Assert.Equal(0, session.UndoCount);

        var changed = session.FillContentControlText(bound.AnchorId, "detached replacement",
            new ContentControlFillOptions { BindingPolicy = ContentControlBindingPolicy.DetachTarget });
        Assert.True(changed.Success, changed.Error?.Message);
        var saved = session.Save();
        Assert.Equal(customBefore, CustomXmlBytes(saved));
        using var doc = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var control = doc.MainDocumentPart!.GetXDocument().Descendants(W + "sdt")
            .Single(value => (string?)value.Element(W + "sdtPr")?.Element(W + "id")?.Attribute(W + "val") == "106");
        Assert.Null(control.Element(W + "sdtPr")?.Element(W + "dataBinding"));
        Assert.Equal("detached replacement",
            string.Concat(control.Descendants(W + "t").Select(text => text.Value)));
    }

    [Fact]
    public void CC004_EffectiveLocksMalformedUnsupportedAndTrackedModeFailWithoutHistory()
    {
        using var session = new DocxSession(BuildFixture());
        var controls = session.ListContentControls();
        EditResult Fill(string native) => session.FillContentControlText(
            controls.First(control => control.NativeId == native).AnchorId, "x");
        Assert.Equal(EditErrorCode.ContentControlLocked, Fill("112").Error!.Code);
        Assert.Equal(EditErrorCode.ContentControlMalformed, Fill("107").Error!.Code);
        Assert.Equal(EditErrorCode.ContentControlUnsupported, Fill("110").Error!.Code);
        Assert.Equal(0, session.UndoCount);

        session.SetTrackedChanges(TrackedChangeMode.RenderInline);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, Fill("101").Error!.Code);
        var paragraph = ParagraphAnchors(session).First(anchor =>
            session.Project().AnchorIndex[anchor].TextPreview.Contains("inner", StringComparison.Ordinal));
        Assert.True(session.ReplaceTextAtSpan(paragraph, 7, 5, "INNER").Success);
    }

    [Fact]
    public void CC005_DefaultSaveReopen_RederivesSameSdtAnchorsFromNativeIds()
    {
        using var session = new DocxSession(BuildFixture());
        var original = session.ListContentControls()
            .Where(control => control.HasValidNativeId && !control.HasDuplicateNativeId)
            .ToDictionary(control => control.NativeId!, control => control.AnchorId);
        Assert.True(session.FillContentControlText(original["101"], "identity-independent value").Success);
        var saved = session.Save(false);
        Assert.DoesNotContain("Unid", Encoding.UTF8.GetString(saved), StringComparison.Ordinal);

        using var reopened = new DocxSession(saved);
        var after = reopened.ListContentControls()
            .Where(control => control.HasValidNativeId && !control.HasDuplicateNativeId)
            .ToDictionary(control => control.NativeId!, control => control.AnchorId);
        Assert.Equal(original, after);
        Assert.Equal("identity-independent value", reopened.GetContentControl(original["101"])!.Text);
    }

    [Fact]
    public void CC006_RepeatingSectionCloneFreshensNestedIds_AndIsUndoable()
    {
        using var session = new DocxSession(BuildFixture());
        var section = session.ListContentControls().Single(control => control.NativeId == "108");
        var add = session.AddRepeatingSectionItem(section.AnchorId);
        Assert.True(add.Success, add.Error?.Message);
        Assert.Single(add.Created);
        var controls = session.ListContentControls();
        var sections = controls.Where(control => control.Type == ContentControlType.RepeatingSectionItem).ToList();
        Assert.Equal(2, sections.Count);
        Assert.Equal(2, sections.Select(control => control.NativeId).Distinct().Count());
        Assert.All(sections, control => Assert.Equal(section.AnchorId, control.ParentAnchorId));
        var repeatedParagraphs = session.Project().AnchorIndex.Values.Where(value =>
            value.Anchor.Kind == "p" && value.TextPreview == "item").ToList();
        Assert.Equal(2, repeatedParagraphs.Count);
        Assert.Equal(2, repeatedParagraphs.Select(value => value.Anchor.Id).Distinct().Count());

        Assert.True(session.Undo());
        Assert.Single(session.ListContentControls().Where(control =>
            control.Type == ContentControlType.RepeatingSectionItem));
        Assert.True(session.Redo());
        var item = session.ListContentControls().Last(control =>
            control.Type == ContentControlType.RepeatingSectionItem);
        Assert.True(session.RemoveRepeatingSectionItem(item.AnchorId).Success);
        Assert.Single(session.ListContentControls().Where(control =>
            control.Type == ContentControlType.RepeatingSectionItem));
    }

    [Fact]
    public void CC007_OracleAndIrExposeSameSdtIndex_WithoutChangingMarkdownBytes()
    {
        var fixture = BuildFixture();
        var settings = new WmlToMarkdownConverterSettings();
        var oracle = WmlToMarkdownConverter.Convert(new WmlDocument("controls.docx", fixture), settings);
        var ir = IrMarkdownEmitter.Emit(IrReader.Read(new WmlDocument("controls.docx", fixture),
            new IrReaderOptions { RetainSources = false }), settings).ToProjection();
        Assert.Equal(oracle.Markdown, ir.Markdown);
        Assert.Equal(oracle.AnchorIndex.Keys, ir.AnchorIndex.Keys);
        Assert.Equal(15, oracle.AnchorIndex.Values.Select(value => value.Anchor.Id).Distinct()
            .Count(id => id.StartsWith("sdt:", StringComparison.Ordinal)));

        using var session = new DocxSession(fixture);
        Assert.DoesNotContain(session.ListBlocks().Body, unit => unit.Kind == "sdt");
        int handle = Docxodus.Internal.DocxSessionOps.OpenSession(fixture, null);
        try
        {
            var html = Docxodus.Internal.DocxSessionOps.RenderHtml(
                handle, "dx-", false, false, 1.0);
            Assert.Contains("inner", html);
            Assert.DoesNotContain("<w:sdt", html, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            Docxodus.Internal.DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void CC008_OpsJson_IsStrictAndRoundTripsRegistryAndMutations()
    {
        int handle = Docxodus.Internal.DocxSessionOps.OpenSession(BuildFixture(), null);
        try
        {
            using var listed = JsonDocument.Parse(
                Docxodus.Internal.DocxSessionOps.ListContentControls(handle));
            var controls = listed.RootElement.EnumerateArray().ToArray();
            Assert.Equal(15, controls.Length);
            var plain = controls.Single(control => control.TryGetProperty("nativeId", out var id)
                && id.GetString() == "101");
            Assert.Equal("plain_text", plain.GetProperty("type").GetString());
            Assert.Equal("inline", plain.GetProperty("placement").GetString());
            Assert.NotEmpty(plain.GetProperty("parentAnchorId").GetString()!);

            var anchor = plain.GetProperty("anchorId").GetString()!;
            using var filled = JsonDocument.Parse(
                Docxodus.Internal.DocxSessionOps.FillContentControlText(
                    handle, anchor, "transport value", "{}"));
            Assert.True(filled.RootElement.GetProperty("success").GetBoolean());

            using var invalidOptions = JsonDocument.Parse(
                Docxodus.Internal.DocxSessionOps.FillContentControlText(
                    handle, anchor, "ignored", "{\"unknown\":true}"));
            Assert.False(invalidOptions.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal("invalid_content_control_value", invalidOptions.RootElement
                .GetProperty("error").GetProperty("code").GetString());

            var dateAnchor = controls.Single(control => control.TryGetProperty("nativeId", out var id)
                    && id.GetString() == "103").GetProperty("anchorId").GetString()!;
            using var emptyDisplayDate = JsonDocument.Parse(
                Docxodus.Internal.DocxSessionOps.SetContentControlDate(
                    handle, dateAnchor, "2031-05-06T00:00:00Z", "", "{}"));
            Assert.True(emptyDisplayDate.RootElement.GetProperty("success").GetBoolean());
            using var relisted = JsonDocument.Parse(
                Docxodus.Internal.DocxSessionOps.ListContentControls(handle));
            Assert.Equal(string.Empty, relisted.RootElement.EnumerateArray().Single(control =>
                control.TryGetProperty("nativeId", out var id) && id.GetString() == "103")
                .GetProperty("text").GetString());

            using var invalidDate = JsonDocument.Parse(
                Docxodus.Internal.DocxSessionOps.SetContentControlDate(
                    handle, dateAnchor, "not-a-date", null, "{}"));
            Assert.Equal("invalid_content_control_value", invalidDate.RootElement
                .GetProperty("error").GetProperty("code").GetString());
        }
        finally
        {
            Docxodus.Internal.DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void CC009_PictureFill_ReusesNativeImageValidationAndRelationshipSeams()
    {
        using var session = new DocxSession(BuildPictureFixture());
        var picture = session.ListContentControls().Single(control => control.NativeId == "113");
        var before = Assert.Single(session.ListImages().Where(image => image.AnchorId is not null
            && image.AnchorId.StartsWith("p:body:", StringComparison.Ordinal)
            && image.IntrinsicWidthPixels == 2));
        Assert.True(session.FillContentControlPicture(picture.AnchorId, Png(7, 9)).Success);
        var after = Assert.Single(session.ListImages().Where(image => image.AnchorId == before.AnchorId));
        Assert.Equal(7, after.IntrinsicWidthPixels);
        Assert.Equal(9, after.IntrinsicHeightPixels);
        Assert.Equal("picture-tag", session.GetContentControl(picture.AnchorId)!.Tag);
        Assert.True(session.Undo());
        Assert.Equal(2, Assert.Single(session.ListImages().Where(image =>
            image.AnchorId == before.AnchorId)).IntrinsicWidthPixels);

        Assert.True(session.Redo());
        var saved = session.Save();
        using var reopened = new DocxSession(saved);
        Assert.Equal(7, Assert.Single(reopened.ListImages().Where(image =>
            image.IntrinsicWidthPixels == 7)).IntrinsicWidthPixels);
        using var document = WordprocessingDocument.Open(new MemoryStream(saved), false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(document)
            .Where(IsMaterialValidationError));
    }

    [Fact]
    public void CC010_Office2013Binding_IsEnumeratedAndFailsClosedUntilExplicitlyDetached()
    {
        var fixture = Transform(BuildFixture(), document =>
        {
            var properties = ControlByNativeId(document, "101").Element(W + "sdtPr")!;
            properties.Add(new XElement(W15 + "dataBinding",
                new XAttribute(W + "storeItemID", "{11111111-1111-1111-1111-111111111111}"),
                new XAttribute(W + "xpath", "/root/value"),
                new XAttribute(W + "prefixMappings", "xmlns:x='urn:test'")));
        });
        var customBefore = CustomXmlBytes(fixture);
        using var session = new DocxSession(fixture);
        var bound = session.ListContentControls().Single(control => control.NativeId == "101");
        Assert.True(bound.IsBound);
        Assert.True(bound.CanDetachTargetBinding);
        Assert.Equal("/root/value", bound.Binding!.XPath);

        var refused = session.FillContentControlText(bound.AnchorId, "refused");
        Assert.Equal(EditErrorCode.ContentControlBound, refused.Error!.Code);
        Assert.Equal(0, session.UndoCount);
        Assert.True(session.FillContentControlText(bound.AnchorId, "detached",
            new ContentControlFillOptions
            {
                BindingPolicy = ContentControlBindingPolicy.DetachTarget,
            }).Success);

        var saved = session.Save();
        Assert.Equal(customBefore, CustomXmlBytes(saved));
        using var document = WordprocessingDocument.Open(new MemoryStream(saved), false);
        Assert.Null(ControlByNativeId(document, "101").Element(W + "sdtPr")?
            .Element(W15 + "dataBinding"));
    }

    [Fact]
    public void CC011_CheckboxMissingChecked_UndoRestoresExactMissingPropertyShape()
    {
        var fixture = Transform(BuildFixture(), document =>
            ControlByNativeId(document, "102").Descendants(W14 + "checked").Single().Remove());
        using var session = new DocxSession(fixture);
        var checkbox = session.ListContentControls().Single(control => control.NativeId == "102");
        Assert.True(session.SetContentControlChecked(checkbox.AnchorId, true).Success);
        Assert.True(session.Undo());

        using var document = WordprocessingDocument.Open(new MemoryStream(session.Save()), false);
        Assert.Empty(ControlByNativeId(document, "102").Descendants(W14 + "checked"));
    }

    [Fact]
    public void CC012_RowAndCellTextControls_ReportUnsupportedPlacementBeforeHistory()
    {
        var fixture = Transform(BuildFixture(), document =>
        {
            var body = document.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!;
            var rowControl = new XElement(W + "sdt",
                new XElement(W + "sdtPr",
                    new XElement(W + "id", new XAttribute(W + "val", "201")),
                    new XElement(W + "text")),
                new XElement(W + "sdtContent",
                    new XElement(W + "tr",
                        new XElement(W + "tc",
                            new XElement(W + "p",
                                new XElement(W + "r", new XElement(W + "t", "row")))))));
            var cellControl = new XElement(W + "sdt",
                new XElement(W + "sdtPr",
                    new XElement(W + "id", new XAttribute(W + "val", "202")),
                    new XElement(W + "text")),
                new XElement(W + "sdtContent",
                    new XElement(W + "tc",
                        new XElement(W + "p",
                            new XElement(W + "r", new XElement(W + "t", "cell"))))));
            body.AddFirst(new XElement(W + "tbl", rowControl,
                new XElement(W + "tr", cellControl)));
        });
        using var session = new DocxSession(fixture);
        var controls = session.ListContentControls();
        var row = controls.Single(control => control.NativeId == "201");
        var cell = controls.Single(control => control.NativeId == "202");
        Assert.Equal(ContentControlPlacement.Row, row.Placement);
        Assert.Equal(ContentControlPlacement.Cell, cell.Placement);
        Assert.False(row.CanMutate);
        Assert.False(cell.CanMutate);
        Assert.Contains("inline and block", row.UnsupportedReason);
        Assert.Equal(EditErrorCode.ContentControlPlacementUnsupported,
            session.FillContentControlText(row.AnchorId, "x").Error!.Code);
        Assert.Equal(EditErrorCode.ContentControlPlacementUnsupported,
            session.FillContentControlText(cell.AnchorId, "x").Error!.Code);
        Assert.Equal(0, session.UndoCount);
    }

    [Fact]
    public void CC013_PictureFill_RefusesNestedControlWithoutBypassingChildLock()
    {
        var fixture = Transform(BuildPictureFixture(), document =>
        {
            var outer = ControlByNativeId(document, "113");
            var run = outer.Descendants(W + "r").Single(value => value.Descendants(W + "drawing").Any());
            run.ReplaceWith(new XElement(W + "sdt",
                new XElement(W + "sdtPr",
                    new XElement(W + "id", new XAttribute(W + "val", "114")),
                    new XElement(W + "lock", new XAttribute(W + "val", "contentLocked")),
                    new XElement(W + "picture")),
                new XElement(W + "sdtContent", new XElement(run))));
        });
        using var session = new DocxSession(fixture);
        var outer = session.ListContentControls().Single(control => control.NativeId == "113");
        var child = session.ListContentControls().Single(control => control.NativeId == "114");
        Assert.False(outer.CanMutate);
        Assert.Contains("nested controls", outer.UnsupportedReason, StringComparison.Ordinal);
        var before = Assert.Single(session.ListImages()).IntrinsicWidthPixels;
        Assert.Equal(EditErrorCode.ContentControlNestedFillUnsupported,
            session.FillContentControlPicture(outer.AnchorId, Png(7, 9)).Error!.Code);
        Assert.Equal(EditErrorCode.ContentControlLocked,
            session.FillContentControlPicture(child.AnchorId, Png(7, 9)).Error!.Code);
        Assert.Equal(0, session.UndoCount);
        Assert.Equal(before, Assert.Single(session.ListImages()).IntrinsicWidthPixels);
    }

    [Fact]
    public void CC014_RepeatingClone_AssignsDistinctDocumentPropertyIdsToEveryDrawing()
    {
        var fixture = Transform(BuildPictureFixture(), document =>
        {
            var drawingRun = ControlByNativeId(document, "113").Descendants(W + "r")
                .Single(value => value.Descendants(W + "drawing").Any());
            var first = new XElement(drawingRun);
            var second = new XElement(drawingRun);
            first.Descendants().Single(value => value.Name.LocalName == "docPr")
                .SetAttributeValue("id", "501");
            second.Descendants().Single(value => value.Name.LocalName == "docPr")
                .SetAttributeValue("id", "502");
            ControlByNativeId(document, "109").Element(W + "sdtContent")!
                .ReplaceNodes(new XElement(W + "p", first, second));
        });
        using var session = new DocxSession(fixture);
        var section = session.ListContentControls().Single(control => control.NativeId == "108");
        Assert.True(session.AddRepeatingSectionItem(section.AnchorId).Success);
        var saved = session.Save();
        using var document = WordprocessingDocument.Open(new MemoryStream(saved), false);
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        var ids = document.MainDocumentPart!.GetXDocument().Descendants(wp + "docPr")
            .Select(value => (string)value.Attribute("id")!).ToList();
        Assert.Equal(ids.Count, ids.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void CC015_RepeatingClone_RejectsCustomXmlMoveAndParagraphIdentities()
    {
        var cases = new Action<XElement>[]
        {
            item => item.Element(W + "sdtContent")!.AddFirst(
                new XElement(W + "customXml", new XElement(W + "p"))),
            item => item.Element(W + "sdtContent")!.AddFirst(
                new XElement(W + "customXmlMoveFromRangeStart",
                    new XAttribute(W + "id", "7"))),
            item => item.Descendants(W + "p").First().SetAttributeValue(W14 + "paraId", "12345678"),
            item => item.Descendants(W + "p").First().SetAttributeValue(W14 + "textId", "87654321"),
        };
        foreach (var arrange in cases)
        {
            var fixture = Transform(BuildFixture(), document =>
                arrange(ControlByNativeId(document, "109")));
            using var session = new DocxSession(fixture);
            var section = session.ListContentControls().Single(control => control.NativeId == "108");
            Assert.Equal(EditErrorCode.RepeatingSectionConstraint,
                session.AddRepeatingSectionItem(section.AnchorId).Error!.Code);
            Assert.Equal(0, session.UndoCount);
        }
    }

    [Fact]
    public void CC016_DuplicateNativeIdsAcrossStories_AreStableDiagnosticsAndNotMutable()
    {
        var fixture = Transform(BuildFixture(), document =>
        {
            var header = document.MainDocumentPart!.AddNewPart<HeaderPart>();
            header.PutXDocument(new XDocument(new XElement(W + "hdr",
                BlockSdt("101", new XElement(W + "text"), "header duplicate"))));
        });
        using var session = new DocxSession(fixture);
        var duplicates = session.ListContentControls().Where(control => control.NativeId == "101").ToList();
        Assert.Equal(2, duplicates.Count);
        Assert.All(duplicates, control =>
        {
            Assert.True(control.HasDuplicateNativeId);
            Assert.False(control.CanMutate);
            Assert.Equal(EditErrorCode.ContentControlMalformed,
                session.FillContentControlText(control.AnchorId, "x").Error!.Code);
        });
        var anchors = duplicates.Select(control => control.AnchorId).OrderBy(value => value).ToArray();
        using var reopened = new DocxSession(session.Save());
        Assert.Equal(anchors, reopened.ListContentControls()
            .Where(control => control.NativeId == "101")
            .Select(control => control.AnchorId).OrderBy(value => value).ToArray());
    }

    [Fact]
    public void CC017_WholeFill_ProtectsBookmarksIncludingTargetsRemovedByTheReplacement()
    {
        static void AddBookmark(XElement control, string name)
        {
            var content = control.Element(W + "sdtContent")!;
            content.AddFirst(new XElement(W + "bookmarkStart",
                new XAttribute(W + "id", "31"), new XAttribute(W + "name", name)));
            content.Add(new XElement(W + "bookmarkEnd", new XAttribute(W + "id", "31")));
        }

        var externallyReferenced = Transform(BuildFixture(), document =>
        {
            AddBookmark(ControlByNativeId(document, "101"), "InnerTarget");
            document.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!.Add(
                new XElement(W + "p", new XElement(W + "hyperlink",
                    new XAttribute(W + "anchor", "InnerTarget"),
                    new XElement(W + "r", new XElement(W + "t", "jump")))));
        });
        using (var session = new DocxSession(externallyReferenced))
        {
            var target = session.ListContentControls().Single(control => control.NativeId == "101");
            var refused = session.FillContentControlText(target.AnchorId, "replacement");
            Assert.Equal(EditErrorCode.BookmarkInUse, refused.Error!.Code);
            Assert.Equal(0, session.UndoCount);
            Assert.Equal("inner", session.GetContentControl(target.AnchorId)!.Text);
        }

        var replacementTarget = Transform(BuildFixture(), document =>
        {
            var control = ControlByNativeId(document, "101");
            control.Element(W + "sdtPr")!.Element(W + "text")!
                .ReplaceWith(new XElement(W + "richText"));
            AddBookmark(control, "RemovedTarget");
        });
        using (var session = new DocxSession(replacementTarget))
        {
            var target = session.ListContentControls().Single(control => control.NativeId == "101");
            var refused = session.FillContentControlRichText(target.AnchorId,
                "[dangling](#RemovedTarget)");
            Assert.Equal(EditErrorCode.MissingBookmarkTarget, refused.Error!.Code);
            Assert.Equal(0, session.UndoCount);
            Assert.Equal("inner", session.GetContentControl(target.AnchorId)!.Text);
        }
    }

    [Fact]
    public void CC018_WholeFill_PromotesNewLinksSweepsOldRelationshipsAndImages_ThroughUndoRedo()
    {
        const string oldUri = "https://old.example.test/value";
        const string newUri = "https://new.example.test/value";
        var fixture = Transform(BuildPictureFixture(), document =>
        {
            var main = document.MainDocumentPart!;
            var control = ControlByNativeId(document, "113");
            control.Element(W + "sdtPr")!.Element(W + "picture")!
                .ReplaceWith(new XElement(W + "richText"));
            var textRun = control.Descendants(W + "r").First(run => run.Descendants(W + "t").Any());
            var relationship = main.AddHyperlinkRelationship(new Uri(oldUri), true);
            textRun.ReplaceWith(new XElement(W + "hyperlink",
                new XAttribute(R + "id", relationship.Id), new XElement(textRun)));
        });

        using var session = new DocxSession(fixture);
        var target = session.ListContentControls().Single(control => control.NativeId == "113");
        Assert.True(session.FillContentControlRichText(target.AnchorId,
            $"[new link]({newUri})").Success);
        Assert.Empty(session.ListImages());
        Assert.Equal(newUri, Assert.Single(session.ListHyperlinks()).Target);

        Assert.True(session.Undo());
        Assert.Single(session.ListImages());
        Assert.Equal(oldUri, Assert.Single(session.ListHyperlinks()).Target);

        Assert.True(session.Redo());
        Assert.Empty(session.ListImages());
        Assert.Equal(newUri, Assert.Single(session.ListHyperlinks()).Target);
        using var saved = WordprocessingDocument.Open(new MemoryStream(session.Save()), false);
        Assert.Empty(saved.MainDocumentPart!.ImageParts);
        var liveRelationship = Assert.Single(saved.MainDocumentPart.HyperlinkRelationships);
        Assert.Equal(newUri, liveRelationship.Uri.ToString());
        Assert.Equal(liveRelationship.Id, saved.MainDocumentPart.GetXDocument().Descendants(W + "hyperlink")
            .Single().Attribute(R + "id")?.Value);
    }

    [Fact]
    public void CC019_RepeatingRemoval_ProtectsBookmarksCleansRelationshipsAndFailsClosedWhenTracked()
    {
        static XElement AddSecondItem(WordprocessingDocument document)
        {
            var first = ControlByNativeId(document, "109");
            var second = new XElement(first);
            second.Element(W + "sdtPr")!.Element(W + "id")!
                .SetAttributeValue(W + "val", "209");
            first.AddAfterSelf(second);
            return second;
        }

        var bookmarked = Transform(BuildFixture(), document =>
        {
            var second = AddSecondItem(document);
            var paragraph = second.Descendants(W + "p").Single();
            paragraph.AddFirst(new XElement(W + "bookmarkStart",
                new XAttribute(W + "id", "41"), new XAttribute(W + "name", "RepeatedTarget")));
            paragraph.Add(new XElement(W + "bookmarkEnd", new XAttribute(W + "id", "41")));
            document.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!.Add(
                new XElement(W + "p", new XElement(W + "hyperlink",
                    new XAttribute(W + "anchor", "RepeatedTarget"),
                    new XElement(W + "r", new XElement(W + "t", "jump")))));
        });
        using (var session = new DocxSession(bookmarked))
        {
            var item = session.ListContentControls().Single(control => control.NativeId == "209");
            var refused = session.RemoveRepeatingSectionItem(item.AnchorId);
            Assert.Equal(EditErrorCode.BookmarkInUse, refused.Error!.Code);
            Assert.Equal(0, session.UndoCount);
            Assert.Equal(2, session.ListContentControls().Count(control =>
                control.Type == ContentControlType.RepeatingSectionItem));
        }

        const string oldUri = "https://removed.example.test/value";
        var relationshipFixture = Transform(BuildPictureFixture(), document =>
        {
            var main = document.MainDocumentPart!;
            var second = AddSecondItem(document);
            var drawingRun = ControlByNativeId(document, "113").Descendants(W + "r")
                .Single(run => run.Descendants(W + "drawing").Any());
            drawingRun.Remove();
            ControlByNativeId(document, "113").Remove();
            var relationship = main.AddHyperlinkRelationship(new Uri(oldUri), true);
            second.Element(W + "sdtContent")!.ReplaceNodes(new XElement(W + "p",
                drawingRun,
                new XElement(W + "hyperlink", new XAttribute(R + "id", relationship.Id),
                    new XElement(W + "r", new XElement(W + "t", "removed link")))));
        });

        using (var tracked = new DocxSession(relationshipFixture))
        {
            tracked.SetTrackedChanges(TrackedChangeMode.RenderInline);
            var item = tracked.ListContentControls().Single(control => control.NativeId == "209");
            Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
                tracked.RemoveRepeatingSectionItem(item.AnchorId).Error!.Code);
            Assert.Equal(0, tracked.UndoCount);
            Assert.Single(tracked.ListImages());
        }

        using (var session = new DocxSession(relationshipFixture))
        {
            var item = session.ListContentControls().Single(control => control.NativeId == "209");
            Assert.True(session.RemoveRepeatingSectionItem(item.AnchorId).Success);
            Assert.Empty(session.ListImages());
            Assert.Empty(session.ListHyperlinks());
            Assert.True(session.Undo());
            Assert.Single(session.ListImages());
            Assert.Equal(oldUri, Assert.Single(session.ListHyperlinks()).Target);
            Assert.True(session.Redo());
            Assert.Empty(session.ListImages());
            Assert.Empty(session.ListHyperlinks());
        }
    }

    private static string[] ParagraphAnchors(DocxSession session) => session.Project().AnchorIndex.Values
        .Where(value => value.Anchor.Kind is "p" or "h" or "li")
        .Select(value => value.Anchor.Id).Distinct().ToArray();

    private static XElement ControlByNativeId(WordprocessingDocument document, string nativeId) =>
        document.MainDocumentPart!.GetXDocument().Descendants(W + "sdt").Concat(
            document.MainDocumentPart.HeaderParts.SelectMany(header =>
                header.GetXDocument().Descendants(W + "sdt")))
        .Single(value => (string?)value.Element(W + "sdtPr")?.Element(W + "id")?
            .Attribute(W + "val") == nativeId);

    private static byte[] Transform(byte[] bytes, Action<WordprocessingDocument> transform)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            transform(document);
            document.MainDocumentPart!.PutXDocument();
            foreach (var header in document.MainDocumentPart.HeaderParts)
                header.PutXDocument();
        }
        return stream.ToArray();
    }

    internal static byte[] BuildFixture()
    {
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var main = doc.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W + "body")!;
            body.Elements().Where(value => value.Name != W + "sectPr").Remove();
            body.AddFirst(
                BlockSdt("100", new XElement(W + "richText"), "outer value",
                    tag: "outer-tag", alias: "Outer alias", nestedInline: Sdt("101",
                        new XElement(W + "text"), "inner")),
                BlockSdt("102", new XElement(W14 + "checkbox",
                    new XElement(W14 + "checked", new XAttribute(W14 + "val", "0")),
                    new XElement(W14 + "checkedState", new XAttribute(W14 + "val", "2612")),
                    new XElement(W14 + "uncheckedState", new XAttribute(W14 + "val", "2610"))), "☐"),
                BlockSdt("103", new XElement(W + "date",
                    new XElement(W + "dateFormat", new XAttribute(W + "val", "MMMM d, yyyy"))), "date"),
                BlockSdt("104", new XElement(W + "dropDownList",
                    Item("Alpha", "a"), Item("Beta", "b")), "pick"),
                BlockSdt("105", new XElement(W + "comboBox",
                    Item("Alpha", "a"), Item("Beta", "b")), "pick"),
                BlockSdt("106", new XElement(W + "text"), "bound",
                    binding: new XElement(W + "dataBinding",
                        new XAttribute(W + "storeItemID", "{11111111-1111-1111-1111-111111111111}"),
                        new XAttribute(W + "xpath", "/root/value"),
                        new XAttribute(W + "prefixMappings", "xmlns:x='urn:test'"))),
                BlockSdt("107", new XElement(W + "text"), "duplicate one"),
                BlockSdt("107", new XElement(W + "text"), "duplicate two"),
                BlockSdt(null, new XElement(W + "text"), "missing id"),
                RepeatingSection(),
                BlockSdt("110", new XElement(W + "group"), "unsupported"),
                BlockSdt("111", new XElement(W + "richText"), "locked outer",
                    lockToken: "contentLocked", nestedInline: Sdt("112",
                        new XElement(W + "text"), "locked child")));
            main.PutXDocument();

            var custom = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var input = new MemoryStream(Encoding.UTF8.GetBytes("<root><value>bound</value></root>"));
            custom.FeedData(input);
        }
        return stream.ToArray();
    }

    private static byte[] BuildPictureFixture()
    {
        var stream = new MemoryStream();
        stream.Write(BuildFixture());
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W + "body")!;
            body.Add(BlockSdt("113", new XElement(W + "picture"),
                "picture placeholder", tag: "picture-tag"));
            main.PutXDocument();
        }
        using var seed = new DocxSession(stream.ToArray());
        var paragraph = seed.Project().AnchorIndex.Values.Single(value =>
            value.Anchor.Kind == "p" && value.TextPreview.Contains("picture placeholder",
                StringComparison.Ordinal));
        Assert.True(seed.InsertImage(paragraph.Anchor.Id, 0, Png(2, 3)).Success);
        return seed.Save();
    }

    private static byte[] Png(int width, int height)
    {
        var bytes = new byte[24];
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R' }.CopyTo(bytes, 0);
        bytes[16] = (byte)(width >> 24); bytes[17] = (byte)(width >> 16);
        bytes[18] = (byte)(width >> 8); bytes[19] = (byte)width;
        bytes[20] = (byte)(height >> 24); bytes[21] = (byte)(height >> 16);
        bytes[22] = (byte)(height >> 8); bytes[23] = (byte)height;
        return bytes;
    }

    private static XElement Sdt(string? id, XElement type, string text,
        string? tag = null, string? alias = null, string? lockToken = null,
        XElement? binding = null, XElement? nestedInline = null) =>
        new(W + "sdt",
            new XElement(W + "sdtPr",
                id is null ? null : new XElement(W + "id", new XAttribute(W + "val", id)),
                tag is null ? null : new XElement(W + "tag", new XAttribute(W + "val", tag)),
                alias is null ? null : new XElement(W + "alias", new XAttribute(W + "val", alias)),
                lockToken is null ? null : new XElement(W + "lock", new XAttribute(W + "val", lockToken)),
                binding, type),
            new XElement(W + "sdtContent",
                new XElement(W + "r", new XElement(W + "t", text)), nestedInline));

    private static XElement BlockSdt(string? id, XElement type, string text,
        string? tag = null, string? alias = null, string? lockToken = null,
        XElement? binding = null, XElement? nestedInline = null)
    {
        var inline = Sdt(id, type, text, tag, alias, lockToken, binding, nestedInline);
        inline.Element(W + "sdtContent")!.ReplaceNodes(new XElement(W + "p",
            new XElement(W + "r", new XElement(W + "t", text)), nestedInline));
        return inline;
    }

    private static XElement Item(string display, string value) =>
        new(W + "listItem", new XAttribute(W + "displayText", display),
            new XAttribute(W + "value", value));

    private static XElement RepeatingSection() =>
        new(W + "sdt",
            new XElement(W + "sdtPr",
                new XElement(W + "id", new XAttribute(W + "val", "108")),
                new XElement(W15 + "repeatingSection")),
            new XElement(W + "sdtContent",
                new XElement(W + "sdt",
                    new XElement(W + "sdtPr",
                        new XElement(W + "id", new XAttribute(W + "val", "109")),
                        new XElement(W15 + "repeatingSectionItem")),
                    new XElement(W + "sdtContent",
                        new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", "item")))))));

    private static byte[] CustomXmlBytes(byte[] bytes)
    {
        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        using var input = doc.MainDocumentPart!.CustomXmlParts.Single().GetStream();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static bool IsMaterialValidationError(ValidationErrorInfo error) =>
        error.Description?.Contains("The 'Ignorable' attribute", StringComparison.Ordinal) != true
        && error.Description?.Contains("http://powertools.codeplex.com/2011:Unid", StringComparison.Ordinal) != true;
}
