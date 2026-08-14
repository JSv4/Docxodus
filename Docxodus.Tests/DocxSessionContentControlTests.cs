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

            using var invalidDate = JsonDocument.Parse(
                Docxodus.Internal.DocxSessionOps.SetContentControlDate(
                handle, controls.Single(control => control.TryGetProperty("nativeId", out var id)
                    && id.GetString() == "103")
                        .GetProperty("anchorId").GetString()!, "not-a-date", null, "{}"));
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

    private static string[] ParagraphAnchors(DocxSession session) => session.Project().AnchorIndex.Values
        .Where(value => value.Anchor.Kind is "p" or "h" or "li")
        .Select(value => value.Anchor.Id).Distinct().ToArray();

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
