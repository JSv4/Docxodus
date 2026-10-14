// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #763: explicit nested-fill semantics for whole-control fills, native tracked
/// representations for the operations that have one, and a per-operation capability matrix
/// evaluated by the same gate the operations apply.
/// </summary>
public class DocxSessionContentControlNestedTrackedTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void CC763a_PreserveKeepsNestedControlsAndFillsNamedChildren()
    {
        using var session = new DocxSession(DocxSessionContentControlTests.BuildFixture());
        var outer = Control(session, "100");
        var inner = Control(session, "101");
        Assert.Equal(new[] { inner.AnchorId }, outer.NestedControlAnchorIds);
        var preserve = outer.Operations.Single(op => op.Operation == "fill_rich_text" && op.NestedControls == "preserve");
        Assert.True(preserve.CanMutate, preserve.Reason);
        Assert.False(outer.Operations.Single(op => op.Operation == "fill_rich_text" && op.NestedControls == "refuse").CanMutate);

        var result = session.FillContentControlRichText(outer.AnchorId, "Outer **text**", new ContentControlFillOptions
        {
            NestedControls = ContentControlNestedPolicy.Preserve,
            ChildFills = new Dictionary<string, string> { [inner.AnchorId] = "inner filled" },
        });

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(new[] { outer.AnchorId, inner.AnchorId }, result.Modified.Select(a => a.Id));
        Assert.Empty(result.Removed);
        var after = Control(session, "100");
        Assert.Equal(new[] { inner.AnchorId }, after.NestedControlAnchorIds);
        Assert.Equal("inner filled", Control(session, "101").Text);
        Assert.Contains("Outer text", after.Text, StringComparison.Ordinal);
        Assert.Equal("outer-tag", after.Tag);
        Assert.Equal("Outer alias", after.Alias);
        Assert.True(session.Undo());
        Assert.Equal("inner", Control(session, "101").Text);
        Assert.Contains("outer value", Control(session, "100").Text, StringComparison.Ordinal);
    }

    [Fact]
    public void CC763b_ReplaceDropsNestedControlsExplicitlyAndProtectsLockedOrBoundChildren()
    {
        using var session = new DocxSession(DocxSessionContentControlTests.BuildFixture());
        var outer = Control(session, "100");
        var inner = Control(session, "101");

        var result = session.FillContentControlText(outer.AnchorId, "flat", new ContentControlFillOptions
        {
            NestedControls = ContentControlNestedPolicy.Replace,
        });

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(new[] { inner.AnchorId }, result.Removed.Select(a => a.Id));
        Assert.Equal("flat", Control(session, "100").Text);
        Assert.Empty(Control(session, "100").NestedControlAnchorIds);
        Assert.DoesNotContain(session.ListContentControls(), c => c.NativeId == "101");
        Assert.True(session.Undo());
        Assert.Contains(session.ListContentControls(), c => c.NativeId == "101");

        // A locked nested child is never discarded by implication.
        var lockedOuter = Control(session, "111");
        var refused = session.FillContentControlText(lockedOuter.AnchorId, "x", new ContentControlFillOptions
        {
            NestedControls = ContentControlNestedPolicy.Replace,
        });
        Assert.Equal(EditErrorCode.ContentControlLocked, refused.Error!.Code);
        Assert.Equal(0, session.UndoCount);
    }

    [Fact]
    public void CC763c_ChildFillsMustNameNestedTextualControlsAndFailWithoutMutation()
    {
        using var session = new DocxSession(DocxSessionContentControlTests.BuildFixture());
        var outer = Control(session, "100");
        var checkbox = Control(session, "102");

        var unknown = session.FillContentControlText(outer.AnchorId, "x", new ContentControlFillOptions
        {
            NestedControls = ContentControlNestedPolicy.Preserve,
            ChildFills = new Dictionary<string, string> { ["sdt:body:missing"] = "y" },
        });
        Assert.Equal(EditErrorCode.ContentControlNotFound, unknown.Error!.Code);

        var notNested = session.FillContentControlText(outer.AnchorId, "x", new ContentControlFillOptions
        {
            NestedControls = ContentControlNestedPolicy.Preserve,
            ChildFills = new Dictionary<string, string> { [checkbox.AnchorId] = "y" },
        });
        Assert.Equal(EditErrorCode.ContentControlNotFound, notNested.Error!.Code);

        var wrongPolicy = session.FillContentControlText(outer.AnchorId, "x", new ContentControlFillOptions
        {
            NestedControls = ContentControlNestedPolicy.Replace,
            ChildFills = new Dictionary<string, string> { [Control(session, "101").AnchorId] = "y" },
        });
        Assert.Equal(EditErrorCode.ContentControlNestedFillUnsupported, wrongPolicy.Error!.Code);

        // The default policy still refuses, and names the option that would not.
        var refused = session.FillContentControlText(outer.AnchorId, "x");
        Assert.Equal(EditErrorCode.ContentControlNestedFillUnsupported, refused.Error!.Code);
        Assert.Contains("nestedControls", refused.Error.Message, StringComparison.Ordinal);
        Assert.Equal(0, session.UndoCount);
        Assert.Equal("outer value", Control(session, "100").Text.Replace("inner", string.Empty).Trim());
    }

    [Fact]
    public void CC763d_TrackedTextFill_AcceptsToThePayloadAndRejectsToTheOriginal()
    {
        var fixture = DocxSessionContentControlTests.BuildFixture();
        foreach (var (rich, payload, expected) in new[]
                 {
                     (false, "tracked plain", "tracked plain"),
                     // Preserve keeps the nested control (and its text) in the host paragraph.
                     (true, "First **bold**\n\nSecond paragraph", "First boldinnerSecond paragraph"),
                 })
        {
            using var session = new DocxSession(fixture, new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "Reviewer",
            });
            var target = Control(session, rich ? "100" : "101");
            var entry = target.Operations.Single(op => op.Operation == (rich ? "fill_rich_text" : "fill_text")
                && op.NestedControls is null or "preserve");
            Assert.True(entry.CanMutate, entry.Reason);

            var options = new ContentControlFillOptions { NestedControls = ContentControlNestedPolicy.Preserve };
            var result = rich
                ? session.FillContentControlRichText(target.AnchorId, payload, options)
                : session.FillContentControlText(target.AnchorId, payload, options);
            Assert.True(result.Success, result.Error?.Message);
            Assert.Equal(1, session.UndoCount);
            var revisions = session.ListRevisions();
            Assert.Contains(revisions, r => r.Type == "insert" && r.Author == "Reviewer");
            Assert.Contains(revisions, r => r.Type == "delete" && r.Author == "Reviewer");

            var accepted = session.Save();
            using (var acceptedSession = new DocxSession(accepted))
            {
                Assert.True(acceptedSession.AcceptAllRevisions().Success);
                Assert.Empty(acceptedSession.ListRevisions());
                var control = ControlByNativeId(acceptedSession, rich ? "100" : "101");
                Assert.Equal(expected, Squash(control.Text));
                Assert.Equal(rich ? "outer-tag" : null, control.Tag);
                if (rich) Assert.Equal("inner", ControlByNativeId(acceptedSession, "101").Text);
            }
            using (var rejectedSession = new DocxSession(accepted))
            {
                Assert.True(rejectedSession.RejectAllRevisions().Success);
                Assert.Empty(rejectedSession.ListRevisions());
                var control = ControlByNativeId(rejectedSession, rich ? "100" : "101");
                Assert.Equal(rich ? "outer valueinner" : "inner", Squash(control.Text));
            }
        }
    }

    [Fact]
    public void CC763e_TrackedStateChangesRefuseWithASpecificReason()
    {
        using var session = new DocxSession(DocxSessionContentControlTests.BuildFixture(), new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
        });
        var checkbox = Control(session, "102");
        var date = Control(session, "103");
        var list = Control(session, "104");
        Assert.False(checkbox.CanMutate);
        Assert.Contains("w14:checked", checkbox.UnsupportedReason, StringComparison.Ordinal);
        Assert.Contains("w:date", date.Operations.Single().Reason, StringComparison.Ordinal);
        Assert.Contains("w:lastValue", list.Operations.Single().Reason, StringComparison.Ordinal);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
            session.SetContentControlChecked(checkbox.AnchorId, true).Error!.Code);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
            session.SetContentControlDate(date.AnchorId, DateTimeOffset.UnixEpoch).Error!.Code);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
            session.SelectContentControlItem(list.AnchorId, "a").Error!.Code);
        Assert.Equal(0, session.UndoCount);

        // Text controls are mutable in tracked mode, and the registry says so before the fact.
        Assert.True(Control(session, "101").CanMutate);
        Assert.Null(Control(session, "101").UnsupportedReason);
    }

    [Fact]
    public void CC763f_TrackedRepeatingItems_InsertAndDeleteAsContentControlRevisions()
    {
        using var session = new DocxSession(DocxSessionContentControlTests.BuildFixture(), new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
        });
        var section = Control(session, "108");
        var item = Control(session, "109");
        Assert.True(section.Operations.Single().CanMutate, section.Operations.Single().Reason);

        var added = session.AddRepeatingSectionItem(section.AnchorId);
        Assert.True(added.Success, added.Error?.Message);
        var createdId = Assert.Single(added.Created).Id;
        Assert.Contains(session.ListRevisions(), r => r.Family == RevisionFamily.ContentControlInsert);

        var removed = session.RemoveRepeatingSectionItem(item.AnchorId);
        Assert.True(removed.Success, removed.Error?.Message);
        Assert.Empty(removed.Removed);
        Assert.Contains(item.AnchorId, removed.Modified.Select(a => a.Id));
        Assert.Contains(session.ListRevisions(), r => r.Family == RevisionFamily.ContentControlDelete);
        Assert.Equal(2, session.ListContentControls().Count(c => c.Type == ContentControlType.RepeatingSectionItem));

        var saved = session.Save();
        using (var accepted = new DocxSession(saved))
        {
            Assert.True(accepted.AcceptAllRevisions().Success);
            Assert.Empty(accepted.ListRevisions());
            var items = accepted.ListContentControls().Where(c => c.Type == ContentControlType.RepeatingSectionItem).ToList();
            Assert.Single(items);
            Assert.NotEqual("109", items[0].NativeId);
        }
        using (var rejected = new DocxSession(saved))
        {
            Assert.True(rejected.RejectAllRevisions().Success);
            Assert.Empty(rejected.ListRevisions());
            var items = rejected.ListContentControls().Where(c => c.Type == ContentControlType.RepeatingSectionItem).ToList();
            Assert.Single(items);
            Assert.Equal("109", items[0].NativeId);
        }
    }

    [Fact]
    public void CC763g_TrackedPictureFill_KeepsBothImagesUntilTheRevisionResolves()
    {
        using var session = new DocxSession(PictureFixture(), new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
        });
        var picture = Control(session, "113");
        Assert.True(picture.CanMutate, picture.UnsupportedReason);
        var before = Assert.Single(session.ListImages());

        var result = session.FillContentControlPicture(picture.AnchorId, Png(6, 7));
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(2, session.ListImages().Count);
        Assert.Contains(session.ListRevisions(), r => r.Type == "insert");
        Assert.Contains(session.ListRevisions(), r => r.Type == "delete");

        var saved = session.Save();
        using (var accepted = new DocxSession(saved))
        {
            Assert.True(accepted.AcceptAllRevisions().Success);
            var image = Assert.Single(accepted.ListImages());
            Assert.Equal(6, image.IntrinsicWidthPixels);
            Assert.NotEqual(before.TargetPartUri, image.TargetPartUri);
        }
        using (var rejected = new DocxSession(saved))
        {
            Assert.True(rejected.RejectAllRevisions().Success);
            var image = Assert.Single(rejected.ListImages());
            Assert.Equal(2, image.IntrinsicWidthPixels);
        }
    }

    [Fact]
    public void CC763h_OperationsMatrix_IsTheSameGateTheOperationsApply()
    {
        using var session = new DocxSession(DocxSessionContentControlTests.BuildFixture());
        foreach (var control in session.ListContentControls())
        {
            foreach (var support in control.Operations)
            {
                var options = new ContentControlFillOptions
                {
                    NestedControls = support.NestedControls switch
                    {
                        "preserve" => ContentControlNestedPolicy.Preserve,
                        "replace" => ContentControlNestedPolicy.Replace,
                        _ => ContentControlNestedPolicy.Refuse,
                    },
                };
                using var probe = new DocxSession(DocxSessionContentControlTests.BuildFixture());
                var attempt = support.Operation switch
                {
                    "fill_text" => probe.FillContentControlText(control.AnchorId, "probe", options),
                    "fill_rich_text" => probe.FillContentControlRichText(control.AnchorId, "probe", options),
                    "set_checked" => probe.SetContentControlChecked(control.AnchorId, true, options),
                    "set_date" => probe.SetContentControlDate(control.AnchorId, DateTimeOffset.UnixEpoch, null, options),
                    "select_item" => probe.SelectContentControlItem(control.AnchorId, "a", options),
                    "add_repeating_item" => probe.AddRepeatingSectionItem(control.AnchorId, null, options),
                    "remove_repeating_item" => probe.RemoveRepeatingSectionItem(control.AnchorId),
                    _ => throw new InvalidOperationException(support.Operation),
                };
                Assert.True(support.CanMutate == attempt.Success,
                    $"{control.NativeId}/{support.Operation}/{support.NestedControls}: registry {support.CanMutate} ({support.Reason}) vs mutation {attempt.Success} ({attempt.Error?.Message})");
                if (!attempt.Success) Assert.Equal(attempt.Error!.Message, support.Reason);
            }
        }
    }

    private static ContentControlInfo Control(DocxSession session, string nativeId) =>
        session.ListContentControls().First(control => control.NativeId == nativeId);

    private static ContentControlInfo ControlByNativeId(DocxSession session, string nativeId) =>
        session.ListContentControls().Single(control => control.NativeId == nativeId);

    private static string Squash(string text) => text.Replace("\n", string.Empty).Replace("\r", string.Empty);

    private static byte[] PictureFixture()
    {
        var stream = new MemoryStream();
        stream.Write(DocxSessionContentControlTests.BuildFixture());
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W + "body")!;
            body.Add(new XElement(W + "sdt",
                new XElement(W + "sdtPr",
                    new XElement(W + "id", new XAttribute(W + "val", "113")),
                    new XElement(W + "tag", new XAttribute(W + "val", "picture-tag")),
                    new XElement(W + "picture")),
                new XElement(W + "sdtContent",
                    new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", "picture placeholder"))))));
            main.PutXDocument();
        }
        using var seed = new DocxSession(stream.ToArray());
        var paragraph = seed.Project().AnchorIndex.Values.Single(value =>
            value.Anchor.Kind == "p" && value.TextPreview.Contains("picture placeholder", StringComparison.Ordinal));
        Assert.True(seed.InsertImage(paragraph.Anchor.Id, 0, Png(2, 3)).Success);
        return seed.Save();
    }

    private static byte[] Png(int width, int height)
    {
        using var stream = new MemoryStream();
        stream.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        WriteChunk(stream, "IHDR", writer =>
        {
            writer.Write(BigEndian(width));
            writer.Write(BigEndian(height));
            writer.Write(new byte[] { 8, 6, 0, 0, 0 });
        });
        WriteChunk(stream, "IDAT", writer => writer.Write(new byte[] { 0x78, 0x9C, 0x63, 0x60, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01 }));
        WriteChunk(stream, "IEND", _ => { });
        return stream.ToArray();
    }

    private static byte[] BigEndian(int value) => new[]
    {
        (byte)(value >> 24), (byte)(value >> 16), (byte)(value >> 8), (byte)value,
    };

    private static void WriteChunk(Stream stream, string type, Action<BinaryWriter> body)
    {
        using var payload = new MemoryStream();
        using (var writer = new BinaryWriter(payload, System.Text.Encoding.ASCII, leaveOpen: true))
            body(writer);
        var data = payload.ToArray();
        stream.Write(BigEndian(data.Length));
        var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        stream.Write(typeBytes);
        stream.Write(data);
        var crc = Crc32(typeBytes.Concat(data).ToArray());
        stream.Write(BigEndian((int)crc));
    }

    private static uint Crc32(byte[] bytes)
    {
        uint crc = 0xFFFFFFFF;
        foreach (var b in bytes)
        {
            crc ^= b;
            for (int i = 0; i < 8; i++)
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
        }
        return ~crc;
    }
}
