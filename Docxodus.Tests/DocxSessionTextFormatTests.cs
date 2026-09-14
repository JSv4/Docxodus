// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Reflection;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionTextFormatTests
{
    [Theory]
    [InlineData(TrackedChangeMode.Accept, 6, 9, "new text", "First new text.")]
    [InlineData(TrackedChangeMode.RenderInline, 6, 9, "new text", "First new text.")]
    [InlineData(TrackedChangeMode.Accept, 0, 0, "New ", "New First paragraph.")]
    [InlineData(TrackedChangeMode.RenderInline, 16, 0, " End", "First paragraph. End")]
    [InlineData(TrackedChangeMode.Accept, 6, 9, "", "First .")]
    public void TF788_ReplacementAndFormat_AreOneUndoUnit(
        TrackedChangeMode tracking, int start, int length, string replacement, string expected)
    {
        using var session = Open(tracking);
        var anchor = session.FindByText("First paragraph.")!.Anchor.Id;
        var before = session.GetPackageContentHash();

        var edit = session.ReplaceTextAtSpanWithFormat(
            anchor, start, length, replacement, new FormatOp { Bold = true });

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal(anchor, Assert.Single(edit.Modified).Id);
        Assert.Null(edit.Patch);
        Assert.Equal(1, session.Version);
        Assert.Equal(1, session.UndoCount);
        var runs = session.GetFormatting(anchor)!.Runs;
        Assert.Equal(expected, string.Concat(runs.Select(r => r.Text)));
        var boldText = string.Concat(runs.Where(r => r.Effective.Bold is true).Select(r => r.Text));
        Assert.Equal(replacement, boldText);
        var after = session.GetPackageContentHash();
        Assert.NotEqual(before, after);

        Assert.True(session.Undo());
        Assert.Equal(before, session.GetPackageContentHash());
        Assert.False(session.Undo());
        Assert.True(session.Redo());
        Assert.Equal(after, session.GetPackageContentHash());
        Assert.False(session.Redo());
    }

    [Theory]
    [InlineData(TrackedChangeMode.Accept)]
    [InlineData(TrackedChangeMode.RenderInline)]
    public void TF788_FormatFailure_RestoresTextStylesGeneratorsAndRedo(TrackedChangeMode tracking)
    {
        using var session = Open(tracking);
        var anchor = session.FindByText("First paragraph.")!.Anchor.Id;
        Assert.True(session.ReplaceTextAtSpan(anchor, 0, 5, "Redo").Success);
        var redoHash = session.GetPackageContentHash();
        Assert.True(session.Undo());
        var before = session.GetPackageContentHash();
        var version = session.Version;
        var counter = Field<int>(session, "_revisionCounter");
        var ticks = Field<long>(session, "_lastFormatRevisionTicks");

        // Code synthesizes a style before the invalid highlight throws. Both the text edit and
        // this partially applied formatting operation must roll back, including the redo stack.
        var edit = session.ReplaceTextAtSpanWithFormat(anchor, 6, 9, "new text",
            new FormatOp { Code = true, Bold = true, Highlight = "invalid-highlight" });

        Assert.False(edit.Success);
        Assert.Equal(EditErrorCode.InternalError, edit.Error!.Code);
        Assert.Equal(before, session.GetPackageContentHash());
        Assert.Equal(version, session.Version);
        Assert.Equal(counter, Field<int>(session, "_revisionCounter"));
        Assert.Equal(ticks, Field<long>(session, "_lastFormatRevisionTicks"));
        Assert.Equal(0, session.UndoCount);
        Assert.Equal(1, session.RedoCount);
        Assert.True(session.Redo());
        Assert.Equal(redoHash, session.GetPackageContentHash());
    }

    [Fact]
    public void TF788_EnclosingBatch_KeepsReceiptAndFullPackageRollback()
    {
        using var session = Open(TrackedChangeMode.Accept);
        var anchor = session.FindByText("First paragraph.")!.Anchor.Id;
        var before = session.GetPackageContentHash();
        var step = new MutationBatchStep("edit", "text_format",
            s => s.ReplaceTextAtSpanWithFormat(anchor, 6, 9, "new text", new FormatOp { Code = true }));

        var failed = session.ExecuteBatch(new[]
        {
            step,
            new MutationBatchStep("create", "header",
                s => s.SetHeaderText(anchor, HeaderFooterKind.Default, "Speculative header")),
            new MutationBatchStep("edit", "missing", s => s.ReplaceText("p:body:missing", "fail")),
        });
        Assert.False(failed.Success);
        Assert.True(failed.RolledBack);
        Assert.Equal(before, failed.PackageHash);
        Assert.Equal(0, session.UndoCount);
        Assert.Equal(0, session.Version);

        var applied = session.ExecuteBatch(new[] { step });
        Assert.True(applied.Success);
        Assert.Equal(1, applied.ResultVersion);
        Assert.Equal(session.GetPackageContentHash(), applied.PackageHash);
        Assert.NotEqual(before, applied.PackageHash);
        Assert.True(session.Undo());
        Assert.Equal(before, session.GetPackageContentHash());
        Assert.True(session.Redo());
        Assert.Equal(applied.PackageHash, session.GetPackageContentHash());
    }

    [Fact]
    public void TF788_DeliveryCapture_PreservesConsecutiveEditVersions()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDocWithoutCodeStyle(),
            new DocxSessionSettings { CaptureDeliveryEvidence = true });
        var anchor = session.FindByText("First paragraph.")!.Anchor.Id;
        Assert.True(session.ReplaceTextAtSpanWithFormat(anchor, 0, 5, "First", new FormatOp { Bold = true }).Success);
        Assert.True(session.ReplaceTextAtSpanWithFormat(anchor, 6, 9, "text", new FormatOp { Italic = true }).Success);
        var status = session.GetDeliveryEvidenceStatus();
        Assert.Null(status.UnavailableReason);
        Assert.Equal(2, status.TransactionCount);
        Assert.Equal(2, status.CurrentVersion);
        Assert.True(session.BuildDeliveryReceipt().Verification.IsValid);
    }

    private static DocxSession Open(TrackedChangeMode tracking) =>
        new(DocxSessionTests.BuildDocWithoutCodeStyle(), new DocxSessionSettings
        {
            PersistAnchorIds = true,
            EmitMarkdownPatch = false,
            TrackedChanges = tracking,
            UndoDepth = 1,
        });

    private static T Field<T>(DocxSession session, string name) =>
        (T)typeof(DocxSession).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!
            .GetValue(session)!;
}
