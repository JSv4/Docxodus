// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
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
    [InlineData(TrackedChangeMode.Accept, 9, false)]
    [InlineData(TrackedChangeMode.RenderInline, 9, false)]
    [InlineData(TrackedChangeMode.Accept, 0, false)]
    [InlineData(TrackedChangeMode.RenderInline, 0, false)]
    [InlineData(TrackedChangeMode.Accept, 0, true)]
    [InlineData(TrackedChangeMode.RenderInline, 0, true)]
    public void TF788_FormatFailure_RestoresTextStylesGeneratorsAndRedo(
        TrackedChangeMode tracking, int length, bool leadingTab)
    {
        using var session = leadingTab
            ? Open(tracking, new XElement(W.r, new XElement(W.tab), new XElement(W.t, "First paragraph.")))
            : Open(tracking);
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
        var edit = session.ReplaceTextAtSpanWithFormat(anchor, 6, length, "new text",
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

    [Theory]
    [InlineData(TrackedChangeMode.Accept, 0)]
    [InlineData(TrackedChangeMode.RenderInline, 0)]
    [InlineData(TrackedChangeMode.Accept, 1)]
    [InlineData(TrackedChangeMode.RenderInline, 1)]
    [InlineData(TrackedChangeMode.Accept, 2)]
    public void TF799_InteriorInsertion_PreservesNeighborFormatting(TrackedChangeMode tracking, int leadingTabs)
    {
        // The caret is between two text nodes in one italic run, after a UTF-16 surrogate pair.
        var source = new XElement(W.r, new XElement(W.rPr, new XElement(W.i)),
            Enumerable.Range(0, leadingTabs).Select(_ => new XElement(W.tab)),
            new XElement(W.t, "doc😀u"), new XElement(W.t, "ment"));
        using var session = Open(tracking,
            new XElement(W.p, ComplexField(new XElement(W.r, new XElement(W.t, "Title")))),
            new XElement(W.p, source));
        var anchor = session.FindByText("doc😀ument")!.Anchor.Id;
        var before = session.GetPackageContentHash();

        var edit = session.ReplaceTextAtSpanWithFormat(anchor, 6, 0, " inserted ",
            new FormatOp { Bold = true, Italic = false });

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal(1, session.Version);
        Assert.Equal(1, session.UndoCount);
        Assert.Equal(new[] { ("doc😀u", false, true), (" inserted ", true, false), ("ment", false, true) },
            session.GetFormatting(anchor)!.Runs.Select(r => (r.Text, r.Effective.Bold is true, r.Effective.Italic is true)));
        var xml = XElement.Parse(session.Raw.GetXml(anchor));
        Assert.Equal(leadingTabs, xml.Descendants(W.tab).Count());
        Assert.Equal(Enumerable.Repeat(W.tab, leadingTabs).Append(W.t),
            xml.Elements(W.r).First().Elements().Where(e => e.Name != W.rPr).Select(e => e.Name));
        Assert.All(xml.Descendants(W.t), t => Assert.Equal("preserve", (string?)t.Attribute(XNamespace.Xml + "space")));
        if (tracking == TrackedChangeMode.RenderInline)
        {
            Assert.Equal(" inserted ", Assert.Single(xml.Elements(W.ins)).Element(W.r)!.Element(W.t)!.Value);
            Assert.Empty(xml.Descendants(W.del));
        }
        var after = session.GetPackageContentHash();
        Assert.True(session.Undo());
        Assert.Equal(before, session.GetPackageContentHash());
        Assert.True(session.Redo());
        Assert.Equal(after, session.GetPackageContentHash());
    }

    [Theory]
    [InlineData("complex field")]
    [InlineData("field across paragraphs")]
    [InlineData("deleted field end")]
    [InlineData("moved field end")]
    [InlineData("simple field")]
    [InlineData("insertion")]
    [InlineData("mixed run")]
    [InlineData("trailing tab")]
    [InlineData("leading break")]
    [InlineData("tab/text field")]
    [InlineData("tab/text hyperlink")]
    [InlineData("tab/text surrogate pair")]
    [InlineData("surrogate pair")]
    public void TF799_UnsafeInteriorInsertion_IsUnchanged(string context)
    {
        var run = new XElement(W.r, new XElement(W.t, "document"));
        var content = context switch
        {
            "complex field" => ComplexField(run),
            "deleted field end" or "moved field end" => ComplexField(
                new XElement(context == "deleted field end" ? W.del : W.moveFrom,
                    new XAttribute(W.id, "1"), new XAttribute(W.author, "Other"), FieldRun("end")), run),
            "field across paragraphs" => new[]
            {
                new XElement(W.p, FieldRun("begin"),
                    new XElement(W.r, new XElement(W.instrText, " TOC ")), FieldRun("separate")),
                new XElement(W.p, run),
                new XElement(W.p, FieldRun("end")),
            },
            "simple field" => new[] { new XElement(W.fldSimple, new XAttribute(W.instr, " DOCPROPERTY Title "), run) },
            "insertion" => new[] { new XElement(W.ins, new XAttribute(W.id, "1"), new XAttribute(W.author, "Other"), run) },
            "mixed run" => new[] { new XElement(W.r, new XElement(W.t, "docu"), new XElement(W.tab), new XElement(W.t, "ment")) },
            "trailing tab" => new[] { new XElement(W.r, new XElement(W.t, "document"), new XElement(W.tab)) },
            "leading break" => new[] { new XElement(W.r, new XElement(W.br), new XElement(W.t, "document")) },
            "tab/text field" => ComplexField(new XElement(W.r, new XElement(W.tab), new XElement(W.t, "document"))),
            "tab/text hyperlink" => new[] { new XElement(W.hyperlink, new XAttribute(W.anchor, "top"),
                new XElement(W.r, new XElement(W.tab), new XElement(W.t, "document"))) },
            "tab/text surrogate pair" => new[] { new XElement(W.r, new XElement(W.tab), new XElement(W.t, "doc😀ument")) },
            "surrogate pair" => new[] { new XElement(W.r, new XElement(W.t, "doc😀ument")) },
            _ => throw new ArgumentOutOfRangeException(nameof(context)),
        };
        using var session = Open(TrackedChangeMode.RenderInline, content);
        var anchor = session.FindByText("doc")!.Anchor.Id;
        var before = session.GetPackageContentHash();

        var edit = session.ReplaceTextAtSpanWithFormat(anchor, 4, 0, " inserted ", new FormatOp { Bold = true });

        Assert.False(edit.Success);
        Assert.Equal(EditErrorCode.OffsetOutOfRange, edit.Error!.Code);
        Assert.Equal(before, session.GetPackageContentHash());
        Assert.Equal(0, session.Version);
        Assert.Equal(0, session.UndoCount);
    }

    [Fact]
    public void TF802_NvcaFootnote_FormattedInsertionPreservesTabAndReference()
    {
        using var session = new DocxSession(File.ReadAllBytes("../../../../TestFiles/NVCA-Model-COI.docx"),
            new DocxSessionSettings { PersistAnchorIds = true, EmitMarkdownPatch = false });
        var anchor = session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p" && a.Anchor.Scope == "fn"
            && a.TextPreview.StartsWith("Consider adding other exceptions", StringComparison.Ordinal)).Anchor.Id;
        var before = string.Concat(session.GetFormatting(anchor)!.Runs.Select(r => r.Text));
        var original = XElement.Parse(session.Raw.GetXml(anchor));
        var reference = Assert.Single(original.Descendants(W.footnoteRef)).Parent;
        Assert.Single(original.Descendants(W.tab));

        var edit = session.ReplaceTextAtSpanWithFormat(anchor, 24, 0, " inserted ", new FormatOp { Bold = true });

        Assert.True(edit.Success, edit.Error?.Message);
        var after = session.GetFormatting(anchor);
        Assert.NotNull(after);
        Assert.Equal(before.Insert(24, " inserted "), string.Concat(after.Runs.Select(r => r.Text)));
        var xml = XElement.Parse(session.Raw.GetXml(anchor));
        Assert.True(XNode.DeepEquals(reference, Assert.Single(xml.Descendants(W.footnoteRef)).Parent));
        Assert.Equal("Consider adding other ex", Assert.Single(xml.Descendants(W.tab)).ElementsAfterSelf(W.t).Single().Value);
    }

    private static DocxSession Open(TrackedChangeMode tracking, params XElement[] content)
    {
        using var stream = new MemoryStream();
        stream.Write(DocxSessionTests.BuildDocWithoutCodeStyle());
        if (content.Length > 0)
        {
            using var document = WordprocessingDocument.Open(stream, true);
            var paragraph = document.MainDocumentPart!.GetXDocument().Descendants(W.p).Single();
            if (content[0].Name == W.p) paragraph.ReplaceWith(content);
            else paragraph.ReplaceNodes(content);
            document.MainDocumentPart.PutXDocument();
        }
        return new(stream.ToArray(), new DocxSessionSettings
        {
            PersistAnchorIds = true,
            EmitMarkdownPatch = false,
            TrackedChanges = tracking,
            UndoDepth = 1,
        });
    }

    private static XElement FieldRun(string kind) =>
        new(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, kind)));

    private static XElement[] ComplexField(params XElement[] result) =>
        new[] { FieldRun("begin"), new XElement(W.r, new XElement(W.instrText, " DOCPROPERTY Title ")), FieldRun("separate") }
            .Concat(result).Append(FieldRun("end")).ToArray();

    private static T Field<T>(DocxSession session, string name) =>
        (T)typeof(DocxSession).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!
            .GetValue(session)!;
}
