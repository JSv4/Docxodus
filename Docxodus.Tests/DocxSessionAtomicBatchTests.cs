#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Complete-package atomic mutation batch regression coverage (issue #445).</summary>
public class DocxSessionAtomicBatchTests
{
    [Fact]
    public void DS454_AtomicSuccess_IsOneVersionAndOneUndoRedoUnit()
    {
        using var session = Open();
        var paragraphs = BodyParagraphs(session);
        var before = session.Save(persistAnchorIds: false);

        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ExecuteMutation(
                    new MutationPreconditions { ExpectedVersion = 0 },
                    x => x.ReplaceText(paragraphs[0], "Atomic first."))),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ExecuteMutation(
                    new MutationPreconditions { ExpectedVersion = 0 },
                    x => x.ReplaceText(paragraphs[1], "Atomic second."))),
        });

        Assert.True(result.Success);
        Assert.False(result.RolledBack);
        Assert.Equal(1, session.Version);
        Assert.Equal(1, session.UndoCount);
        Assert.Equal(0, session.RedoCount);
        var edited = session.Save(persistAnchorIds: false);
        Assert.Contains("Atomic first.", session.Project().Markdown);
        Assert.Contains("Atomic second.", session.Project().Markdown);

        Assert.True(session.Undo());
        Assert.Equal(2, session.Version);
        AssertSamePackage(before, session.Save(persistAnchorIds: false));
        Assert.True(session.Redo());
        Assert.Equal(3, session.Version);
        AssertSamePackage(edited, session.Save(persistAnchorIds: false));
    }

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("note")]
    [InlineData("comment")]
    [InlineData("table")]
    [InlineData("relationship")]
    [InlineData("annotation")]
    public void DS455_AtomicFailure_RestoresEveryStoryAndRelationship(string mutationKind)
    {
        using var session = Open();
        var body = BodyParagraphs(session)[0];
        var normalBefore = session.Save(persistAnchorIds: false);
        var persistedBefore = session.Save(persistAnchorIds: true);
        var anchorsBefore = session.Project().AnchorIndex.Keys.OrderBy(x => x).ToArray();
        var hyperlinkCountBefore = session.LiveDocument.MainDocumentPart!.HyperlinkRelationships.Count();
        var hyperlinkCountDuring = hyperlinkCountBefore;

        EditResult Mutate(DocxSession s) => mutationKind switch
        {
            "body" => s.ReplaceText(body, "Changed body."),
            "header" => s.SetHeaderText(body, HeaderFooterKind.Default, "Atomic header."),
            "footer" => s.SetFooterText(body, HeaderFooterKind.Default, "Atomic footer."),
            "note" => s.InsertFootnote(body, 3, "Atomic note."),
            "comment" => s.AddComment(body, null, "Alice", "Atomic comment."),
            "table" => s.InsertTable(body, Position.After, 2, 2),
            "relationship" => AddHyperlink(s),
            "annotation" => s.AddAnnotation(body, new CharSpan(0, 5), new DocumentAnnotation
            {
                Id = "atomic-annotation",
                LabelId = "RISK",
                Label = "Risk",
                Color = "#FFCC00",
            }),
            _ => throw new ArgumentOutOfRangeException(nameof(mutationKind)),
        };

        EditResult AddHyperlink(DocxSession s)
        {
            var edit = s.ReplaceText(body, "Changed [link](https://example.test/atomic-batch)");
            hyperlinkCountDuring = s.LiveDocument.MainDocumentPart!.HyperlinkRelationships.Count();
            return edit;
        }

        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", mutationKind, Mutate),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText("p:body:missing", "must fail")),
        });

        Assert.False(result.Success);
        Assert.True(result.RolledBack);
        Assert.Equal(1, result.Failure?.Index);
        Assert.Equal("docx_edit", result.Failure?.Tool);
        Assert.Equal("replace_text", result.Failure?.Action);
        Assert.Equal(EditErrorCode.AnchorNotFound, result.Failure?.Error.Code);
        Assert.True(result.Failure?.RolledBack);
        Assert.All(result.Steps, step => Assert.True(step.RolledBack));
        Assert.Equal(0, session.Version);
        Assert.Equal(0, session.UndoCount);
        Assert.Equal(0, session.RedoCount);
        AssertSamePackage(normalBefore, session.Save(persistAnchorIds: false));
        AssertSamePackage(persistedBefore, session.Save(persistAnchorIds: true));
        Assert.Equal(anchorsBefore, session.Project().AnchorIndex.Keys.OrderBy(x => x).ToArray());
        Assert.Equal(hyperlinkCountBefore, session.LiveDocument.MainDocumentPart!.HyperlinkRelationships.Count());
        Assert.Empty(session.ListAnnotations());
        if (mutationKind == "relationship")
            Assert.True(hyperlinkCountDuring > hyperlinkCountBefore);
    }

    [Fact]
    public void DS455B_ThrowingStep_AfterCrossPartMutationIsStructuredAndRolledBack()
    {
        using var session = Open();
        var body = BodyParagraphs(session)[0];
        var before = session.Save(persistAnchorIds: true);

        var result = session.ExecuteBatch(new MutationBatchStep[]
        {
            new("docx_create", "set_header_text",
                s => s.SetHeaderText(body, HeaderFooterKind.Default, "Speculative header.")),
            new("docx_edit", "throw",
                (Func<DocxSession, EditResult>)(_ => throw new InvalidOperationException("deliberate batch fault"))),
        });

        Assert.False(result.Success);
        Assert.True(result.RolledBack);
        Assert.Equal(1, result.Failure?.Index);
        Assert.Equal("docx_edit", result.Failure?.Tool);
        Assert.Equal("throw", result.Failure?.Action);
        Assert.Equal(EditErrorCode.InternalError, result.Failure?.Error.Code);
        Assert.Contains("deliberate batch fault", result.Failure?.Error.Message);
        Assert.Equal(0, session.Version);
        Assert.Equal(0, session.UndoCount);
        AssertSamePackage(before, session.Save(persistAnchorIds: true));
        Assert.DoesNotContain("Speculative header.", session.Project().Markdown);
    }

    [Fact]
    public void DS455C_AtomicFailure_PreservesOpaqueCustomXmlPayloadAndRelationship()
    {
        var payload = Encoding.UTF8.GetBytes(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?><vendor:data xmlns:vendor=\"urn:vendor:test\">\n  <vendor:item id=\"7\"> untouched </vendor:item>\n</vendor:data>");
        var seeded = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var source = new MemoryStream();
        source.Write(seeded);
        source.Position = 0;
        string partUri;
        string relationshipId;
        using (var package = WordprocessingDocument.Open(source, isEditable: true))
        {
            var custom = package.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            custom.FeedData(new MemoryStream(payload));
            partUri = custom.Uri.ToString();
            relationshipId = package.MainDocumentPart.GetIdOfPart(custom);
            package.Save();
        }

        using var session = new DocxSession(source.ToArray(), new DocxSessionSettings
        {
            CaptureInitialProjection = false,
            PersistAnchorIds = false,
        });
        var anchor = BodyParagraphs(session)[0];
        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchor, "Speculative body edit.")),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText("p:body:missing", "failure")),
        });

        Assert.False(result.Success);
        Assert.True(result.RolledBack);
        var restored = session.LiveDocument.MainDocumentPart!.CustomXmlParts
            .Single(part => part.Uri.ToString() == partUri);
        Assert.Equal(relationshipId, session.LiveDocument.MainDocumentPart.GetIdOfPart(restored));
        Assert.Equal(payload, ReadPartBytes(restored));

        // A normal output save may inspect custom XML while finding Docxodus annotations, but
        // opaque application payloads must remain byte-for-byte untouched.
        var saved = session.Save(persistAnchorIds: false);
        using var reopened = WordprocessingDocument.Open(new MemoryStream(saved), isEditable: false);
        var savedPart = reopened.MainDocumentPart!.CustomXmlParts
            .Single(part => part.Uri.ToString() == partUri);
        Assert.Equal(relationshipId, reopened.MainDocumentPart.GetIdOfPart(savedPart));
        Assert.Equal(payload, ReadPartBytes(savedPart));
    }

    [Fact]
    public void DS456_AtomicFailure_RestoresRedoCursorAndHistoryDiagnostics()
    {
        using var session = Open(new DocxSessionSettings
        {
            PersistAnchorIds = true,
            UndoDepth = 1,
        });
        var body = BodyParagraphs(session)[0];
        Assert.True(session.ReplaceText(body, "History edit.").Success);
        Assert.True(session.Undo());
        Assert.Equal(0, session.UndoCount);
        Assert.Equal(1, session.RedoCount);
        var trimmedBefore = session.UndoHistoryTrimmedForMemory;
        var memoryBefore = session.UndoMemoryBytes;
        var versionBefore = session.Version;

        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text", s => s.ReplaceText(body, "Speculative.")),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText("p:body:missing", "failure")),
        });

        Assert.False(result.Success);
        Assert.Equal(versionBefore, session.Version);
        Assert.Equal(0, session.UndoCount);
        Assert.Equal(1, session.RedoCount);
        Assert.Equal(trimmedBefore, session.UndoHistoryTrimmedForMemory);
        Assert.Equal(memoryBefore, session.UndoMemoryBytes);
        Assert.False(session.Undo());
        Assert.True(session.Redo());
        Assert.Contains("History edit.", session.Project().Markdown);
    }

    [Fact]
    public void DS457_Transactions_AreNestedLifoAndOuterCommitSquashesInnerWork()
    {
        using var session = Open();
        var paragraphs = BodyParagraphs(session);
        var before = session.Save(persistAnchorIds: false);

        using (var outer = session.BeginTransaction())
        {
            Assert.True(session.ReplaceText(paragraphs[0], "Outer edit.").Success);
            using (var inner = session.BeginTransaction())
            {
                Assert.True(session.SetHeaderText(
                    paragraphs[0], HeaderFooterKind.Default, "Rolled-back header.").Success);
                inner.Rollback();
            }

            Assert.DoesNotContain("Rolled-back header.", session.Project().Markdown);
            using (var inner = session.BeginTransaction())
            {
                Assert.True(session.ReplaceText(paragraphs[1], "Inner committed edit.").Success);
                inner.Commit();
            }

            Assert.Equal(0, session.Version);
            Assert.False(session.Undo());
            outer.Commit();
        }

        Assert.Equal(1, session.Version);
        Assert.Equal(1, session.UndoCount);
        Assert.Contains("Outer edit.", session.Project().Markdown);
        Assert.Contains("Inner committed edit.", session.Project().Markdown);
        Assert.True(session.Undo());
        AssertSamePackage(before, session.Save(persistAnchorIds: false));
    }

    [Fact]
    public void DS457B_OutOfOrderCompletionLeavesBothScopesRecoverable()
    {
        using var session = Open();
        var anchor = BodyParagraphs(session)[0];
        var before = session.Save(persistAnchorIds: true);
        var outer = session.BeginTransaction();
        Assert.True(session.ReplaceText(anchor, "Outer speculative edit.").Success);
        var inner = session.BeginTransaction();
        Assert.True(session.SetHeaderText(
            anchor, HeaderFooterKind.Default, "Inner speculative header.").Success);

        Assert.Throws<InvalidOperationException>(() => outer.Commit());
        Assert.False(outer.IsCompleted);
        Assert.False(inner.IsCompleted);

        inner.Commit();
        Assert.True(inner.IsCompleted);
        Assert.False(outer.IsCompleted);
        outer.Rollback();

        Assert.True(outer.IsCompleted);
        Assert.Equal(0, session.Version);
        Assert.Equal(0, session.UndoCount);
        AssertSamePackage(before, session.Save(persistAnchorIds: true));
    }

    [Fact]
    public void DS457C_WrongThreadCompletionLeavesScopeRecoverableByOwner()
    {
        using var session = Open();
        var before = session.Save(persistAnchorIds: false);
        var transaction = session.BeginTransaction();
        Assert.True(session.ReplaceText(
            BodyParagraphs(session)[0], "Wrong-thread speculative edit.").Success);

        Exception? failure = null;
        var wrongThread = new Thread(() => failure = Record.Exception(transaction.Rollback));
        wrongThread.Start();
        wrongThread.Join();
        Assert.IsType<InvalidOperationException>(failure);
        Assert.False(transaction.IsCompleted);

        transaction.Rollback();
        Assert.True(transaction.IsCompleted);
        AssertSamePackage(before, session.Save(persistAnchorIds: false));
        Assert.Equal(0, session.Version);
        Assert.Equal(0, session.UndoCount);
    }

    [Fact]
    public void DS457D_DisposingSessionAbandonsScopesWithoutHoldingGateOrResurrectingPackage()
    {
        var session = Open();
        var mutationGate = PrivateField<object>(session, "_mutationGate");
        var transaction = session.BeginTransaction();
        Assert.True(session.ReplaceText(
            BodyParagraphs(session)[0], "Abandoned speculative edit.").Success);

        session.Dispose();

        Assert.True(transaction.IsCompleted);
        transaction.Dispose();
        Assert.Throws<ObjectDisposedException>(() => session.Project());
        Assert.True(Task.Run(() =>
        {
            if (!Monitor.TryEnter(mutationGate, TimeSpan.FromSeconds(1))) return false;
            Monitor.Exit(mutationGate);
            return true;
        }).GetAwaiter().GetResult());
    }

    [Fact]
    public void DS457E_WrongThreadSessionDisposeLeavesOwnerAbleToRollback()
    {
        var session = Open();
        var transaction = session.BeginTransaction();

        Exception? failure = null;
        var wrongThread = new Thread(() => failure = Record.Exception(session.Dispose));
        wrongThread.Start();
        wrongThread.Join();
        Assert.IsType<InvalidOperationException>(failure);
        Assert.False(transaction.IsCompleted);
        Assert.NotEmpty(session.Project().Markdown);

        transaction.Rollback();
        session.Dispose();
    }

    [Fact]
    public void DS458_NoOpAndPreflightFailure_AreBytePureAndDoNotReassignAnchors()
    {
        using var session = Open(new DocxSessionSettings
        {
            CaptureInitialProjection = false,
            PersistAnchorIds = false,
        });
        var anchors = session.Project().AnchorIndex.Keys.OrderBy(x => x).ToArray();
        var normalBefore = session.Save(persistAnchorIds: false);

        using (var transaction = session.BeginTransaction())
            transaction.Commit();

        AssertSamePackage(normalBefore, session.Save(persistAnchorIds: false));
        Assert.Equal(anchors, session.Project().AnchorIndex.Keys.OrderBy(x => x).ToArray());
        Assert.Equal(0, session.Version);
        Assert.Equal(0, session.UndoCount);

        var persistedBefore = session.Save(persistAnchorIds: true);
        var invoked = false;
        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep(
                "docx_edit", "replace_text",
                s => { invoked = true; return s.ReplaceText(anchors[0], "not run"); },
                _ => new EditError(EditErrorCode.PreconditionFailed, "preflight rejected")),
        });

        Assert.False(result.Success);
        Assert.True(result.RolledBack);
        Assert.False(invoked);
        Assert.Equal(0, session.Version);
        Assert.Equal(0, session.UndoCount);
        AssertSamePackage(normalBefore, session.Save(persistAnchorIds: false));
        AssertSamePackage(persistedBefore, session.Save(persistAnchorIds: true));
        Assert.Equal(anchors, session.Project().AnchorIndex.Keys.OrderBy(x => x).ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DS458B_DisposedNoOpTransaction_WithWarmedReadCachesIsStrictlyBytePure(
        bool persistAnchorIds)
    {
        using var session = Open(new DocxSessionSettings
        {
            CaptureInitialProjection = false,
            PersistAnchorIds = persistAnchorIds,
        });

        // Warm the XDocument-backed projection, story, relationship and custom-XML reads before
        // checkpointing. A no-op rollback must not serialize or otherwise perturb those caches.
        _ = session.Project();
        _ = session.ListBlocks();
        _ = session.ListNotes();
        _ = session.ListComments();
        _ = session.ListRevisions();
        _ = session.ListAnnotations();
        var before = session.Save(persistAnchorIds);

        using (session.BeginTransaction())
        {
        }

        var after = session.Save(persistAnchorIds);
        Assert.Equal(before, after);
        Assert.Equal(0, session.Version);
        Assert.Equal(0, session.UndoCount);
        Assert.Equal(0, session.RedoCount);
    }

    [Fact]
    public void DS458C_SaveFalse_PreservesPtDeclarationNamedByMcIgnorable()
    {
        XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        var seeded = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var source = new MemoryStream();
        source.Write(seeded);
        source.Position = 0;
        using (var package = WordprocessingDocument.Open(source, isEditable: true))
        {
            var root = package.MainDocumentPart!.GetXDocument().Root!;
            root.SetAttributeValue(XNamespace.Xmlns + "mc", mc.NamespaceName);
            root.SetAttributeValue(XNamespace.Xmlns + "pt", PtOpenXml.pt.NamespaceName);
            root.SetAttributeValue(mc + "Ignorable", "pt");
            root.Descendants().First().SetAttributeValue(PtOpenXml.Unid, "save-false-fixture");
            package.MainDocumentPart.PutXDocument();
            package.Save();
        }

        using var session = new DocxSession(source.ToArray(), new DocxSessionSettings
        {
            CaptureInitialProjection = false,
            PersistAnchorIds = false,
        });
        var saved = session.Save(persistAnchorIds: false);

        using var reopened = WordprocessingDocument.Open(new MemoryStream(saved), isEditable: false);
        // Access the typed DOM to force the SDK's markup-compatibility prefix resolution.
        Assert.NotNull(reopened.MainDocumentPart!.Document.Body);
        var savedRoot = reopened.MainDocumentPart.GetXDocument().Root!;
        Assert.Equal(PtOpenXml.pt.NamespaceName,
            (string?)savedRoot.Attribute(XNamespace.Xmlns + "pt"));
        Assert.Contains("pt", ((string?)savedRoot.Attribute(mc + "Ignorable") ?? string.Empty)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        Assert.DoesNotContain(savedRoot.DescendantsAndSelf().Attributes(),
            attribute => attribute.Name == PtOpenXml.Unid);
    }

    [Fact]
    public void DS459_Rollback_RestoresTrackingConfigurationAndRevisionGenerators()
    {
        using var seed = Open();
        var baseBytes = seed.Save(persistAnchorIds: true);
        var settings = new DocxSessionSettings
        {
            PersistAnchorIds = true,
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Alice",
        };
        using var tested = new DocxSession(baseBytes, settings);
        using var control = new DocxSession(baseBytes, settings);
        var anchor = BodyParagraphs(tested)[0];
        var controlAnchor = BodyParagraphs(control)[0];
        var revisionCounter = PrivateField<int>(tested, "_revisionCounter");
        var formatTicks = PrivateField<long>(tested, "_lastFormatRevisionTicks");

        var result = tested.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchor, "Failed tracked edit.")),
            new MutationBatchStep("docx_edit", "apply_format",
                s => s.ApplyFormat(anchor, new CharSpan(0, 5), new FormatOp { Bold = true })),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText("p:body:missing", "failure")),
        });

        Assert.False(result.Success);
        Assert.Equal(revisionCounter, PrivateField<int>(tested, "_revisionCounter"));
        Assert.Equal(formatTicks, PrivateField<long>(tested, "_lastFormatRevisionTicks"));
        Assert.Equal(TrackedChangeMode.RenderInline, tested.TrackedChanges);
        Assert.Equal("Alice", tested.RevisionAuthor);
        Assert.Empty(tested.ListRevisions());

        Assert.True(tested.ReplaceText(anchor, "Committed tracked edit.").Success);
        Assert.True(control.ReplaceText(controlAnchor, "Committed tracked edit.").Success);
        Assert.Equal(
            control.ListRevisions().Select(r => (r.Id, r.Type, r.Author, r.Text)),
            tested.ListRevisions().Select(r => (r.Id, r.Type, r.Author, r.Text)));

        using (var transaction = tested.BeginTransaction())
        {
            tested.SetTrackedChanges(TrackedChangeMode.Accept);
            tested.SetRevisionAuthor("Speculative");
            transaction.Rollback();
        }
        Assert.Equal(TrackedChangeMode.RenderInline, tested.TrackedChanges);
        Assert.Equal("Alice", tested.RevisionAuthor);
    }

    [Fact]
    public void DS460_BestEffort_IsExplicitAndRetainsPartialSuccess()
    {
        using var session = Open();
        var anchor = BodyParagraphs(session)[0];

        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchor, "Retained partial edit.")),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText("p:body:missing", "failure")),
        }, MutationBatchMode.BestEffort);

        Assert.False(result.Success);
        Assert.False(result.RolledBack);
        Assert.False(result.Failure?.RolledBack);
        Assert.Contains("Retained partial edit.", session.Project().Markdown);
        Assert.Equal(1, session.Version);
        Assert.Equal(1, session.UndoCount);
    }

    [Fact]
    public void DS460B_BestEffortPreflightsEachStepAgainstSequentialState()
    {
        using var session = Open();
        var paragraphs = BodyParagraphs(session);

        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(paragraphs[0], "State created by step zero.")),
            new MutationBatchStep(
                "docx_edit", "replace_text",
                s => s.ReplaceText(paragraphs[1], "Preflight observed the new state."),
                s => s.Project().Markdown.Contains("State created by step zero.", StringComparison.Ordinal)
                    ? null
                    : new EditError(EditErrorCode.PreconditionFailed,
                        "step zero's state was not visible")),
        }, MutationBatchMode.BestEffort);

        Assert.True(result.Success);
        Assert.Equal(2, result.Steps.Count);
        Assert.All(result.Steps, step => Assert.True(step.Success));
        Assert.Contains("State created by step zero.", session.Project().Markdown);
        Assert.Contains("Preflight observed the new state.", session.Project().Markdown);
        Assert.Equal(2, session.Version);
        Assert.Equal(2, session.UndoCount);
    }

    private static T PrivateField<T>(DocxSession session, string name) =>
        (T)typeof(DocxSession).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!
            .GetValue(session)!;

    private static void AssertSamePackage(byte[] expected, byte[] actual) =>
        PackageEquivalence.AssertSamePackage(
            new WmlDocument("expected.docx", expected),
            new WmlDocument("actual.docx", actual));

    private static byte[] ReadPartBytes(OpenXmlPart part)
    {
        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        source.CopyTo(copy);
        return copy.ToArray();
    }

    private static DocxSession Open(DocxSessionSettings? settings = null) =>
        new(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            settings ?? new DocxSessionSettings { PersistAnchorIds = true });

    private static string[] BodyParagraphs(DocxSession session) =>
        session.Project().AnchorIndex.Keys
            .Where(id => id.StartsWith("p:body:", StringComparison.Ordinal))
            .ToArray();
}
