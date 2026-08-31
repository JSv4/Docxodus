#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Complete-package isolated preview regression coverage (issue #446).</summary>
public class DocxSessionPreviewBatchTests
{
    [Fact]
    public void DS461_PreviewSuccessFailureThrowAndBestEffort_NeverTouchLiveState()
    {
        using var session = OpenRich(new DocxSessionSettings
        {
            PersistAnchorIds = true,
            UndoDepth = 1,
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Preview Author",
        });
        var anchors = BodyParagraphs(session);

        // Seed a redo cursor and a tight history ring: the former apply-and-undo implementation
        // destroyed redo here and could underflow after more preview steps than UndoDepth.
        Assert.True(session.ReplaceText(anchors[0], "Redo target.").Success);
        Assert.True(session.Undo());
        _ = session.Project();
        _ = session.AnchorIndex();
        var before = Fingerprint.Capture(session);

        var success = session.PreviewBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchors[0], "Predicted tracked replacement.")),
            new MutationBatchStep("docx_create", "set_header_text",
                s => s.SetHeaderText(anchors[0], HeaderFooterKind.Default, "Predicted header.")),
            new MutationBatchStep("docx_comment", "add",
                s => s.AddComment(anchors[1], null, "Alice", "Predicted comment.",
                    date: new DateTime(2025, 1, 2, 3, 4, 5, DateTimeKind.Utc))),
            new MutationBatchStep("docx_annotate", "add",
                s => s.AddAnnotation(anchors[1], new CharSpan(0, 3), new DocumentAnnotation
                {
                    Id = "preview-ann",
                    LabelId = "RISK",
                    Label = "Risk",
                    Color = "#FFCC00",
                    Created = new DateTime(2025, 1, 2, 3, 4, 5, DateTimeKind.Utc),
                })),
        }, options: new MutationBatchPreviewOptions
        {
            HtmlMode = MutationPreviewHtmlMode.Full,
        });

        Assert.True(success.Preview);
        Assert.True(success.Success,
            success.Failure is null
                ? "preview failed without a failure envelope"
                : $"{success.Failure.Index}:{success.Failure.Action}:{success.Failure.Error.Code}:{success.Failure.Error.Message}");
        Assert.Equal(before.Version, success.BaseVersion);
        Assert.Equal(before.Version + 1, success.ResultVersion);
        Assert.NotNull(success.PackageHash);
        Assert.NotEmpty(success.PackageHash);
        Assert.Equal(4, success.Steps.Count);
        Assert.NotEmpty(success.RevisionChanges.Added);
        Assert.Single(success.CommentChanges.Added);
        Assert.Single(success.AnnotationChanges.Added);
        Assert.Contains(success.Warnings,
            warning => warning.Contains("Comment date attributes", StringComparison.Ordinal));
        Assert.Contains("Predicted tracked replacement.", success.Html);
        before.AssertUnchanged(session);

        var failed = session.PreviewBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchors[0], "Rolled back in the shadow.")),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText("p:body:missing", "failure")),
        });
        Assert.False(failed.Success);
        Assert.True(failed.RolledBack);
        Assert.Empty(failed.RevisionChanges.Added);
        before.AssertUnchanged(session);

        var thrown = session.PreviewBatch(new MutationBatchStep[]
        {
            new("docx_edit", "replace_text",
                s => s.ReplaceText(anchors[0], "Thrown away in shadow.")),
            new("docx_edit", "throw",
                (Func<DocxSession, EditResult>)(_ => throw new InvalidOperationException("preview fault"))),
        });
        Assert.False(thrown.Success);
        Assert.Equal(EditErrorCode.InternalError, thrown.Failure?.Error.Code);
        before.AssertUnchanged(session);

        var partial = session.PreviewBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchors[0], "Retained only in best-effort shadow.")),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText("p:body:missing", "failure")),
            new MutationBatchStep("docx_create", "set_footer_text",
                s => s.SetFooterText(anchors[1], HeaderFooterKind.Default, "Shadow footer.")),
        }, MutationBatchMode.BestEffort);
        Assert.False(partial.Success);
        Assert.False(partial.RolledBack);
        Assert.Equal(before.Version + 2, partial.ResultVersion);
        Assert.Contains(partial.Warnings, value => value.Contains("Best-effort", StringComparison.Ordinal));
        before.AssertUnchanged(session);

        // The original redo remains usable after every preview, including batches longer than
        // UndoDepth. This explicitly supersedes the undo-too-many failure mode from #468.
        Assert.False(session.Undo());
        Assert.True(session.Redo());
        Assert.Contains("Redo target.", session.Project().Markdown);
    }

    [Fact]
    public void DS462_DisposedOrAbandonedShadow_IsIntrinsicallySafe()
    {
        using var live = OpenRich();
        var anchor = BodyParagraphs(live)[0];
        var before = Fingerprint.Capture(live);

        var shadow = live.CreateShadowSession();
        var liveSettings = PrivateField<DocxSessionSettings>(live, "_settings");
        var shadowSettings = PrivateField<DocxSessionSettings>(shadow, "_settings");
        Assert.NotSame(liveSettings, shadowSettings);
        Assert.NotSame(liveSettings.ProjectionSettings, shadowSettings.ProjectionSettings);
        shadowSettings.ProjectionSettings.HeadingLevelOffset++;
        Assert.True(shadow.ReplaceText(anchor, "Only the abandoned clone changes.").Success);
        Assert.Contains("Only the abandoned clone changes.", shadow.Project().Markdown);
        before.AssertUnchanged(live); // live is safe even while the shadow is still in flight
        shadow.Dispose();
        before.AssertUnchanged(live);

        // Timeout-style abandonment: work can fault/dispose independently because no rollback of
        // live state is ever needed.
        var task = Task.Run(() =>
        {
            using var timedOutShadow = live.CreateShadowSession();
            Assert.True(timedOutShadow.SetHeaderText(
                anchor, HeaderFooterKind.Default, "Timed-out shadow.").Success);
            throw new TimeoutException("simulated caller abandonment");
        });
        Assert.IsType<TimeoutException>(Record.Exception(() => task.GetAwaiter().GetResult()));
        before.AssertUnchanged(live);
    }

    [Fact]
    public void DS463_DeterministicPreviewAndApply_HaveIdenticalReceiptsAndPackageHash()
    {
        using var session = OpenRich();
        var anchors = BodyParagraphs(session);
        Assert.True(session.ReplaceText(anchors[0], "Existing live change from initial baseline.").Success);
        var expectedDiff = session.GetDiff();
        var expectedTransactionState = PrivateField<long>(session, "_nextTransactionId");
        string? previewDiff = null;
        long previewPreflightTransaction = -1;
        long previewMutationTransaction = -1;
        var previewSteps = new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s =>
                {
                    previewMutationTransaction = PrivateField<long>(s, "_nextTransactionId");
                    return s.ReplaceText(anchors[0], "Deterministic replacement.");
                },
                s =>
                {
                    previewDiff = s.GetDiff();
                    previewPreflightTransaction = PrivateField<long>(s, "_nextTransactionId");
                    return null;
                }),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchors[1], "Deterministic second replacement.")),
        };

        var preview = session.PreviewBatch(previewSteps);
        Assert.Equal(expectedDiff, previewDiff);
        Assert.Equal(expectedTransactionState + 1, previewPreflightTransaction);
        Assert.Equal(expectedTransactionState + 1, previewMutationTransaction);
        Assert.Equal(1, session.Version);

        string? applyDiff = null;
        long applyPreflightTransaction = -1;
        long applyMutationTransaction = -1;
        var applySteps = new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s =>
                {
                    applyMutationTransaction = PrivateField<long>(s, "_nextTransactionId");
                    return s.ReplaceText(anchors[0], "Deterministic replacement.");
                },
                s =>
                {
                    applyDiff = s.GetDiff();
                    applyPreflightTransaction = PrivateField<long>(s, "_nextTransactionId");
                    return null;
                }),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchors[1], "Deterministic second replacement.")),
        };
        var applied = session.ExecuteBatch(applySteps);

        Assert.Equal(previewDiff, applyDiff);
        Assert.Equal(previewPreflightTransaction, applyPreflightTransaction);
        Assert.Equal(previewMutationTransaction, applyMutationTransaction);

        Assert.True(preview.Preview);
        Assert.False(applied.Preview);
        Assert.Equal(preview.BaseVersion, applied.BaseVersion);
        Assert.Equal(preview.ResultVersion, applied.ResultVersion);
        Assert.Equal(preview.PackageHash, applied.PackageHash);
        Assert.Equal(
            preview.Steps.Select(Receipt),
            applied.Steps.Select(Receipt));
        Assert.Equal(ChangeReceipt(preview.RevisionChanges), ChangeReceipt(applied.RevisionChanges));
        Assert.Equal(ChangeReceipt(preview.CommentChanges), ChangeReceipt(applied.CommentChanges));
        Assert.Equal(ChangeReceipt(preview.AnnotationChanges), ChangeReceipt(applied.AnnotationChanges));

        static string Receipt(MutationBatchStepResult step) =>
            $"{step.Index}|{step.Tool}|{step.Action}|{step.Success}|" +
            string.Join(";", step.Results.Select(result =>
                $"{result.Success}:{string.Join(',', result.Created.Select(a => a.Id))}:" +
                $"{string.Join(',', result.Removed.Select(a => a.Id))}:" +
                $"{string.Join(',', result.Modified.Select(a => a.Id))}"));

        static string ChangeReceipt<T>(MutationBatchChangeSet<T> changes) =>
            $"{changes.Added.Count}|{changes.Removed.Count}|{changes.Modified.Count}";
    }

    [Fact]
    public void DS464_HandlePreviewFactory_CannotAccidentallyTargetTheLiveHandle()
    {
        var handle = DocxSessionOps.OpenSession(RichBytes(), new DocxSessionSettings
        {
            PersistAnchorIds = true,
            UndoDepth = 1,
        });
        try
        {
            using var projection = System.Text.Json.JsonDocument.Parse(DocxSessionOps.Project(handle));
            var anchor = projection.RootElement.GetProperty("anchorIndex")
                .EnumerateObject().First(property => property.Name.StartsWith("p:body:", StringComparison.Ordinal)).Name;
            _ = DocxSessionOps.Save(handle, persistAnchorIds: false);
            _ = DocxSessionOps.Save(handle, persistAnchorIds: true);
            var beforeNormal = DocxSessionOps.Save(handle, persistAnchorIds: false);
            var beforePersisted = DocxSessionOps.Save(handle, persistAnchorIds: true);
            var beforeVersion = DocxSessionOps.GetVersion(handle);

            var json = DocxSessionOps.PreviewBatch(
                handle,
                MutationBatchMode.Atomic,
                shadowHandle => new[]
                {
                    DocxSessionOps.SerializedBatchStep(
                        "docx_scalpel",
                        "replace_text",
                        () => DocxSessionOps.ReplaceText(
                            shadowHandle, anchor, "Handle-only predicted edit.")),
                });

            using var result = System.Text.Json.JsonDocument.Parse(json);
            Assert.True(result.RootElement.GetProperty("preview").GetBoolean());
            Assert.True(result.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(beforeVersion, DocxSessionOps.GetVersion(handle));
            var afterNormal = DocxSessionOps.Save(handle, persistAnchorIds: false);
            var afterPersisted = DocxSessionOps.Save(handle, persistAnchorIds: true);
            Assert.Equal(beforeNormal, afterNormal);
            Assert.Equal(beforePersisted, afterPersisted);
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void DS465_PostCommitInspectionFailure_IsWarningNotApparentMutationFailure()
    {
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var package = WordprocessingDocument.Open(stream, isEditable: true))
        {
            const string paraId = "A1B2C3D4";
            var main = package.MainDocumentPart!;
            var comments = main.AddNewPart<WordprocessingCommentsPart>();
            comments.PutXDocument(new XDocument(
                new XElement(W.comments,
                    new XElement(W.comment,
                        new XAttribute(W.id, "1"),
                        new XAttribute(W.author, "Observer"),
                        new XElement(W.p,
                            new XAttribute(W14.paraId, paraId),
                            new XElement(W.r, new XElement(W.t, "comment")))))));
            package.Save();
        }

        using var session = new DocxSession(stream.ToArray());
        Assert.Single(session.ListComments());
        var anchor = BodyParagraphs(session)[0];
        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s =>
                {
                    var edit = s.ReplaceText(anchor, "The mutation still commits.");
                    if (!edit.Success) return edit;

                    // Simulate a failure in optional receipt enrichment only after the mutation
                    // has committed its ordinary operation state.
                    var document = PrivateField<WordprocessingDocument>(s, "_doc");
                    var commentsEx = document.MainDocumentPart!
                        .AddNewPart<WordprocessingCommentsExPart>();
                    commentsEx.FeedData(new MemoryStream(Encoding.UTF8.GetBytes("<malformed")));
                    return edit;
                }),
        });

        Assert.True(result.Success,
            result.Failure is null
                ? string.Join("; ", result.Warnings)
                : $"{result.Failure.Error.Code}: {result.Failure.Error.Message}");
        Assert.Equal(1, session.Version);
        Assert.Contains("The mutation still commits.", session.Project().Markdown);
        Assert.Contains(result.Warnings,
            value => value.Contains("Comment delta inspection unavailable", StringComparison.Ordinal));
        Assert.Empty(result.CommentChanges.Added);
    }

    [Fact]
    public void DS466_InvalidPreviewHtmlMode_IsRejectedBeforeShadowExecution()
    {
        using var session = OpenRich();
        var invoked = false;
        Assert.Throws<ArgumentOutOfRangeException>(() => session.PreviewBatch(new[]
        {
            new MutationBatchStep("docx_edit", "never",
                s => { invoked = true; return s.ReplaceText(BodyParagraphs(s)[0], "not run"); }),
        }, options: new MutationBatchPreviewOptions
        {
            HtmlMode = (MutationPreviewHtmlMode)12345,
        }));
        Assert.False(invoked);
        Assert.Equal(0, session.Version);
    }

    [Fact]
    public void DS467_CreatePreviewApply_AreSemanticallyEquivalentModuloGeneratedIds()
    {
        using var session = OpenRich();
        var anchor = BodyParagraphs(session)[0];
        var steps = new[]
        {
            new MutationBatchStep("docx_create", "insert_paragraph",
                s => s.InsertParagraph(anchor, Position.After, "Generated-id paragraph.")),
        };

        var preview = session.PreviewBatch(steps, options: new MutationBatchPreviewOptions
        {
            HtmlMode = MutationPreviewHtmlMode.Full,
        });
        var applied = session.ExecuteBatch(steps);
        var appliedHtml = HtmlConversionOps.ConvertToHtml(session, new HtmlConversionOptions
        {
            CommentRenderMode = 0,
            RenderAnnotations = true,
            RenderFootnotesAndEndnotes = true,
            RenderHeadersAndFooters = true,
            RenderTrackedChanges = true,
            StampAnchors = true,
        });

        var previewCreated = Assert.Single(Assert.Single(preview.Steps).Results).Created;
        var appliedCreated = Assert.Single(Assert.Single(applied.Steps).Results).Created;
        Assert.Equal(previewCreated.Select(anchor => (anchor.Kind, anchor.Scope)),
            appliedCreated.Select(anchor => (anchor.Kind, anchor.Scope)));
        Assert.NotEqual(previewCreated.Select(anchor => anchor.Id), appliedCreated.Select(anchor => anchor.Id));
        Assert.NotEqual(preview.PackageHash, applied.PackageHash);
        Assert.Contains(preview.Warnings,
            warning => warning.Contains("equivalence is semantic", StringComparison.Ordinal));
        Assert.Contains(applied.Warnings,
            warning => warning.Contains("equivalence is semantic", StringComparison.Ordinal));
        Assert.Equal(NormalizeGeneratedIds(preview.Html!), NormalizeGeneratedIds(appliedHtml));
        Assert.Contains("Generated-id paragraph.", appliedHtml);
    }

    [Fact]
    public void DS468_ReceiptEnrichment_IsSerializedWithConcurrentMutations()
    {
        using var session = OpenRich();
        var anchors = BodyParagraphs(session);
        using var receiptInspectionEntered = new ManualResetEventSlim();
        using var releaseReceiptInspection = new ManualResetEventSlim();
        using var concurrentMutationStarted = new ManualResetEventSlim();
        var blockingCreated = new BlockingAnchorList(
            receiptInspectionEntered, releaseReceiptInspection);

        var batchTask = Task.Run(() => session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s =>
                {
                    var edit = s.ReplaceText(anchors[0], "Batch mutation.");
                    return new EditResult
                    {
                        Success = edit.Success,
                        Error = edit.Error,
                        Created = blockingCreated,
                        Removed = edit.Removed,
                        Modified = edit.Modified,
                        Patch = edit.Patch,
                        AnnotationId = edit.AnnotationId,
                    };
                }),
        }));

        Task<EditResult>? concurrentTask = null;
        try
        {
            // Reaching the hook needs a thread-pool thread to pick up the Task.Run and complete
            // a full mutation; on a loaded 2-core CI runner with the suite's heavier comparison
            // collections executing in parallel, that has been observed to exceed 10 s (issue
            // #635). The window is not part of the claim — the claim is ORDERING, asserted by the
            // Assert.False below — so it is generous: the event either fires or the batch is
            // genuinely wedged, and a wedged batch fails at any timeout.
            Assert.True(receiptInspectionEntered.Wait(TimeSpan.FromSeconds(60)),
                "batch did not reach receipt enrichment");
            concurrentTask = Task.Run(() =>
            {
                concurrentMutationStarted.Set();
                return session.ExecuteMutation(
                    preconditions: null,
                    s => s.ReplaceText(anchors[1], "Concurrent mutation."));
            });
            Assert.True(concurrentMutationStarted.Wait(TimeSpan.FromSeconds(60)),
                "concurrent mutation task did not start");
            Assert.False(concurrentTask.Wait(TimeSpan.FromSeconds(1)),
                "concurrent mutation interleaved with batch receipt enrichment");
        }
        finally
        {
            releaseReceiptInspection.Set();
        }

        var batch = batchTask.GetAwaiter().GetResult();
        var concurrent = concurrentTask!.GetAwaiter().GetResult();
        Assert.True(batch.Success);
        Assert.True(concurrent.Success);
        Assert.Equal(0, batch.BaseVersion);
        Assert.Equal(1, batch.ResultVersion);
        Assert.Equal(2, session.Version);
        Assert.Contains("Batch mutation.", session.Project().Markdown);
        Assert.Contains("Concurrent mutation.", session.Project().Markdown);
    }

    /// <summary>
    /// Revision classification on an ALREADY-REDLINED document. The receipt's change sets are a
    /// before∩after comparison, so a comparison that is not value-based reports every surviving
    /// pre-existing revision as modified — and then cascades into the execution-clock warning
    /// that tells callers not to trust <c>packageHash</c>. Both the apply and the preview path
    /// run the same enrichment, so both are asserted.
    /// </summary>
    [Fact]
    public void DS469_PreExistingRevisions_AreNeverReclassifiedByAnUnrelatedBatch()
    {
        var redlined = RedlinedBytes(out var redlinedAnchors);

        // Untracked batch: nothing about the document's revisions changes, so every change set
        // must be empty and the revision-date warning must not fire.
        using (var untracked = new DocxSession(redlined, new DocxSessionSettings
        {
            PersistAnchorIds = true,
            TrackedChanges = TrackedChangeMode.Accept,
        }))
        {
            var existing = untracked.ListRevisions();
            Assert.NotEmpty(existing);

            var preview = untracked.PreviewBatch(new[]
            {
                new MutationBatchStep("docx_edit", "replace_text",
                    s => s.ReplaceText(redlinedAnchors[1], "Untouched by the redlines.")),
            });

            Assert.True(preview.Success);
            Assert.Empty(preview.RevisionChanges.Added);
            Assert.Empty(preview.RevisionChanges.Removed);
            Assert.Empty(preview.RevisionChanges.Modified);
            Assert.DoesNotContain(preview.Warnings,
                warning => warning.Contains("Tracked-revision date attributes", StringComparison.Ordinal));
        }

        // Tracked batch: the batch's own revision is added; the pre-existing ones it never
        // touched stay out of every bucket.
        using (var tracked = new DocxSession(redlined, new DocxSessionSettings
        {
            PersistAnchorIds = true,
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Batch Author",
        }))
        {
            var existingIds = tracked.ListRevisions()
                .Select(revision => revision.Id).ToHashSet(StringComparer.Ordinal);
            Assert.NotEmpty(existingIds);

            var applied = tracked.ExecuteBatch(new[]
            {
                new MutationBatchStep("docx_edit", "replace_text",
                    s => s.ReplaceText(redlinedAnchors[1], "Tracked batch replacement.")),
            });

            Assert.True(applied.Success);
            Assert.NotEmpty(applied.RevisionChanges.Added);
            Assert.Empty(applied.RevisionChanges.Removed);
            Assert.Empty(applied.RevisionChanges.Modified);
            Assert.DoesNotContain(applied.RevisionChanges.Added,
                revision => existingIds.Contains(revision.Id));
        }
    }

    /// <summary>
    /// Cross-surface preview HTML profile. The typed core renders preview HTML directly; the
    /// callback-shaped npm client cannot, so it renders its shadow through the handle façade.
    /// Both must resolve to ONE profile, or the same batch previewed from a browser and from
    /// stdio/MCP describes two different documents. The editor's own render profile is asserted
    /// to be the wrong answer here on purpose — that is what npm used to call, and it silently
    /// drops comments, annotations and headers/footers.
    /// </summary>
    [Fact]
    public void DS470_PreviewHtmlProfile_IsOwnedByTheFacadeAndNotTheEditorRenderProfile()
    {
        var commented = CommentedBytes(out var commentedAnchors);
        var settings = new DocxSessionSettings { PersistAnchorIds = true };

        using var typed = new DocxSession(commented, settings);
        var typedPreview = typed.PreviewBatch(
            new[]
            {
                new MutationBatchStep("docx_edit", "replace_text",
                    s => s.ReplaceText(commentedAnchors[1], "Predicted body text.")),
            },
            options: new MutationBatchPreviewOptions { HtmlMode = MutationPreviewHtmlMode.Full });
        Assert.True(typedPreview.Success);

        // The façade path the browser client uses: clone, mutate the clone, render the clone.
        var liveHandle = SessionRegistry.OpenSession(commented, settings);
        try
        {
            var shadowHandle = SessionRegistry.CloneSessionForPreview(liveHandle);
            string facadeHtml;
            string editorHtml;
            try
            {
                Assert.True(SessionRegistry.Get(shadowHandle)
                    .ReplaceText(commentedAnchors[1], "Predicted body text.").Success);
                facadeHtml = DocxSessionOps.RenderPreviewHtml(shadowHandle);
                editorHtml = DocxSessionOps.RenderHtml(shadowHandle, "docx-", false, false, 1);
            }
            finally
            {
                SessionRegistry.CloseSession(shadowHandle);
            }

            Assert.Equal(typedPreview.Html, facadeHtml);
            Assert.Contains("Predicted body text.", facadeHtml, StringComparison.Ordinal);

            // What the profiles disagree about, stated rather than implied.
            Assert.Contains("Reviewer comment body.", facadeHtml, StringComparison.Ordinal);
            Assert.Contains("Preview header.", facadeHtml, StringComparison.Ordinal);
            Assert.DoesNotContain("Reviewer comment body.", editorHtml, StringComparison.Ordinal);
            Assert.DoesNotContain("Preview header.", editorHtml, StringComparison.Ordinal);
        }
        finally
        {
            SessionRegistry.CloseSession(liveHandle);
        }
    }

    /// <summary>
    /// An unavailable package hash is <c>null</c> on the wire, never <c>""</c>. Two receipts
    /// that both failed to hash must NOT satisfy
    /// <c>preview.packageHash == applied.packageHash</c> — the replay assertion the docs
    /// describe has to fail loudly when it has nothing to compare.
    /// </summary>
    [Fact]
    public void DS471_UnavailablePackageHash_IsNullOnTheWireNotAnEmptySentinel()
    {
        var unavailable = DocxSessionJson.SerializeMutationBatchResult(new MutationBatchResult
        {
            Mode = MutationBatchMode.Atomic,
            Success = true,
            Warnings = new[] { "Package equivalence hash unavailable: simulated." },
        });
        Assert.Contains("\"packageHash\":null", unavailable, StringComparison.Ordinal);
        Assert.DoesNotContain("\"packageHash\":\"\"", unavailable, StringComparison.Ordinal);

        using var session = OpenRich();
        var applied = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(BodyParagraphs(s)[0], "Hashed.")),
        });
        Assert.NotNull(applied.PackageHash);
        Assert.Contains($"\"packageHash\":\"{applied.PackageHash}\"",
            DocxSessionJson.SerializeMutationBatchResult(applied), StringComparison.Ordinal);
    }

    /// <summary>Bytes carrying a comment and a default header — the parts the editor's render
    /// profile drops and a preview's must keep.</summary>
    private static byte[] CommentedBytes(out string[] anchors)
    {
        using var seed = OpenRich(new DocxSessionSettings { PersistAnchorIds = true });
        var seedAnchors = BodyParagraphs(seed);
        Assert.True(seed.AddComment(seedAnchors[0], null, "Reviewer", "Reviewer comment body.",
            date: new DateTime(2025, 1, 2, 3, 4, 5, DateTimeKind.Utc)).Success);
        Assert.True(seed.SetHeaderText(
            seedAnchors[0], HeaderFooterKind.Default, "Preview header.").Success);
        var bytes = seed.Save(persistAnchorIds: true);

        using var probe = new DocxSession(bytes, new DocxSessionSettings { PersistAnchorIds = true });
        anchors = BodyParagraphs(probe);
        return bytes;
    }

    /// <summary>Bytes carrying tracked revisions authored into the FIRST body paragraph only.</summary>
    private static byte[] RedlinedBytes(out string[] anchors)
    {
        using var seed = OpenRich(new DocxSessionSettings
        {
            PersistAnchorIds = true,
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Original Reviewer",
        });
        var seedAnchors = BodyParagraphs(seed);
        Assert.True(seed.ReplaceText(seedAnchors[0], "Redlined first paragraph.").Success);
        var bytes = seed.Save(persistAnchorIds: true);

        using var probe = new DocxSession(bytes, new DocxSessionSettings { PersistAnchorIds = true });
        anchors = BodyParagraphs(probe);
        return bytes;
    }

    private static string NormalizeGeneratedIds(string value) =>
        Regex.Replace(value, "[0-9a-fA-F]{32}", "<generated-id>");

    private sealed class BlockingAnchorList : IReadOnlyList<Anchor>
    {
        private readonly ManualResetEventSlim _entered;
        private readonly ManualResetEventSlim _release;

        internal BlockingAnchorList(ManualResetEventSlim entered, ManualResetEventSlim release)
        {
            _entered = entered;
            _release = release;
        }

        public int Count
        {
            get
            {
                _entered.Set();
                _release.Wait();
                return 0;
            }
        }

        public Anchor this[int index] => throw new ArgumentOutOfRangeException(nameof(index));

        public IEnumerator<Anchor> GetEnumerator() =>
            Enumerable.Empty<Anchor>().GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
            GetEnumerator();
    }

    private sealed record Fingerprint(
        IReadOnlyDictionary<string, string> NormalOpcEntries,
        IReadOnlyDictionary<string, string> PersistedOpcEntries,
        string Markdown,
        string[] Anchors,
        long Version,
        int RevisionCounter,
        long FormatRevisionTicks,
        long NextTransactionId,
        string Revisions,
        string Comments,
        string Annotations,
        TrackedChangeMode TrackedChanges,
        string? RevisionAuthor,
        string Settings,
        int UndoCount,
        int RedoCount,
        long UndoMemoryBytes,
        bool UndoTrimmed,
        object? CachedProjection,
        object? InitialProjection,
        object? CachedAnchorIndex,
        object? RawOps)
    {
        internal static Fingerprint Capture(DocxSession session)
        {
            // Observe Save output on complete package clones. Calling Save on the live session
            // would itself replace its read caches and make the invariant probe perturb state.
            // Compare entry payloads rather than raw ZIP bytes: DOS entry timestamps are transport
            // metadata and may advance across successive clone saves. This matches
            // GetPackageContentHash's policy of excluding ZIP timestamps and compression details.
            var projection = session.Project();
            _ = session.AnchorIndex();
            byte[] normal;
            byte[] persisted;
            using (var normalClone = session.CreateShadowSession())
                normal = normalClone.Save(persistAnchorIds: false);
            using (var persistedClone = session.CreateShadowSession())
                persisted = persistedClone.Save(persistAnchorIds: true);
            return new Fingerprint(
                HashOpcEntries(normal),
                HashOpcEntries(persisted),
                projection.Markdown,
                projection.AnchorIndex.Select(pair =>
                        $"{pair.Key}|{pair.Value.Anchor.Kind}|{pair.Value.Anchor.Scope}|{pair.Value.Unid}|{pair.Value.PartUri}")
                    .OrderBy(value => value, StringComparer.Ordinal).ToArray(),
                session.Version,
                PrivateField<int>(session, "_revisionCounter"),
                PrivateField<long>(session, "_lastFormatRevisionTicks"),
                PrivateField<long>(session, "_nextTransactionId"),
                DocxSessionJson.SerializeRevisionList(session.ListRevisions()),
                DocxSessionJson.SerializeCommentList(session.ListComments()),
                DocxSessionJson.SerializeAnnotations(session.ListAnnotations()),
                session.TrackedChanges,
                session.RevisionAuthor,
                SettingsReceipt(PrivateField<DocxSessionSettings>(session, "_settings")),
                session.UndoCount,
                session.RedoCount,
                session.UndoMemoryBytes,
                session.UndoHistoryTrimmedForMemory,
                PrivateField<object?>(session, "_cachedProjection"),
                PrivateField<object?>(session, "_initialProjection"),
                PrivateField<object?>(session, "_cachedAnchorIndex"),
                PrivateField<object?>(session, "_raw"));
        }

        internal void AssertUnchanged(DocxSession session)
        {
            // Check identity-sensitive state before any observational API is invoked.
            Assert.Same(CachedProjection, PrivateField<object?>(session, "_cachedProjection"));
            Assert.Same(InitialProjection, PrivateField<object?>(session, "_initialProjection"));
            Assert.Same(CachedAnchorIndex, PrivateField<object?>(session, "_cachedAnchorIndex"));
            Assert.Same(RawOps, PrivateField<object?>(session, "_raw"));
            var after = Capture(session);
            Assert.Equal(
                NormalOpcEntries.OrderBy(pair => pair.Key, StringComparer.Ordinal),
                after.NormalOpcEntries.OrderBy(pair => pair.Key, StringComparer.Ordinal));
            Assert.Equal(
                PersistedOpcEntries.OrderBy(pair => pair.Key, StringComparer.Ordinal),
                after.PersistedOpcEntries.OrderBy(pair => pair.Key, StringComparer.Ordinal));
            Assert.Equal(Markdown, after.Markdown);
            Assert.Equal(Anchors, after.Anchors);
            Assert.Equal(Version, after.Version);
            Assert.Equal(RevisionCounter, after.RevisionCounter);
            Assert.Equal(FormatRevisionTicks, after.FormatRevisionTicks);
            Assert.Equal(NextTransactionId, after.NextTransactionId);
            Assert.Equal(Revisions, after.Revisions);
            Assert.Equal(Comments, after.Comments);
            Assert.Equal(Annotations, after.Annotations);
            Assert.Equal(TrackedChanges, after.TrackedChanges);
            Assert.Equal(RevisionAuthor, after.RevisionAuthor);
            Assert.Equal(Settings, after.Settings);
            Assert.Equal(UndoCount, after.UndoCount);
            Assert.Equal(RedoCount, after.RedoCount);
            Assert.Equal(UndoMemoryBytes, after.UndoMemoryBytes);
            Assert.Equal(UndoTrimmed, after.UndoTrimmed);
            Assert.Same(CachedProjection, after.CachedProjection);
            Assert.Same(InitialProjection, after.InitialProjection);
            Assert.Same(CachedAnchorIndex, after.CachedAnchorIndex);
            Assert.Same(RawOps, after.RawOps);
        }

        private static string SettingsReceipt(DocxSessionSettings settings)
        {
            var projection = settings.ProjectionSettings;
            return string.Join('|',
                settings.UndoDepth,
                settings.UndoMemoryBudgetBytes,
                settings.ValidateRawOps,
                settings.TrackedChanges,
                settings.RevisionAuthor,
                settings.PersistAnchorIds,
                settings.SmartQuotes,
                settings.EmitMarkdownPatch,
                settings.CaptureInitialProjection,
                projection.Scopes,
                projection.HeadingLevelOffset,
                projection.AnchorMode,
                projection.TableMode,
                projection.TableInlineCellMax,
                projection.TrackedChanges,
                projection.ResolveNumbering,
                projection.EmptyParagraphs,
                projection.AnchorIdRendering);
        }
    }

    private static IReadOnlyDictionary<string, string> HashOpcEntries(byte[] bytes)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        return archive.Entries.OrderBy(entry => entry.FullName, StringComparer.Ordinal)
            .ToDictionary(
                entry => entry.FullName,
                entry =>
                {
                    using var stream = entry.Open();
                    using var copy = new MemoryStream();
                    stream.CopyTo(copy);
                    return Convert.ToHexString(SHA256.HashData(copy.ToArray()));
                },
                StringComparer.Ordinal);
    }

    private static T PrivateField<T>(DocxSession session, string name) =>
        (T)typeof(DocxSession).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!
            .GetValue(session)!;

    private static DocxSession OpenRich(DocxSessionSettings? settings = null) =>
        new(RichBytes(), settings ?? new DocxSessionSettings { PersistAnchorIds = true });

    private static byte[] RichBytes()
    {
        using var stream = new MemoryStream();
        stream.Write(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        stream.Position = 0;
        using (var package = WordprocessingDocument.Open(stream, isEditable: true))
        {
            var main = package.MainDocumentPart!;
            var custom = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            custom.FeedData(new MemoryStream(Encoding.UTF8.GetBytes(
                "<vendor:data xmlns:vendor=\"urn:preview\"><vendor:item>opaque</vendor:item></vendor:data>")));
            var image = main.AddImagePart(ImagePartType.Png);
            image.FeedData(new MemoryStream(Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")));
            main.AddHyperlinkRelationship(new Uri("https://example.test/preview"), true);
            package.Save();
        }
        return stream.ToArray();
    }

    private static string[] BodyParagraphs(DocxSession session) =>
        session.Project().AnchorIndex.Keys
            .Where(id => id.StartsWith("p:body:", StringComparison.Ordinal))
            .ToArray();
}
