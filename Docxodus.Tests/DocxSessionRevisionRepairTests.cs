// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issues #754–#758: explicit, previewable repair of native revision markup the registry
/// refuses. Every shape here is one the registry lists as malformed, ambiguous, or an orphan
/// payload; the proposals name the carriers and the evidence, the apply is atomic and undoable,
/// and anything the package cannot prove stays refused without mutation.
/// </summary>
[Collection("MCP session registry isolation")]
public sealed class DocxSessionRevisionRepairTests
{
    private const string Author = "Reviewer";
    private const string Date = "2026-01-01T00:00:00Z";

    private static XElement Mark(XName name, string? id, string author = Author, string date = Date)
    {
        var mark = new XElement(name, new XAttribute(W.author, author), new XAttribute(W.date, date));
        if (id is not null) mark.SetAttributeValue(W.id, id);
        return mark;
    }

    private static byte[] Document(string shape) =>
        MutateMain(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), root =>
        {
            var paragraphs = root.Descendants(W.p).Take(2).ToArray();
            var firstRun = paragraphs[0].Elements(W.r).First();
            switch (shape)
            {
                case "missing_id":
                    firstRun.ReplaceWith(Wrap(W.ins, null, firstRun));
                    break;
                case "nonnumeric_id":
                    firstRun.ReplaceWith(Wrap(W.ins, "not-an-integer", firstRun));
                    break;
                case "duplicate_ids":
                    foreach (var paragraph in paragraphs)
                    {
                        var run = paragraph.Elements(W.r).First();
                        run.ReplaceWith(Wrap(W.ins, "777", run));
                    }
                    break;
                case "orphan_numbering_change":
                    paragraphs[0].AddFirst(new XElement(W.pPr,
                        new XElement(W.numPr,
                            new XElement(W.ilvl, new XAttribute(W.val, 0)),
                            new XElement(W.numId, new XAttribute(W.val, 1))),
                        Mark(W.numberingChange, "5")));
                    break;
                case "orphan_numbering_change_in_run":
                    paragraphs[0].AddFirst(new XElement(W.pPr,
                        new XElement(W.numPr,
                            new XElement(W.ilvl, new XAttribute(W.val, 0)),
                            new XElement(W.numId, new XAttribute(W.val, 1)))));
                    firstRun.Add(Mark(W.numberingChange, "5"));
                    break;
                case "orphan_numbering_change_without_owner":
                    paragraphs[0].AddFirst(new XElement(W.pPr, Mark(W.numberingChange, "5")));
                    break;
                case "orphan_cell_marker":
                    paragraphs[0].AddAfterSelf(new XElement(W.tbl,
                        new XElement(W.tblPr), new XElement(W.tblGrid, new XElement(W.gridCol)),
                        new XElement(W.tr,
                            new XElement(W.tc, new XElement(W.p, new XElement(W.r, new XElement(W.t, "A")))),
                            new XElement(W.tc,
                                Mark(W.cellIns, "6"),
                                new XElement(W.p, new XElement(W.r, new XElement(W.t, "B")))))));
                    break;
                case "invalid_merge_state":
                    paragraphs[0].AddAfterSelf(new XElement(W.tbl,
                        new XElement(W.tblPr), new XElement(W.tblGrid, new XElement(W.gridCol)),
                        new XElement(W.tr,
                            new XElement(W.tc,
                                new XElement(W.tcPr, Mark(W.cellMerge, "6").WithAttribute(W.vMerge, "sideways")),
                                new XElement(W.p, new XElement(W.r, new XElement(W.t, "A")))))));
                    break;
                case "orphan_deleted_text":
                    firstRun.Element(W.t)!.ReplaceWith(new XElement(W.delText, "gone"));
                    break;
                case "orphan_deleted_text_beside_live_text":
                    firstRun.Add(new XElement(W.delText, "gone"));
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(shape));
            }
        });

    private static XElement Wrap(XName name, string? id, XElement run) =>
        Mark(name, id).WithContent(new XElement(run));

    // ─── Identity (#754, #755) ──────────────────────────────────────────

    [Theory]
    [InlineData("missing_id", "missing_revision_id")]
    [InlineData("nonnumeric_id", "invalid_revision_id")]
    public void RVR754_IdentityDefect_IsProposedRepairedAndBecomesResolvable(string shape, string code)
    {
        using var session = new DocxSession(Document(shape));
        var listed = Assert.Single(session.ListRevisions());
        Assert.Equal(code, listed.Diagnostic!.Code);
        var proposal = Assert.Single(session.ListRevisionRepairs());
        Assert.Equal(listed.Id, proposal.RevisionId);
        Assert.Equal(RevisionRepairKind.AssignIdentity, proposal.Kind);
        Assert.True(proposal.Repairable);
        Assert.Single(proposal.Carriers, c => c.StartsWith("w:ins@", StringComparison.Ordinal));
        Assert.Equal(code, proposal.Diagnostic.Code);

        var result = session.RepairRevisions(new[]
        {
            new RevisionRepairRequest { RevisionId = listed.Id, Kind = RevisionRepairKind.AssignIdentity },
        });

        Assert.True(result.Success, result.Error?.Message);
        var identity = Assert.Single(Assert.Single(result.Repairs).Identities);
        Assert.Equal(shape == "missing_id" ? null : "not-an-integer", identity.OldId);
        Assert.True(long.TryParse(identity.NewId, out var assigned) && assigned > 0);
        var repaired = Assert.Single(session.ListRevisions());
        Assert.Equal(RevisionResolutionStatus.Supported, repaired.ResolutionStatus);
        Assert.NotEqual(listed.Id, repaired.Id);
        Assert.Equal(identity.NewId, Assert.Single(repaired.ConstituentIds));
        Assert.Empty(session.ListRevisionRepairs());
        Assert.True(session.AcceptRevision(repaired.Id).Success);

        Assert.True(session.Undo()); // the accept
        Assert.True(session.Undo()); // the repair
        Assert.Equal(code, Assert.Single(session.ListRevisions()).Diagnostic!.Code);
        Assert.True(session.Redo());
        Assert.Equal(RevisionResolutionStatus.Supported, Assert.Single(session.ListRevisions()).ResolutionStatus);
        using var reopened = new DocxSession(session.Save());
        Assert.Equal(identity.NewId, Assert.Single(Assert.Single(reopened.ListRevisions()).ConstituentIds));
    }

    [Fact]
    public void RVR755_DuplicateIdentity_RepairingOneGroupMakesBothUniqueAndKeepsThemDistinct()
    {
        using var session = new DocxSession(Document("duplicate_ids"));
        var listed = session.ListRevisions();
        Assert.Equal(2, listed.Count);
        Assert.All(listed, r => Assert.Equal("duplicate_revision_id", r.Diagnostic!.Code));
        var proposals = session.ListRevisionRepairs();
        Assert.Equal(2, proposals.Count);
        Assert.All(proposals, p => Assert.Equal(RevisionRepairKind.AssignIdentity, p.Kind));

        var second = listed[1];
        var result = session.RepairRevisions(new[]
        {
            new RevisionRepairRequest { RevisionId = second.Id, Kind = RevisionRepairKind.AssignIdentity },
        });

        Assert.True(result.Success, result.Error?.Message);
        var identity = Assert.Single(Assert.Single(result.Repairs).Identities);
        Assert.Equal("777", identity.OldId);
        Assert.NotEqual("777", identity.NewId);
        var after = session.ListRevisions();
        Assert.Equal(2, after.Count);
        Assert.All(after, r => Assert.Equal(RevisionResolutionStatus.Supported, r.ResolutionStatus));
        Assert.Equal(new[] { "777", identity.NewId }.OrderBy(x => x),
            after.Select(r => Assert.Single(r.ConstituentIds)).OrderBy(x => x));
        // The untouched group keeps its native id; its public id changes only because the
        // collision suffix that disambiguated the two listings is no longer needed.
        Assert.Contains(after, r => Assert.Single(r.ConstituentIds) == "777");
        Assert.Empty(session.ListRevisionRepairs());
        Assert.True(session.AcceptAllRevisions().Success);
    }

    [Fact]
    public void RVR755b_SameIdInDifferentStories_IsLegalAndOffersNoRepair()
    {
        var input = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(BuildCrossPartDuplicate(input));
        Assert.Equal(2, session.ListRevisions().Count);
        Assert.All(session.ListRevisions(), r => Assert.Equal(RevisionResolutionStatus.Supported, r.ResolutionStatus));
        Assert.Empty(session.ListRevisionRepairs());
    }

    // ─── Owners (#756, #757) ────────────────────────────────────────────

    // A marker directly under w:pPr is never walked as a unit (unsupported_revision_family);
    // one inside a run is (orphan_numbering_revision). Both are the same misplacement.
    [Theory]
    [InlineData("orphan_numbering_change", "unsupported_revision_family")]
    [InlineData("orphan_numbering_change_in_run", "orphan_numbering_revision")]
    public void RVR756_OrphanNumberingChange_ReattachesToItsParagraphsNumPrOnly(string shape, string code)
    {
        using var session = new DocxSession(Document(shape));
        var listed = Assert.Single(session.ListRevisions());
        Assert.Equal(code, listed.Diagnostic!.Code);
        var proposal = Assert.Single(session.ListRevisionRepairs());
        Assert.Equal(RevisionRepairKind.ReattachNumberingChange, proposal.Kind);
        Assert.True(proposal.Repairable);

        var result = session.RepairRevisions(new[]
        {
            new RevisionRepairRequest { RevisionId = listed.Id, Kind = proposal.Kind },
        });

        Assert.True(result.Success, result.Error?.Message);
        var repaired = Assert.Single(session.ListRevisions());
        Assert.Equal(RevisionFamily.NumberingChange, repaired.Family);
        Assert.Equal(RevisionResolutionStatus.Supported, repaired.ResolutionStatus);
        Assert.Equal("5", Assert.Single(repaired.ConstituentIds));
        var numPr = MainRoot(session.Save()).Descendants(W.numPr).Single();
        Assert.Equal(new[] { W.ilvl, W.numId, W.numberingChange }, numPr.Elements().Select(e => e.Name));
        Assert.True(session.AcceptRevision(repaired.Id).Success);

        using var unowned = new DocxSession(Document("orphan_numbering_change_without_owner"));
        var refused = Assert.Single(unowned.ListRevisionRepairs());
        Assert.False(refused.Repairable);
        Assert.Contains("no w:numPr", refused.Reason);
        var before = MainRoot(unowned.Save());
        var attempt = unowned.RepairRevisions(new[]
        {
            new RevisionRepairRequest { RevisionId = refused.RevisionId, Kind = refused.Kind },
        });
        Assert.False(attempt.Success);
        Assert.Equal(EditErrorCode.RevisionRepairRejected, attempt.Error!.Code);
        Assert.True(XNode.DeepEquals(before, MainRoot(unowned.Save())));
        Assert.False(unowned.Undo());
    }

    [Fact]
    public void RVR757_OrphanCellMarker_ReattachesToItsCell_ButMergeStateIsNotRepairable()
    {
        using var session = new DocxSession(Document("orphan_cell_marker"));
        var listed = Assert.Single(session.ListRevisions());
        Assert.Equal("unsupported_revision_family", listed.Diagnostic!.Code);
        var proposal = Assert.Single(session.ListRevisionRepairs());
        Assert.Equal(RevisionRepairKind.ReattachCellMarker, proposal.Kind);
        Assert.True(proposal.Repairable);

        var result = session.RepairRevisions(new[]
        {
            new RevisionRepairRequest { RevisionId = listed.Id, Kind = proposal.Kind },
        });

        Assert.True(result.Success, result.Error?.Message);
        var repaired = Assert.Single(session.ListRevisions());
        Assert.Equal(RevisionFamily.CellInsert, repaired.Family);
        Assert.Equal(RevisionResolutionStatus.Supported, repaired.ResolutionStatus);
        var cell = MainRoot(session.Save()).Descendants(W.tc).Last();
        Assert.Equal(W.tcPr, cell.Elements().First().Name);
        Assert.NotNull(cell.Element(W.tcPr)!.Element(W.cellIns));
        Assert.True(session.RejectRevision(repaired.Id).Success);

        using var merge = new DocxSession(Document("invalid_merge_state"));
        Assert.Equal("invalid_cell_merge_state", Assert.Single(merge.ListRevisions()).Diagnostic!.Code);
        Assert.Empty(merge.ListRevisionRepairs());
        var attempt = merge.RepairRevisions(new[]
        {
            new RevisionRepairRequest
            {
                RevisionId = merge.ListRevisions()[0].Id, Kind = RevisionRepairKind.ReattachCellMarker,
            },
        });
        Assert.False(attempt.Success);
        Assert.Equal(EditErrorCode.RevisionRepairRejected, attempt.Error!.Code);
        Assert.False(merge.Undo());
    }

    // ─── Orphan deleted text (#758) ─────────────────────────────────────

    [Fact]
    public void RVR758_OrphanDeletedText_RestoresAsLiveText_OrWrapsWithCallerAuthorship()
    {
        using var restore = new DocxSession(Document("orphan_deleted_text"));
        var listed = Assert.Single(restore.ListRevisions());
        Assert.Equal("unsupported_revision_family", listed.Diagnostic!.Code);
        var proposals = restore.ListRevisionRepairs();
        Assert.Equal(new[] { RevisionRepairKind.RestoreOrphanText, RevisionRepairKind.WrapOrphanTextAsDeletion },
            proposals.Select(p => p.Kind));
        Assert.All(proposals, p => Assert.True(p.Repairable));
        Assert.True(proposals[1].RequiresAuthorship);

        var restored = restore.RepairRevisions(new[]
        {
            new RevisionRepairRequest { RevisionId = listed.Id, Kind = RevisionRepairKind.RestoreOrphanText },
        });
        Assert.True(restored.Success, restored.Error?.Message);
        Assert.Empty(restore.ListRevisions());
        Assert.Equal("gone", MainRoot(restore.Save()).Descendants(W.p).First().Value);

        using var wrap = new DocxSession(Document("orphan_deleted_text"));
        var entry = Assert.Single(wrap.ListRevisions());
        var before = MainRoot(wrap.Save());
        var withoutAuthorship = wrap.RepairRevisions(new[]
        {
            new RevisionRepairRequest { RevisionId = entry.Id, Kind = RevisionRepairKind.WrapOrphanTextAsDeletion },
        });
        Assert.False(withoutAuthorship.Success);
        Assert.Equal(EditErrorCode.RevisionRepairRejected, withoutAuthorship.Error!.Code);
        Assert.Contains("author", withoutAuthorship.Error.Message);
        Assert.True(XNode.DeepEquals(before, MainRoot(wrap.Save())));
        Assert.False(wrap.Undo());

        var wrapped = wrap.RepairRevisions(new[]
        {
            new RevisionRepairRequest
            {
                RevisionId = entry.Id, Kind = RevisionRepairKind.WrapOrphanTextAsDeletion,
                Author = "Recovering Reviewer", Date = "2026-02-02T10:00:00Z",
            },
        });
        Assert.True(wrapped.Success, wrapped.Error?.Message);
        var deletion = Assert.Single(wrap.ListRevisions());
        Assert.Equal(RevisionFamily.ContentDelete, deletion.Family);
        Assert.Equal(RevisionResolutionStatus.Supported, deletion.ResolutionStatus);
        Assert.Equal("Recovering Reviewer", deletion.Author);
        Assert.Equal("gone", deletion.Text);
        Assert.True(wrap.AcceptRevision(deletion.Id).Success);
        Assert.DoesNotContain("gone", MainRoot(wrap.Save()).Value);

        using var mixed = new DocxSession(Document("orphan_deleted_text_beside_live_text"));
        var mixedProposals = mixed.ListRevisionRepairs();
        Assert.True(mixedProposals.Single(p => p.Kind == RevisionRepairKind.RestoreOrphanText).Repairable);
        var wrapProposal = mixedProposals.Single(p => p.Kind == RevisionRepairKind.WrapOrphanTextAsDeletion);
        Assert.False(wrapProposal.Repairable);
        Assert.Contains("live text", wrapProposal.Reason);
    }

    // ─── Atomicity and contract ─────────────────────────────────────────

    [Fact]
    public void RVR759_OneRefusedRequest_RollsBackTheWholeCall()
    {
        using var session = new DocxSession(MutateMain(Document("missing_id"), root =>
            root.Descendants(W.p).Skip(1).First().AddFirst(new XElement(W.pPr, Mark(W.numberingChange, "9")))));
        var repairs = session.ListRevisionRepairs();
        var identity = repairs.Single(p => p.Kind == RevisionRepairKind.AssignIdentity);
        var orphan = repairs.Single(p => p.Kind == RevisionRepairKind.ReattachNumberingChange);
        Assert.False(orphan.Repairable);
        var before = MainRoot(session.Save());

        var result = session.RepairRevisions(new[]
        {
            new RevisionRepairRequest { RevisionId = identity.RevisionId, Kind = identity.Kind },
            new RevisionRepairRequest { RevisionId = orphan.RevisionId, Kind = orphan.Kind },
        });

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.RevisionRepairRejected, result.Error!.Code);
        Assert.True(XNode.DeepEquals(before, MainRoot(session.Save())));
        Assert.False(session.Undo());

        var unknown = session.RepairRevisions(new[]
        {
            new RevisionRepairRequest { RevisionId = "rev2-nope", Kind = RevisionRepairKind.AssignIdentity },
        });
        Assert.Equal(EditErrorCode.RevisionNotFound, unknown.Error!.Code);
        var wrongKind = session.RepairRevisions(new[]
        {
            new RevisionRepairRequest { RevisionId = identity.RevisionId, Kind = RevisionRepairKind.RestoreOrphanText },
        });
        Assert.Equal(EditErrorCode.RevisionRepairRejected, wrongKind.Error!.Code);
        Assert.Contains("offers no RestoreOrphanText", wrongKind.Error.Message);
    }

    [Fact]
    public void RVR760_McpTrackChanges_ListsAndAppliesRepairsThroughTheSameOwner()
    {
        var root = Path.Combine(Path.GetTempPath(), $"revision-repair-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var path = Path.Combine(root, "document.docx");
        File.WriteAllBytes(path, Document("missing_id"));
        var store = new SessionStore(new LocalFileDocumentStore(root));
        try
        {
            var sessionId = JsonDocument.Parse(Dispatcher.Call(store, "docxodus_open",
                JsonSerializer.SerializeToElement(new { path }))).RootElement.GetProperty("sessionId").GetString()!;
            using var repairs = JsonDocument.Parse(Dispatcher.Call(store, "docxodus_track_changes",
                JsonSerializer.SerializeToElement(new { sessionId, action = "repairs" })));
            var proposal = repairs.RootElement.GetProperty("repairs").EnumerateArray().Single();
            Assert.Equal("assign_identity", proposal.GetProperty("kind").GetString());
            Assert.True(proposal.GetProperty("repairable").GetBoolean());
            Assert.Equal("missing_revision_id", proposal.GetProperty("diagnostic").GetProperty("code").GetString());

            using var applied = JsonDocument.Parse(Dispatcher.Call(store, "docxodus_track_changes",
                JsonSerializer.SerializeToElement(new
                {
                    sessionId, action = "repair",
                    repairs = new[] { new { revisionId = proposal.GetProperty("revisionId").GetString(), kind = "assign_identity" } },
                })));
            Assert.True(applied.RootElement.GetProperty("success").GetBoolean());
            var mapping = applied.RootElement.GetProperty("repairs")[0].GetProperty("identities")[0];
            Assert.Equal(JsonValueKind.Null, mapping.GetProperty("oldId").ValueKind);
            Assert.NotEqual("", mapping.GetProperty("newId").GetString());

            using var listed = JsonDocument.Parse(Dispatcher.Call(store, "docxodus_track_changes",
                JsonSerializer.SerializeToElement(new { sessionId, action = "list" })));
            Assert.Equal("supported", listed.RootElement.GetProperty("revisions")[0]
                .GetProperty("resolutionStatus").GetString());

            using var refused = JsonDocument.Parse(Dispatcher.Call(store, "docxodus_track_changes",
                JsonSerializer.SerializeToElement(new
                {
                    sessionId, action = "repair",
                    repairs = new[] { new { revisionId = "rev2-nope", kind = "assign_identity" } },
                })));
            Assert.False(refused.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal("revision_not_found",
                refused.RootElement.GetProperty("error").GetProperty("code").GetString());
        }
        finally
        {
            store.CloseAll();
            Directory.Delete(root, recursive: true);
        }
    }

    // ─── helpers ────────────────────────────────────────────────────────

    private static byte[] MutateMain(byte[] input, Action<XElement> mutate)
    {
        using var stream = new MemoryStream();
        stream.Write(input);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var xDocument = document.MainDocumentPart!.GetXDocument();
            mutate(xDocument.Root!);
            document.MainDocumentPart.PutXDocument();
        }
        return stream.ToArray();
    }

    /// <summary>The same numeric id on an insertion in the body and one in a header part.</summary>
    private static byte[] BuildCrossPartDuplicate(byte[] input)
    {
        using var stream = new MemoryStream();
        stream.Write(input);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var header = main.AddNewPart<HeaderPart>();
            header.PutXDocument(new XDocument(new XElement(W.hdr,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XElement(W.p, Wrap(W.ins, "777", new XElement(W.r, new XElement(W.t, "Header")))))));
            var mainDocument = main.GetXDocument();
            var firstRun = mainDocument.Root!.Descendants(W.p).First().Elements(W.r).First();
            firstRun.ReplaceWith(Wrap(W.ins, "777", firstRun));
            var body = mainDocument.Root.Element(W.body)!;
            var sectPr = body.Element(W.sectPr);
            if (sectPr is null)
            {
                sectPr = new XElement(W.sectPr);
                body.Add(sectPr);
            }
            sectPr.AddFirst(new XElement(W.headerReference,
                new XAttribute(R.id, main.GetIdOfPart(header)),
                new XAttribute(W.type, "default")));
            main.PutXDocument();
        }
        return stream.ToArray();
    }

    private static XElement MainRoot(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var root = new XElement(document.MainDocumentPart!.GetXDocument().Root!);
        root.DescendantsAndSelf().Attributes()
            .Where(a => a.Name.Namespace == PtOpenXml.pt || a.IsNamespaceDeclaration)
            .Remove();
        return root;
    }
}

internal static class RevisionRepairTestExtensions
{
    public static XElement WithAttribute(this XElement element, XName name, string value)
    {
        element.SetAttributeValue(name, value);
        return element;
    }

    public static XElement WithContent(this XElement element, params object[] content)
    {
        element.Add(content);
        return element;
    }
}
