// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #963: when a failed op's rollback also fails, the document may be half-mutated. The
/// failing call must say so (<see cref="EditErrorCode.SessionCorrupted"/>), and every later mutation
/// must be refused with the same code instead of building on the damaged package, through the
/// facade and every transport. Test IDs use the SC prefix.
/// </summary>
public class DocxSessionCorruptionTests
{
    private static string FirstBodyParagraph(DocxSession s) =>
        s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    private static string BodyText(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.GetXDocument().Root!.ToString();
    }

    private static JsonElement J(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone();
    }

    private static string Code(string resultJson) =>
        J(resultJson).GetProperty("error").GetProperty("code").GetString()!;

    /// <summary>Make the next rollback fail, then run an op that throws mid-mutation.</summary>
    private static EditResult CorruptByFailedRollback(DocxSession s)
    {
        s.BeforeRestoreSnapshotForTests = () => throw new IOException("simulated restore failure");
        var anchor = FirstBodyParagraph(s);
        var result = DocxSessionRollbackTests.FailDuringMutation(s, () => s.InsertFootnote(anchor, 0, "Note"));
        s.BeforeRestoreSnapshotForTests = null;
        return result;
    }

    [Fact]
    public void SC001_FailingCall_ReportsCorruption_WhenItsRollbackAlsoFails()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());

        var failed = CorruptByFailedRollback(s);

        Assert.False(failed.Success);
        Assert.Equal(EditErrorCode.SessionCorrupted, failed.Error!.Code);
        Assert.Contains("simulated", s.LastRollbackError!.Message);
        Assert.True(s.IsCorrupted);
    }

    [Fact]
    public void SC002_CorruptedSession_RefusesEveryLaterMutation_ButStillReads()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        CorruptByFailedRollback(s);
        var anchor = FirstBodyParagraph(s);
        var version = s.Version;

        var insert = s.InsertParagraph(anchor, Position.After, "more");
        Assert.Equal(EditErrorCode.SessionCorrupted, insert.Error!.Code);
        Assert.Equal(EditErrorCode.SessionCorrupted, s.ReplaceText(anchor, "x").Error!.Code);
        Assert.Equal(EditErrorCode.SessionCorrupted, s.DeleteBlock(anchor).Error!.Code);
        Assert.Equal(EditErrorCode.SessionCorrupted,
            Assert.Single(s.ReplaceTextRange(anchor, "a", "b")).Error!.Code);
        Assert.Equal(EditErrorCode.SessionCorrupted,
            s.ExecuteMutation(null, session => session.InsertParagraph(anchor, Position.Before, "y")).Error!.Code);
        Assert.False(s.Undo());
        Assert.False(s.Redo());
        Assert.Throws<InvalidOperationException>(() => s.CompactRuns());
        Assert.Equal(version, s.Version);

        // Reads are not refused, so a caller can inspect what happened before reopening. Save is not
        // refused either, but the failed rollback left a partially constructed note and citation,
        // which is why the session must stop accepting edits.
        Assert.NotEmpty(s.Project().AnchorIndex);
    }

    [Fact]
    public void SC003_HealthyFailure_StaysAnInternalError()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());

        var anchor = FirstBodyParagraph(s);
        var failed = DocxSessionRollbackTests.FailDuringMutation(s, () => s.InsertFootnote(anchor, 0, "Note"));

        Assert.Equal(EditErrorCode.InternalError, failed.Error!.Code);
        Assert.False(s.IsCorrupted);
        Assert.True(s.InsertParagraph(FirstBodyParagraph(s), Position.After, "fine").Success);
    }

    [Fact]
    public void SC004_TransactionRollback_RestoresTheCheckpoint_AndClearsCorruption()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var before = s.Save();
        using (var tx = s.BeginTransaction())
        {
            CorruptByFailedRollback(s);
            Assert.True(s.IsCorrupted);
            tx.Rollback();
        }

        Assert.False(s.IsCorrupted);
        Assert.Equal(BodyText(before), BodyText(s.Save()));
        Assert.True(s.InsertParagraph(FirstBodyParagraph(s), Position.After, "fine").Success);
    }

    // ─── Through the facade and the transports ──────────────────────────

    [Fact]
    public void SC010_Facade_ReportsSessionCorrupted_OnTheFailingCallAndAfter()
    {
        var handle = DocxSessionOps.OpenSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), null);
        try
        {
            var session = SessionRegistry.Get(handle);
            var anchor = FirstBodyParagraph(session);
            session.BeforeRestoreSnapshotForTests = () => throw new IOException("simulated restore failure");

            Assert.Equal("session_corrupted", Code(DocxSessionRollbackTests.FailDuringMutation(session,
                () => DocxSessionOps.InsertFootnote(handle, anchor, 0, "Note"))));
            session.BeforeRestoreSnapshotForTests = null;

            Assert.Equal("session_corrupted", Code(DocxSessionOps.InsertParagraph(handle, anchor, Position.After, "x")));
            Assert.Equal("session_corrupted", Code(DocxSessionOps.UndoChecked(handle, null)));
            Assert.Contains("session_corrupted",
                Assert.Throws<InvalidOperationException>(() => DocxSessionOps.Undo(handle)).Message);
            Assert.Equal("session_corrupted",
                Code(Docxodus.PyHost.Dispatcher.Dispatch("insert_paragraph", J(
                    $$$"""{"handle":{{{handle}}},"anchorId":"{{{anchor}}}","position":"after","markdown":"x"}"""))));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void SC011_Mcp_ReportsSessionCorrupted()
    {
        var root = Path.Combine(Path.GetTempPath(), $"sc011-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var store = new SessionStore(new LocalFileDocumentStore(root));
        try
        {
            var path = Path.Combine(root, "d.docx");
            File.WriteAllBytes(path, DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
            var sessionId = J(Dispatcher.Call(store, "docxodus_open", J(
                $$"""{"path":{{JsonSerializer.Serialize(path)}}}"""))).GetProperty("sessionId").GetString()!;
            var session = SessionRegistry.Get(store.Get(sessionId).Handle);
            var anchor = FirstBodyParagraph(session);
            CorruptByFailedRollback(session);

            var result = Dispatcher.Call(store, "docxodus_edit", J(
                $$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"action":"insert_paragraph","anchorId":"{{anchor}}","position":"after","markdown":"x"}"""));
            Assert.Equal("session_corrupted", Code(result));
        }
        finally
        {
            store.CloseAll();
            Directory.Delete(root, recursive: true);
        }
    }
}
