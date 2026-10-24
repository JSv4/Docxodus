using System.Text.Json;
using System.Xml.Linq;
using Docxodus.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

[Collection("MCP session registry isolation")]
public class DocxSessionRevisionIdentityTests
{
    private static readonly XNamespace People = "http://schemas.microsoft.com/office/word/2012/wordml";
    private const string OurAccount = "S::pat@ourfirm.example::22222222-2222-2222-2222-222222222222";
    private const string OtherAccount = "S::pat@counterparty.example::11111111-1111-1111-1111-111111111111";
    private const string Proposed = "Counterparty's new clause.";

    [Theory]
    [InlineData("Pat", OtherAccount, true, false)]
    [InlineData("Pat", null, true, false)]
    [InlineData("Word Pat", "S::PAT@OURFIRM.EXAMPLE::33333333-3333-3333-3333-333333333333", true, true)]
    [InlineData("Word Pat", "S::former@ourfirm.example::22222222-2222-2222-2222-222222222222", true, true)]
    [InlineData("Other Reviewer", OtherAccount, true, false)]
    [InlineData("Pat", OtherAccount, false, true)]
    public void Deleting_a_pending_paragraph_uses_accounts_when_configured_and_names_otherwise(
        string recordedAuthor, string? recordedAccount, bool identities, bool own)
    {
        var input = Document(recordedAuthor, recordedAccount);
        AssertValid(input);
        // Enter through the shared settings parser, so the reproducer runs on current main
        // before the typed setting exists rather than failing only at compilation.
        var settings = DocxSessionJson.ParseSettings(JsonSerializer.Serialize(new
        {
            trackedChanges = "render_inline",
            revisionAuthor = "Pat",
            revisionIdentities = identities ? new[] { new { providerId = "AD", userId = OurAccount } } : null,
        }));
        using var session = new DocxSession(input, settings);
        var anchor = session.FindByText(Proposed)!.Anchor.Id;

        var result = session.DeleteBlock(anchor);

        Assert.True(result.Success, result.Error?.Message);
        var saved = session.Save();
        AssertValid(saved);
        var body = Body(saved);
        if (own)
        {
            Assert.Equal(new[] { "Before.", "After." }, body.Elements(W.p).Select(p => p.Value));
            Assert.Empty(body.Descendants(W.del));
        }
        else
        {
            Assert.Equal(3, body.Elements(W.p).Count());
            Assert.Contains(Proposed, string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
            Assert.Contains(body.Descendants(W.del), d => (string?)d.Attribute(W.author) == "Pat");
        }
    }

    [Fact]
    public void First_tracked_change_claims_a_free_name_and_undo_redo_restores_the_people_part()
    {
        using var session = OpenAccounts(Document("Earlier", null));
        var version = session.Version;
        Assert.Equal(RevisionAuthorNameStatus.Free, session.GetRevisionAuthorNameStatus());
        Assert.Null(PeopleRoot(session.Save()));
        Assert.Equal(0, session.UndoCount);

        Assert.True(session.ReplaceText(session.FindByText("Before.")!.Anchor.Id, "Edited.").Success);

        var saved = session.Save();
        AssertValid(saved);
        AssertPerson(saved, "Pat", OurAccount);
        Assert.Equal(RevisionAuthorNameStatus.Own, session.GetRevisionAuthorNameStatus());
        Assert.Equal(version + 1, session.Version);
        Assert.True(session.Undo());
        Assert.Null(PeopleRoot(session.Save()));
        Assert.Equal(RevisionAuthorNameStatus.Free, session.GetRevisionAuthorNameStatus());
        Assert.True(session.Redo());
        AssertPerson(session.Save(), "Pat", OurAccount);
        AssertValid(session.Save());
        using var reopened = OpenAccounts(session.Save());
        Assert.Equal(RevisionAuthorNameStatus.Own, reopened.GetRevisionAuthorNameStatus());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void First_comment_or_reply_claims_only_the_session_author_and_is_undoable(bool reply)
    {
        using var session = new DocxSession(Document("Earlier", null), new DocxSessionSettings { RevisionAuthor = "Pat" });
        var anchor = session.FindByText("Before.")!.Anchor.Id;
        if (reply) Assert.True(session.AddComment(anchor, null, "Earlier", "Parent.").Success);
        session.SetRevisionIdentities(new[] { new RevisionIdentity("AD", OurAccount) });
        var result = reply
            ? session.AddCommentReply(session.ListComments().Single().DefAnchorId, "Pat", "Reply.")
            : session.AddComment(anchor, null, "Pat", "Note.");
        Assert.True(result.Success, result.Error?.Message);
        AssertPerson(session.Save(), "Pat", OurAccount);
        AssertValid(session.Save(), FileFormatVersions.Microsoft365);
        Assert.True(session.Undo());
        Assert.Null(PeopleRoot(session.Save()));
        Assert.True(session.Redo());
        AssertPerson(session.Save(), "Pat", OurAccount);
    }

    [Fact]
    public void An_explicit_other_comment_author_is_not_claimed_as_the_session_account()
    {
        using var session = OpenAccounts(Document("Earlier", null));
        Assert.True(session.AddComment(session.FindByText("Before.")!.Anchor.Id, null, "Another reviewer", "Note.").Success);
        Assert.Null(PeopleRoot(session.Save()));
    }

    [Theory]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("comment")]
    [InlineData("people")]
    public void Names_used_outside_body_revisions_are_taken(string kind)
    {
        var bytes = AddUsedName(Document("Earlier", null), kind);
        AssertValid(bytes);
        using var session = OpenAccounts(bytes);
        Assert.Equal(RevisionAuthorNameStatus.Taken, session.GetRevisionAuthorNameStatus());
        Assert.True(session.ReplaceText(session.FindByText("Before.")!.Anchor.Id, "Edited.").Success);
        Assert.DoesNotContain(PeopleRoot(session.Save())?.Elements(People + "person") ?? Enumerable.Empty<XElement>(),
            person => (string?)person.Attribute(People + "author") == "Pat"
                && (string?)person.Element(People + "presenceInfo")?.Attribute(People + "userId") == OurAccount);
        AssertValid(session.Save());
    }

    [Fact]
    public void Changing_the_display_name_does_not_reclassify_an_existing_unattributed_author_as_free()
    {
        using var session = OpenAccounts(Document("Reserved", null));
        session.SetRevisionAuthor("Reserved");
        Assert.Equal(RevisionAuthorNameStatus.Taken, session.GetRevisionAuthorNameStatus());
        Assert.True(session.DeleteBlock(session.FindByText(Proposed)!.Anchor.Id).Success);
        Assert.Equal(3, Body(session.Save()).Elements(W.p).Count());
        Assert.Null(PeopleRoot(session.Save()));
    }

    [Fact]
    public void Clearing_identities_restores_legacy_ownership_and_settings_take_a_defensive_copy()
    {
        var identities = new List<RevisionIdentity> { new("AD", OurAccount) };
        using var session = new DocxSession(Document("Pat", OtherAccount), new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Pat", RevisionIdentities = identities,
        });
        identities[0] = new("AD", OtherAccount);
        Assert.Equal(OurAccount, Assert.Single(session.RevisionIdentities).UserId);
        Assert.Equal(RevisionAuthorNameStatus.Taken, session.GetRevisionAuthorNameStatus());
        var version = session.Version;
        session.SetRevisionIdentities(null);
        Assert.Equal(version, session.Version);
        Assert.Equal(0, session.UndoCount);
        Assert.True(session.DeleteBlock(session.FindByText(Proposed)!.Anchor.Id).Success);
        Assert.Equal(2, Body(session.Save()).Elements(W.p).Count());
    }

    [Fact]
    public void A_configuration_only_transaction_restores_the_account_and_free_name_baseline()
    {
        using var session = OpenAccounts(Document("Word Pat", OurAccount));
        var hash = session.GetPackageContentHash();
        using (session.BeginTransaction())
        {
            session.SetRevisionIdentities(new[] { new RevisionIdentity("AD", OtherAccount) });
            Assert.Equal(RevisionAuthorNameStatus.Taken, session.GetRevisionAuthorNameStatus("Word Pat"));
            session.SetRevisionAuthor("Changed");
        }
        Assert.Equal("Pat", session.RevisionAuthor);
        Assert.Equal(OurAccount, Assert.Single(session.RevisionIdentities).UserId);
        Assert.Equal(RevisionAuthorNameStatus.Own, session.GetRevisionAuthorNameStatus("Word Pat"));
        Assert.Equal(hash, session.GetPackageContentHash());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Retained_previews_bind_account_configuration_and_allow_an_equivalent_reset(bool changed)
    {
        using var session = OpenAccounts(Document("Earlier", null));
        var anchor = session.FindByText("Before.")!.Anchor.Id;
        var preview = session.PreviewBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text", s => s.ReplaceText(anchor, "Edited.")),
        }, options: new MutationBatchPreviewOptions { Retain = true });
        Assert.True(preview.Success);
        var hash = session.GetPackageContentHash();
        Assert.Null(PeopleRoot(session.Save()));
        session.SetRevisionIdentities(new[] { new RevisionIdentity("AD", changed ? OtherAccount : OurAccount) });

        var commit = session.CommitPreview(preview.Retention!.PreviewId);

        if (changed)
        {
            Assert.False(commit.Success);
            Assert.Equal(EditErrorCode.PreviewStale, commit.Failure?.Error.Code);
            Assert.Equal(hash, session.GetPackageContentHash());
        }
        else
        {
            Assert.True(commit.Success);
            AssertPerson(session.Save(), "Pat", OurAccount);
            Assert.True(session.Undo());
            Assert.Null(PeopleRoot(session.Save()));
        }
    }

    [Fact]
    public void A_preview_keeps_the_original_taken_name_baseline_after_old_revisions_are_removed()
    {
        using var session = OpenAccounts(Document("Reserved", null));
        session.SetTrackedChanges(TrackedChangeMode.Accept);
        Assert.True(session.DeleteBlock(session.FindByText(Proposed)!.Anchor.Id).Success);
        session.SetRevisionAuthor("Reserved");
        session.SetTrackedChanges(TrackedChangeMode.RenderInline);
        Assert.Equal(RevisionAuthorNameStatus.Taken, session.GetRevisionAuthorNameStatus());
        var anchor = session.FindByText("Before.")!.Anchor.Id;
        var preview = session.PreviewBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text", s => s.ReplaceText(anchor, "Edited.")),
        }, options: new MutationBatchPreviewOptions { Retain = true });
        Assert.True(preview.Success);
        Assert.True(session.CommitPreview(preview.Retention!.PreviewId).Success);
        Assert.Null(PeopleRoot(session.Save()));
        Assert.Equal(RevisionAuthorNameStatus.Taken, session.GetRevisionAuthorNameStatus());
    }

    [Theory]
    [InlineData("Windows Live", "live-123", "Windows Live", "live-123", true)]
    [InlineData("Windows Live", "LIVE-123", "Windows Live", "live-123", false)]
    [InlineData("Other Provider", OurAccount, "AD", OurAccount, false)]
    [InlineData("AD", "S::::", "AD", OurAccount, false)]
    public void Provider_accounts_match_exactly_and_empty_AD_components_do_not_match(
        string provider, string account, string configuredProvider, string configuredAccount, bool own)
    {
        var input = Document("Word Pat", account, provider);
        AssertValid(input);
        using var session = OpenAccounts(input, new RevisionIdentity(configuredProvider, configuredAccount));
        Assert.Equal(own ? RevisionAuthorNameStatus.Own : RevisionAuthorNameStatus.Taken,
            session.GetRevisionAuthorNameStatus("Word Pat"));
        Assert.True(session.DeleteBlock(session.FindByText(Proposed)!.Anchor.Id).Success);
        Assert.Equal(own ? 2 : 3, Body(session.Save()).Elements(W.p).Count());
    }

    [Fact]
    public void Any_configured_identity_can_own_a_revision_and_new_claims_use_the_first()
    {
        using var session = OpenAccounts(Document("Word Pat", OurAccount),
            new RevisionIdentity("Windows Live", "live-123"), new RevisionIdentity("AD", OurAccount));
        Assert.Equal(RevisionAuthorNameStatus.Own, session.GetRevisionAuthorNameStatus("Word Pat"));
        Assert.True(session.ReplaceText(session.FindByText("Before.")!.Anchor.Id, "Edited.").Success);
        AssertPerson(session.Save(), "Pat", "live-123", "Windows Live");
        Assert.True(session.DeleteBlock(session.FindByText(Proposed)!.Anchor.Id).Success);
        Assert.Equal(2, Body(session.Save()).Elements(W.p).Count());
        AssertValid(session.Save());
    }

    [Fact]
    public void The_default_author_can_claim_its_account_without_an_explicit_display_name()
    {
        using var session = OpenAccounts(Document("Earlier", null));
        session.SetRevisionAuthor(null);
        Assert.Equal(RevisionAuthorNameStatus.Free, session.GetRevisionAuthorNameStatus());
        Assert.True(session.ReplaceText(session.FindByText("Before.")!.Anchor.Id, "Edited.").Success);
        AssertPerson(session.Save(), "docxodus", OurAccount);
    }

    [Theory]
    [InlineData("Pat", OtherAccount, false)]
    [InlineData("Pat", null, false)]
    [InlineData("Word Pat", OurAccount, true)]
    [InlineData("Pat", OurAccount, true)]
    public void Span_replacement_preserves_foreign_insertions_and_uninserts_the_same_account(
        string author, string? account, bool own)
    {
        using var session = OpenAccounts(Document(author, account));
        var result = session.ReplaceTextAtSpan(session.FindByText(Proposed)!.Anchor.Id, 0, Proposed.Length, "Replacement.");
        Assert.True(result.Success, result.Error?.Message);
        var saved = session.Save();
        var body = Body(saved);
        Assert.Equal(own ? "" : Proposed, string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        Assert.Contains(body.Descendants(W.t), t => t.Value == "Replacement.");
        AssertValid(saved);
    }

    [Theory]
    [InlineData("Pat", OtherAccount, false)]
    [InlineData("Pat", null, false)]
    [InlineData("Word Pat", OurAccount, true)]
    [InlineData("Pat", OurAccount, true)]
    public void Image_edits_use_the_insertion_authors_account(string author, string? account, bool own)
    {
        byte[] bytes;
        using (var seed = new DocxSession(Document(author, account)))
        {
            Assert.True(seed.InsertImage(seed.FindByText("Before.")!.Anchor.Id, 0, DocxSessionImageTests.Png(2, 3)).Success);
            bytes = seed.Save();
        }
        using (var stream = new MemoryStream())
        {
            stream.Write(bytes);
            stream.Position = 0;
            using (var package = WordprocessingDocument.Open(stream, true))
            {
                var main = package.MainDocumentPart!;
                var body = main.GetXDocument().Root!.Element(W.body)!;
                var image = body.Descendants(W.drawing).Single().Parent!;
                image.Remove();
                body.Elements(W.p).ElementAt(1).Element(W.ins)!.AddFirst(image);
                main.PutXDocument();
            }
            bytes = stream.ToArray();
        }
        AssertValid(bytes);
        using var session = OpenAccounts(bytes);
        var occurrence = Assert.Single(session.ListImages());
        Assert.Equal(own, occurrence.CanMutate);
        var edit = session.SetImageMetadata(occurrence.Id, "Updated", null);
        Assert.Equal(own, edit.Success);
        if (own)
        {
            Assert.Equal("Updated", Assert.Single(session.ListImages()).AltText);
            Assert.True(session.Undo());
        }
        else Assert.Equal(EditErrorCode.UnsupportedImageMarkup, edit.Error?.Code);
        AssertValid(session.Save());
    }

    [Theory]
    [InlineData("Pat", OtherAccount, false)]
    [InlineData("Word Pat", OurAccount, true)]
    [InlineData("Pat", OurAccount, true)]
    public void Merging_a_split_paragraph_cannot_resolve_a_namesakes_pending_properties(
        string author, string account, bool own)
    {
        byte[] bytes;
        using (var seed = new DocxSession(Document(author, account, pageBreakBefore: true), new DocxSessionSettings
        {
            RevisionAuthor = author, TrackedChanges = TrackedChangeMode.RenderInline,
        }))
        {
            Assert.True(seed.SplitParagraph(seed.FindByText("Before.")!.Anchor.Id, 3).Success);
            bytes = seed.Save();
        }
        AssertValid(bytes);
        Assert.NotNull(Body(bytes).Elements(W.p).ElementAt(1).Element(W.pPr)?.Element(W.pPrChange));
        using var session = OpenAccounts(bytes);
        var before = session.GetPackageContentHash();
        var first = session.FindByText("Bef")!.Anchor.Id;
        var second = session.FindByText("ore.")!.Anchor.Id;
        var result = session.MergeParagraphs(first, second);
        Assert.Equal(own, result.Success);
        if (!own)
        {
            Assert.Equal(EditErrorCode.UnresolvedStructuralRevision, result.Error?.Code);
            Assert.Equal(before, session.GetPackageContentHash());
            Assert.Equal(0, session.UndoCount);
        }
        AssertValid(session.Save());
    }

    [Fact]
    public void A_failed_atomic_batch_rolls_back_its_first_identity_claim()
    {
        using var session = OpenAccounts(Document("Earlier", null));
        var hash = session.GetPackageContentHash();
        var anchor = session.FindByText("Before.")!.Anchor.Id;
        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text", s => s.ReplaceText(anchor, "Edited.")),
            new MutationBatchStep("docx_edit", "delete_block", s => s.DeleteBlock("p:body:missing")),
        }, MutationBatchMode.Atomic);
        Assert.False(result.Success);
        Assert.True(result.RolledBack);
        Assert.Null(PeopleRoot(session.Save()));
        Assert.Equal(hash, session.GetPackageContentHash());
        Assert.Equal(RevisionAuthorNameStatus.Free, session.GetRevisionAuthorNameStatus());
        Assert.Equal(0, session.UndoCount);
    }

    [Fact]
    public void An_invalid_text_edit_does_not_claim_the_free_author_name()
    {
        using var session = OpenAccounts(Document("Earlier", null));
        var hash = session.GetPackageContentHash();
        var result = session.ReplaceText(session.FindByText("Before.")!.Anchor.Id, "Invalid\0text");
        Assert.False(result.Success);
        Assert.Null(PeopleRoot(session.Save()));
        Assert.Equal(hash, session.GetPackageContentHash());
        Assert.Equal(0, session.UndoCount);
    }

    [Theory]
    [InlineData("stdio", false)]
    [InlineData("stdio", true)]
    [InlineData("mcp", false)]
    [InlineData("mcp", true)]
    public void Transport_open_configuration_and_name_queries_share_account_ownership(string protocol, bool atOpen)
    {
        var bytes = Document("Pat", OtherAccount);
        var identities = new[] { new { providerId = "AD", userId = OurAccount } };
        var root = Path.Combine(Path.GetTempPath(), "revision-identities-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var store = new Docxodus.McpServer.SessionStore(new Docxodus.McpServer.LocalFileDocumentStore(root));
        var handle = 0;
        try
        {
            string? sessionId = null;
            if (protocol == "stdio")
                handle = int.Parse(Docxodus.PyHost.Dispatcher.Dispatch("open_session", Args(new
                {
                    docxB64 = Convert.ToBase64String(bytes),
                    settings = new { trackedChanges = "render_inline", revisionAuthor = "Pat", revisionIdentities = atOpen ? identities : null },
                })), System.Globalization.CultureInfo.InvariantCulture);
            else
            {
                var path = Path.Combine(root, "input.docx");
                File.WriteAllBytes(path, bytes);
                var opened = ReadJson(Docxodus.McpServer.Dispatcher.Call(store, "docxodus_open", Args(new
                {
                    path, trackedChanges = "render_inline", revisionAuthor = "Pat", revisionIdentities = atOpen ? identities : null,
                })));
                sessionId = opened.GetProperty("sessionId").GetString()!;
                handle = store.Get(sessionId).Handle;
            }
            void Configure(object? accounts)
            {
                if (protocol == "stdio")
                    Docxodus.PyHost.Dispatcher.Dispatch("set_revision_identities", Args(new { handle, revisionIdentities = accounts }));
                else Docxodus.McpServer.Dispatcher.Call(store, "docxodus_track_changes",
                    Args(new { sessionId, action = "set_identities", revisionIdentities = accounts }));
            }
            JsonElement Query() => ReadJson(protocol == "stdio"
                ? Docxodus.PyHost.Dispatcher.Dispatch("get_revision_author_name_status", Args(new { handle }))
                : Docxodus.McpServer.Dispatcher.Call(store, "docxodus_track_changes", Args(new { sessionId, action = "author_name_status" })));
            if (!atOpen) Configure(identities);
            Assert.Equal("taken", Query().GetProperty("status").GetString());
            Assert.Equal("Pat", Query().GetProperty("author").GetString());
            Configure(null);
            Assert.Equal("own", Query().GetProperty("status").GetString());
            Configure(identities);
            var session = SessionRegistry.Get(handle);
            Assert.Equal(0, session.UndoCount);
            Assert.Equal(0, session.Version);
            Assert.True(session.DeleteBlock(session.FindByText(Proposed)!.Anchor.Id).Success);
            Assert.Equal(3, Body(session.Save()).Elements(W.p).Count());
            AssertValid(session.Save());
        }
        finally
        {
            store.CloseAll();
            if (protocol == "stdio" && handle != 0) DocxSessionOps.CloseSession(handle);
            Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("[null]")]
    [InlineData("[{\"providerId\":\"AD\",\"userId\":\"\"}]")]
    [InlineData("[{\"providerId\":\"AD\"}]")]
    [InlineData("[{\"providerId\":\"AD\",\"userId\":\"\\u0000\"}]")]
    [InlineData("[{\"providerId\":\"\\u0000\",\"userId\":\"account\"}]")]
    public void Invalid_identity_configuration_does_not_replace_the_existing_accounts(string json)
    {
        var handle = DocxSessionOps.OpenSession(Document("Word Pat", OurAccount), new DocxSessionSettings
        {
            RevisionAuthor = "Pat", RevisionIdentities = new[] { new RevisionIdentity("AD", OurAccount) },
        });
        try
        {
            var session = SessionRegistry.Get(handle);
            var hash = session.GetPackageContentHash();
            Assert.Throws<ArgumentException>(() => DocxSessionOps.SetRevisionIdentities(handle, json));
            Assert.Equal(RevisionAuthorNameStatus.Own, session.GetRevisionAuthorNameStatus("Word Pat"));
            Assert.Equal(hash, session.GetPackageContentHash());
            Assert.Equal(0, session.UndoCount);
        }
        finally { DocxSessionOps.CloseSession(handle); }
    }

    private static JsonElement Args(object value) => ReadJson(JsonSerializer.Serialize(value));

    private static JsonElement ReadJson(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }

    private static DocxSession OpenAccounts(byte[] bytes, params RevisionIdentity[] identities) => new(bytes, new DocxSessionSettings
    {
        TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Pat",
        RevisionIdentities = identities.Length == 0 ? new[] { new RevisionIdentity("AD", OurAccount) } : identities,
    });

    private static XElement? PeopleRoot(byte[] bytes)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return package.MainDocumentPart!.WordprocessingPeoplePart?.GetXDocument().Root is { } root ? new XElement(root) : null;
    }

    private static void AssertPerson(byte[] bytes, string author, string account, string provider = "AD")
    {
        var person = Assert.Single(PeopleRoot(bytes)!.Elements(People + "person"),
            person => (string?)person.Attribute(People + "author") == author);
        Assert.Equal(provider, (string?)person.Element(People + "presenceInfo")?.Attribute(People + "providerId"));
        Assert.Equal(account, (string?)person.Element(People + "presenceInfo")?.Attribute(People + "userId"));
    }

    private static byte[] AddUsedName(byte[] bytes, string kind)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var package = WordprocessingDocument.Open(stream, true))
        {
            var main = package.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W.body)!;
            if (kind is "header" or "footer")
            {
                OpenXmlPart part = kind == "header" ? main.AddNewPart<HeaderPart>() : main.AddNewPart<FooterPart>();
                part.PutXDocument(new XDocument(new XElement(kind == "header" ? W.hdr : W.ftr,
                    new XElement(W.p, new XElement(W.ins, new XAttribute(W.id, "7"), new XAttribute(W.author, "Pat"),
                        new XAttribute(W.date, "2026-05-01T09:00:00Z"), new XElement(W.r, new XElement(W.t, "Story text.")))))));
                body.Add(new XElement(W.sectPr, new XElement(kind == "header" ? W.headerReference : W.footerReference,
                    new XAttribute(W.type, "default"), new XAttribute(R.id, main.GetIdOfPart(part)))));
            }
            else if (kind == "comment")
            {
                main.AddNewPart<WordprocessingCommentsPart>().PutXDocument(new XDocument(new XElement(W.comments,
                    new XElement(W.comment, new XAttribute(W.id, "7"), new XAttribute(W.author, "Pat"), Paragraph("Note.")))));
                var paragraph = body.Elements(W.p).First();
                paragraph.AddFirst(new XElement(W.commentRangeStart, new XAttribute(W.id, "7")));
                paragraph.Add(new XElement(W.commentRangeEnd, new XAttribute(W.id, "7")),
                    new XElement(W.r, new XElement(W.commentReference, new XAttribute(W.id, "7"))));
            }
            else main.AddNewPart<WordprocessingPeoplePart>().PutXDocument(new XDocument(
                new XElement(People + "people", new XElement(People + "person", new XAttribute(People + "author", "Pat"),
                    new XElement(People + "presenceInfo", new XAttribute(People + "providerId", "AD"),
                        new XAttribute(People + "userId", OtherAccount))))));
            main.PutXDocument();
        }
        return stream.ToArray();
    }

    private static byte[] Document(string author, string? account, string provider = "AD", bool pageBreakBefore = false)
    {
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            var stamp = new[] { new XAttribute(W.id, "1"), new XAttribute(W.author, author),
                new XAttribute(W.date, "2026-05-01T09:00:00Z") };
            main.PutXDocument(new XDocument(new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w), new XElement(W.body,
                    Paragraph("Before.", pageBreakBefore),
                    new XElement(W.p,
                        new XElement(W.pPr, new XElement(W.rPr, new XElement(W.ins, stamp))),
                        new XElement(W.ins, stamp.Where(a => a.Name != W.id).Select(a => new XAttribute(a)),
                            new XAttribute(W.id, "2"),
                            new XElement(W.r, new XElement(W.t, Proposed)))),
                    Paragraph("After.")))));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
            if (account is not null)
                main.AddNewPart<WordprocessingPeoplePart>().PutXDocument(new XDocument(
                    new XElement(People + "people", new XElement(People + "person",
                        new XAttribute(People + "author", author), new XElement(People + "presenceInfo",
                            new XAttribute(People + "providerId", provider), new XAttribute(People + "userId", account))))));
        }
        return stream.ToArray();
    }

    private static XElement Paragraph(string text, bool pageBreakBefore = false) => new(W.p,
        pageBreakBefore ? new XElement(W.pPr, new XElement(W.pageBreakBefore)) : null,
        new XElement(W.r, new XElement(W.t, text)));

    private static XElement Body(byte[] bytes)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return new XElement(package.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!);
    }

    private static void AssertValid(byte[] bytes, FileFormatVersions version = FileFormatVersions.Office2013)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator(version).Validate(package));
    }
}
