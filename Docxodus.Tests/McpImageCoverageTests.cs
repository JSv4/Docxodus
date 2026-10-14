// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Issue #762 over the MCP server: the operation matrix on listed pictures, the
/// embed_linked action, and tracked image mutations under preview and apply.</summary>
[Collection("MCP session registry isolation")]
public sealed class McpImageCoverageTests : IDisposable
{
    private readonly string _root;
    private readonly string _path;
    private readonly SessionStore _store;

    public McpImageCoverageTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"mcp-image-coverage-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
        _path = Path.Combine(_root, "images.docx");
        File.WriteAllBytes(_path, DocxSessionImageCoverageTests.BuildCoverageFixture());
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public void Dispose()
    {
        _store.CloseAll();
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    [Fact]
    public void MCP762a_ListingCarriesTheMatrix_AndEmbedLinkedIsAnAction()
    {
        var capabilities = Call("docxodus_images", new { action = "capabilities" }).GetProperty("capabilities");
        Assert.Contains("embed_linked", capabilities.GetProperty("operations").EnumerateArray().Select(value => value.GetString()));
        Assert.Contains("tight", capabilities.GetProperty("mutableWrapModes").EnumerateArray().Select(value => value.GetString()));
        Assert.Equal(7, capabilities.GetProperty("trackedOperations").GetArrayLength());
        Assert.Contains(capabilities.GetProperty("markups").EnumerateArray(),
            markup => markup.GetProperty("markup").GetString() == "legacy_vml");
        Assert.True(Assert.Single(capabilities.GetProperty("formats").EnumerateArray(),
            format => format.GetProperty("format").GetString() == "webp").GetProperty("canInsert").GetBoolean());

        var sessionId = OpenSession(tracked: false);
        var images = Images(sessionId);
        Assert.Equal(2, images.Length);
        var linked = Assert.Single(images, image => image.GetProperty("isLinked").GetBoolean());
        var embedded = Assert.Single(images, image => !image.GetProperty("isLinked").GetBoolean());
        Assert.False(linked.GetProperty("canMutate").GetBoolean());
        var operations = linked.GetProperty("operations").EnumerateArray().ToArray();
        Assert.False(Operation(operations, "replace").GetProperty("canMutate").GetBoolean());
        Assert.Contains("embed_linked", Operation(operations, "replace").GetProperty("reason").GetString());
        Assert.True(Operation(operations, "embed_linked").GetProperty("canMutate").GetBoolean());
        Assert.True(Operation(operations, "set_metadata").GetProperty("canMutate").GetBoolean());
        Assert.True(embedded.GetProperty("canMutate").GetBoolean());
        Assert.Equal("picture is already embedded",
            Operation(embedded.GetProperty("operations").EnumerateArray().ToArray(), "embed_linked").GetProperty("reason").GetString());

        var refused = Call("docxodus_images", new
        {
            sessionId, action = "replace", imageId = linked.GetProperty("id").GetString(), imageBase64 = Png(6, 7),
        });
        Assert.Equal("linked_image_read_only", refused.GetProperty("error").GetProperty("code").GetString());
        var converted = Call("docxodus_images", new
        {
            sessionId, action = "embed_linked", imageId = linked.GetProperty("id").GetString(), imageBase64 = Png(6, 7),
        });
        Assert.True(converted.GetProperty("success").GetBoolean(), converted.ToString());
        var after = Assert.Single(Images(sessionId), image => image.GetProperty("id").GetString() == linked.GetProperty("id").GetString());
        Assert.False(after.GetProperty("isLinked").GetBoolean());
        Assert.True(after.GetProperty("canMutate").GetBoolean());
        Assert.Equal(6, after.GetProperty("intrinsicWidthPixels").GetInt32());

        // Batch steps validate the new action's arguments before anything runs.
        var invalid = Call("docxodus_mutations", new
        {
            sessionId,
            steps = new[] { new { tool = "docxodus_images", args = new { action = "embed_linked", imageId = "img:body:x" } } },
        });
        Assert.False(invalid.GetProperty("success").GetBoolean());
        Assert.Equal("invalid_batch_step", invalid.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void MCP762b_TrackedReplaceIsARevisionPair_AndPreviewRollsBackMediaAndRevisions()
    {
        var sessionId = OpenSession(tracked: true);
        var embedded = Assert.Single(Images(sessionId), image => !image.GetProperty("isLinked").GetBoolean());
        Assert.True(embedded.GetProperty("canMutate").GetBoolean(), embedded.ToString());
        var imageId = embedded.GetProperty("id").GetString()!;

        var preview = Call("docxodus_mutations", new
        {
            sessionId,
            mode = "preview",
            steps = new[] { new { tool = "docxodus_images", args = new { action = "replace", imageId, imageBase64 = Png(9, 9) } } },
        });
        Assert.Equal("ok", preview.GetProperty("status").GetString());
        Assert.Equal(1, preview.GetProperty("editsApplied").GetInt32());
        // Preview restored the run, the revision markup and the media layer.
        Assert.Equal(2, Images(sessionId).Length);
        Assert.Empty(Call("docxodus_track_changes", new { sessionId, action = "list" }).GetProperty("revisions").EnumerateArray());
        Assert.Single(ImageParts(sessionId, "after-preview.docx"));

        var applied = Call("docxodus_images", new { sessionId, action = "replace", imageId, imageBase64 = Png(9, 9) });
        Assert.True(applied.GetProperty("success").GetBoolean(), applied.ToString());
        var newId = applied.GetProperty("imageId").GetString()!;
        Assert.NotEqual(imageId, newId);
        var images = Images(sessionId);
        Assert.Equal(3, images.Length);
        var deleted = Assert.Single(images, image => image.GetProperty("id").GetString() == imageId);
        Assert.False(deleted.GetProperty("canMutate").GetBoolean());
        Assert.Contains("tracked deletion", deleted.GetProperty("unsupportedReason").GetString());
        var inserted = Assert.Single(images, image => image.GetProperty("id").GetString() == newId);
        Assert.True(inserted.GetProperty("canMutate").GetBoolean());
        var revisions = Call("docxodus_track_changes", new { sessionId, action = "list" })
            .GetProperty("revisions").EnumerateArray().Select(revision => revision.GetProperty("type").GetString()).ToArray();
        Assert.Contains("insert", revisions);
        Assert.Contains("delete", revisions);
        Assert.Equal(2, ImageParts(sessionId, "after-apply.docx").Length);

        Assert.True(Call("docxodus_track_changes", new { sessionId, action = "reject_all" }).GetProperty("success").GetBoolean());
        var restored = Assert.Single(Images(sessionId), image => !image.GetProperty("isLinked").GetBoolean());
        Assert.Equal(imageId, restored.GetProperty("id").GetString());
        Assert.Equal(2, restored.GetProperty("intrinsicWidthPixels").GetInt32());
        Assert.Single(ImageParts(sessionId, "after-reject.docx"));
    }

    private JsonElement[] Images(string sessionId) =>
        Call("docxodus_images", new { sessionId, action = "list" }).GetProperty("images").EnumerateArray().ToArray();

    private static JsonElement Operation(JsonElement[] operations, string name) =>
        Assert.Single(operations, operation => operation.GetProperty("operation").GetString() == name);

    private string[] ImageParts(string sessionId, string fileName)
    {
        var path = Path.Combine(_root, fileName);
        Call("docxodus_save", new { sessionId, path });
        using var stream = new MemoryStream(File.ReadAllBytes(path));
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.ImageParts.Select(part => part.Uri.ToString()).OrderBy(uri => uri).ToArray();
    }

    private static string Png(int width, int height) =>
        Convert.ToBase64String(DocxSessionImageCoverageTests.Png(width, height));

    private JsonElement Call(string tool, object args) =>
        J(Dispatcher.Call(_store, tool, J(JsonSerializer.Serialize(args))));

    private string OpenSession(bool tracked)
    {
        var opened = J(Dispatcher.Call(_store, "docxodus_open", J(JsonSerializer.Serialize(new
        {
            path = _path,
            trackedChanges = tracked ? "render_inline" : "accept",
        }))));
        return opened.GetProperty("sessionId").GetString()!;
    }

    private static JsonElement J(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }
}
