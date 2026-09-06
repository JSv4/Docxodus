#nullable enable

using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.History;
using Xunit;

namespace Docxodus.Tests;

public sealed class DocxBackendReconciliationTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DelayedConflictingStreamPreservesBothCompatibleEditsAndRecoverableContenders(bool filesystem)
    {
        using var store = new HistoryFaultHarness(filesystem);
        var history = new DocxVersionHistory(store, store);
        var original = Package("abcdefghij café 📝");
        var first = await history.CreateVersionAsync("doc", "initial", null, original, Metadata("initial"));
        var a = Text("a", first.Head, 2, 0, "A");
        var b = Text("b", first.Head, 2, 0, "B");
        var c = Text("c", first.Head, 6, 2, "");
        var d = Text("d", first.Head, 7, 1, "X");
        var ra = await history.SubmitOperationAsync("doc", a);
        var rb = await history.SubmitOperationAsync("doc", b);
        var rc = await history.SubmitOperationAsync("doc", c);
        var rd = await history.SubmitOperationAsync("doc", d);
        Assert.Equal("accepted", rb.Operation.Record.Status);
        Assert.Equal(3, rb.Operation.Record.AppliedText!.Offset);
        Assert.Equal("OverlappingText", rd.Operation.Record.Conflict);
        Assert.Equal(rc.View.Version.Id, rd.View.Version.Id);
        Assert.Equal(rc.View.State.Sequence, rd.View.State.Sequence);
        Assert.Equal(rc.View.Head.Revision + 1, rd.View.Head.Revision);
        Assert.Equal("abABcdefij café 📝", BodyText(await history.ExportVersionAsync("doc", rd.View.Version.Id)));
        Assert.Equal("abcdefgXij café 📝", BodyText(await history.ExportOperationProposalAsync("doc", rd.Operation.Id)));
        var resolution = Text("resolve-d", rd.View.Head, 8, 0, "R") with { Resolves = rd.Operation.Id };
        var resolved = await history.SubmitOperationAsync("doc", resolution);
        Assert.Equal("accepted", resolved.Operation.Record.Status);
        Assert.Equal("abABcdefRij café 📝", BodyText(await history.ExportVersionAsync("doc", resolved.View.Version.Id)));
        // Reopen, delayed duplicate, and a competing new resolution do not apply the same work twice.
        history = new DocxVersionHistory(store, store);
        Assert.Equal(ra.View.Head, (await history.SubmitOperationAsync("doc", a)).View.Head);
        Assert.Equal(rd.View.Head, (await history.SubmitOperationAsync("doc", d)).View.Head);
        var repeatedResolution = await history.SubmitOperationAsync("doc", resolution with { RequestId = "other-resolution" });
        Assert.Equal("AlreadyResolved", repeatedResolution.Operation.Record.Conflict);
        var outcomes = await history.ReadOperationsSinceAsync("doc", first.Head);
        Assert.Equal(new[] { "a", "b", "c", "d", "resolve-d", "other-resolution" }, outcomes.Operations.Select(o => o.Input.Request.RequestId));
        Assert.Equal(new long[] { 2, 3, 4, 5, 6, 7 }, outcomes.Operations.Select(o => o.Record.Revision));
        Assert.Empty((await history.ReadOperationsSinceAsync("doc", repeatedResolution.View.Head)).Operations);
        Assert.Equal(5, (await history.ListVersionsAsync("doc")).Versions.Count);
        var snapshots = new[] { first, ra.View, rb.View, rc.View, resolved.View };
        foreach (var view in snapshots)
        {
            var exact = await history.ExportVersionAsync("doc", view.Version.Id);
            SameEntries(exact, await history.ReplayAsync("doc", view.State.Sequence));
            Untouched(original, exact, "word/document.xml");
            NoNewValidationErrors(original, exact);
        }
        Assert.Equal(original, await history.ExportVersionAsync("doc", first.Version.Id));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task RealCASLoserReconcilesDifferentIntentAndIdenticalRetriesReturnOneOutcome(bool filesystem, bool rightFirst)
    {
        using var store = new HistoryFaultHarness(filesystem);
        var history = new DocxVersionHistory(store, store);
        var first = await history.CreateVersionAsync("doc", null, Package("abcd"), Metadata("initial"));
        var a = Text("a", first.Head, 2, 0, "A");
        var b = Text("b", first.Head, 2, 0, "B");
        store.Arm(HistoryFaultHarness.Fault.None);
        store.HoldCompetingPublications(rightFirst ? "right" : "left");
        DocxOperationResult[] results;
        try
        {
            results = await Task.WhenAll(
                Task.Run(() => store.AsContenderAsync("left", async () => await new DocxVersionHistory(store, store).SubmitOperationAsync("doc", a))),
                Task.Run(() => store.AsContenderAsync("right", async () => await new DocxVersionHistory(store, store).SubmitOperationAsync("doc", b))));
        }
        finally { store.ReleaseCompetingPublications(); }
        Assert.All(results, r => Assert.Equal("accepted", r.Operation.Record.Status));
        Assert.Equal(3, store.Publications); // Both attempt the old CAS; the loser then reconciles and commits.
        var current = (await history.ReadAsync("doc"))!;
        Assert.Equal(rightFirst ? "abBAcd" : "abABcd", BodyText(await history.ExportVersionAsync("doc", current.Version.Id)));
        var shared = Text("shared", current.Head, 0, 0, "!");
        store.Arm(HistoryFaultHarness.Fault.None); store.HoldCompetingPublications(rightFirst ? "right" : "left");
        try
        {
            results = await Task.WhenAll(
                Task.Run(() => store.AsContenderAsync("left", async () => await new DocxVersionHistory(store, store).SubmitOperationAsync("doc", shared))),
                Task.Run(() => store.AsContenderAsync("right", async () => await new DocxVersionHistory(store, store).SubmitOperationAsync("doc", shared))));
        }
        finally { store.ReleaseCompetingPublications(); }
        Assert.Equal(2, store.Publications);
        Assert.Equal(results[0].View.Head, results[1].View.Head);
        Assert.Equal(results[0].Operation.Id, results[1].Operation.Id);
        Assert.Equal(3, (await history.ReadOperationsSinceAsync("doc", null)).Operations.Count);
    }

    [Fact]
    public async Task DisjointPackagePartsMergeReadDependenciesConflictAndDiscardIsAuditable()
    {
        using var store = new HistoryFaultHarness();
        var history = new DocxVersionHistory(store, store);
        var original = Package("body");
        var first = await history.CreateVersionAsync("doc", null, original, Metadata("initial"));
        var header = ReplacePartText(original, "word/header1.xml", "header", "new header");
        var footnote = ReplacePartText(original, "word/footnotes.xml", "note", "new note");
        var headerResult = await history.SubmitOperationAsync("doc", PackageRequest("header", first.Head), header);
        var noteResult = await history.SubmitOperationAsync("doc", PackageRequest("note", first.Head), footnote);
        Assert.Equal("accepted", noteResult.Operation.Record.Status);
        var merged = await history.ExportVersionAsync("doc", noteResult.View.Version.Id);
        Assert.Contains("new header", Encoding.UTF8.GetString(Entries(merged)["word/header1.xml"]));
        Assert.Contains("new note", Encoding.UTF8.GetString(Entries(merged)["word/footnotes.xml"]));
        Untouched(original, merged, "word/header1.xml", "word/footnotes.xml");
        NoNewValidationErrors(original, merged);
        var dependency = PackageRequest("dependency", first.Head) with { ReadParts = new[] { "/word/header1.xml" } };
        var conflict = await history.SubmitOperationAsync("doc", dependency, footnote);
        Assert.Equal("ReadDependencyChanged", conflict.Operation.Record.Conflict);
        Assert.Equal(footnote, await history.ExportOperationProposalAsync("doc", conflict.Operation.Id));
        var discard = new DocxOperationRequest
        {
            RequestId = "discard", Base = conflict.View.Head, Kind = "discard", Metadata = Metadata("discard"), Resolves = conflict.Operation.Id,
        };
        var discarded = await history.SubmitOperationAsync("doc", discard);
        Assert.Equal("accepted", discarded.Operation.Record.Status);
        Assert.Null(discarded.Operation.Record.ContentCommit);
        Assert.Equal(noteResult.View.Version.Id, discarded.View.Version.Id);
        Assert.Empty((await history.ReadChangesSinceAsync("doc", noteResult.View.Head)).Entries);
        // Same desired package content is accepted without manufacturing a version or content edit.
        var noOp = await history.SubmitOperationAsync("doc", PackageRequest("same-header", first.Head), header);
        Assert.Equal("accepted", noOp.Operation.Record.Status); Assert.Null(noOp.Operation.Record.ContentCommit);
        Assert.Equal(discarded.View.Version.Id, noOp.View.Version.Id);
        Assert.Equal(headerResult.View.Head, (await history.SubmitOperationAsync("doc", PackageRequest("header", first.Head), header)).View.Head);
    }

    [Fact]
    public async Task EpochAndUnknownStructuralChangesAreExplicitBoundariesAndIdsBindOriginalInput()
    {
        using var store = new HistoryFaultHarness();
        var history = new DocxVersionHistory(store, store);
        var original = Package("abcd");
        var first = await history.CreateVersionAsync("doc", null, original, Metadata("initial"));
        var stale = Text("stale", first.Head, 1, 0, "X");
        var imported = await history.CreateVersionAsync("doc", first.Head,
            ReplacePartText(original, "word/document.xml", "abcd", "new paragraph"), Metadata("external edit"));
        var unknown = await history.SubmitOperationAsync("doc", stale);
        Assert.Equal("UnknownTextChange", unknown.Operation.Record.Conflict);
        var reset = await history.RestoreVersionAsync("doc", unknown.View.Head, first.Version.Id, Metadata("restore"));
        var epoch = await history.SubmitOperationAsync("doc", stale with { RequestId = "across-reset" });
        Assert.Equal("EpochChanged", epoch.Operation.Record.Conflict);
        Assert.Equal(reset.Version.Id, epoch.View.Version.Id);
        Assert.Equal(DocxHistoryError.RequestConflict, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await history.SubmitOperationAsync("doc", stale with { Base = epoch.View.Head }))).Code);
        Assert.Equal(unknown.View.Head, (await history.SubmitOperationAsync("doc", stale)).View.Head);
        Assert.Equal(imported.Version.Id, unknown.View.Version.Id);
    }

    [Fact]
    public async Task InvalidUnicodeOrBaseInputsNeverPublishAndCancellationAfterCASDoesNotEraseSuccess()
    {
        using var store = new HistoryFaultHarness();
        var history = new DocxVersionHistory(store, store);
        var first = await history.CreateVersionAsync("doc", null, Package("a📝b"), Metadata("initial"));
        foreach (var invalid in new[] { Text("split", first.Head, 2, 0, "x"), Text("range", first.Head, 100, 0, "x") })
            await Assert.ThrowsAsync<PackageChangeException>(async () => await history.SubmitOperationAsync("doc", invalid));
        Assert.Equal(first.Head, (await history.ReadAsync("doc"))!.Head);
        store.Cancellation = new CancellationTokenSource(); store.Arm(HistoryFaultHarness.Fault.CancelAfterHead);
        var request = Text("committed", first.Head, 3, 0, "!");
        var result = await history.SubmitOperationAsync("doc", request, cancellationToken: store.Cancellation.Token);
        Assert.True(store.Cancellation.IsCancellationRequested);
        store.Arm(HistoryFaultHarness.Fault.None);
        Assert.Equal(result.View.Head, (await history.SubmitOperationAsync("doc", request)).View.Head);
        Assert.Equal("a📝!b", BodyText(await history.ExportVersionAsync("doc", result.View.Version.Id)));
    }

    internal static DocxVersionMetadata Metadata(string label) => DocxVersionRequestTests.Metadata(label);
    internal static DocxOperationRequest Text(string id, HistoryHead baseline, int offset, int delete, string insert) => new()
    { RequestId = id, Base = baseline, Kind = "text", Metadata = Metadata(id), Text = new DocxTextSplice("/word/document.xml", 0, offset, delete, insert) };
    internal static DocxOperationRequest PackageRequest(string id, HistoryHead baseline) => new()
    { RequestId = id, Base = baseline, Kind = "package", Metadata = Metadata(id) };

    internal static byte[] Package(string text)
    {
        using var stream = new MemoryStream(); stream.Write(DocxSession.CreateBlankDocxBytes());
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var main = doc.MainDocumentPart!;
            var header = main.AddNewPart<HeaderPart>(); header.Header = new Header(new Paragraph(new Run(new Text("header"))));
            var notes = main.AddNewPart<FootnotesPart>();
            notes.Footnotes = new Footnotes(new Footnote(new Paragraph(new Run(new Text("note")))) { Id = 1 });
            var comments = main.AddNewPart<WordprocessingCommentsPart>();
            comments.Comments = new Comments(new Comment(new Paragraph(new Run(new Text("review"))))
            { Id = "0", Author = "Reviewer", Date = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc) });
            main.Document.Body = new Body(new Paragraph(new Run(new Text(text)), new Run(new FootnoteReference { Id = 1 })),
                new Paragraph(new CommentRangeStart { Id = "0" }, new Run(new Text("commented")), new CommentRangeEnd { Id = "0" },
                    new Run(new CommentReference { Id = "0" })),
                new SectionProperties(new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) }));
            var opaque = main.AddCustomXmlPart("application/octet-stream");
            using var payload = new MemoryStream(Enumerable.Range(0, 257).Select(i => (byte)(i * 37)).ToArray()); opaque.FeedData(payload);
        }
        return stream.ToArray();
    }
    internal static string BodyText(byte[] package) => XDocument.Parse(Encoding.UTF8.GetString(Entries(package)["word/document.xml"]))
        .Descendants(XName.Get("t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")).First().Value;
    internal static byte[] ReplacePartText(byte[] package, string name, string before, string after)
    {
        using var stream = new MemoryStream(); stream.Write(package);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(name)!;
            using var content = entry.Open(); using var reader = new StreamReader(content, leaveOpen: true);
            var xml = XDocument.Parse(reader.ReadToEnd(), LoadOptions.PreserveWhitespace);
            foreach (var text in xml.DescendantNodes().OfType<XText>().Where(t => t.Value == before)) text.Value = after;
            content.Position = 0; content.SetLength(0); content.Write(Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)));
        }
        return stream.ToArray();
    }
    internal static SortedDictionary<string, byte[]> Entries(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        var entries = new SortedDictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries)
        { using var input = entry.Open(); using var output = new MemoryStream(); input.CopyTo(output); entries.Add(entry.FullName, output.ToArray()); }
        return entries;
    }
    internal static void SameEntries(byte[] expected, byte[] actual)
    {
        var a = Entries(expected); var b = Entries(actual); Assert.Equal(a.Keys, b.Keys);
        foreach (var pair in a) Assert.Equal(pair.Value, b[pair.Key]);
    }
    internal static void Untouched(byte[] before, byte[] after, params string[] changed)
    {
        var a = Entries(before); var b = Entries(after); Assert.Equal(a.Keys, b.Keys);
        foreach (var pair in a.Where(p => !changed.Contains(p.Key))) Assert.Equal(pair.Value, b[pair.Key]);
    }
    internal static void NoNewValidationErrors(byte[] before, byte[] after)
    {
        static HashSet<string> Errors(byte[] bytes)
        {
            using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc)
                .Select(e => e.Id + "|" + e.Part?.Uri + "|" + e.Path?.XPath).ToHashSet();
        }
        var baseline = Errors(before);
        Assert.Empty(Errors(after).Except(baseline));
    }
}
