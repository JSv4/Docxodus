// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.Verification;

namespace Docxodus.History;

public enum DocxHistoryError { StaleHead, ForeignDocument, InvalidHistory, HistoryUnavailable, TraversalLimit, RequestConflict, Contention }

/// <summary>A publication precondition or cross-record history invariant failed.</summary>
public sealed class DocxHistoryException : Exception
{
    internal DocxHistoryException(DocxHistoryError code, string message) : base(message) => Code = code;
    public DocxHistoryError Code { get; }
}

public sealed record DocxStoredVersion(HistoryBlobReference Id, DocxVersionRecord Record);
public sealed record DocxHistoryView(HistoryHead Head, DocxHistoryStateRecord State, DocxStoredVersion Version);
public sealed record DocxVersionPage(IReadOnlyList<DocxStoredVersion> Versions, HistoryBlobReference? Next);

/// <summary>
/// Provider-neutral named-version history over exact snapshots, immutable records, reversible
/// package contributions, and atomic heads. These are package/version boundary operations, not
/// a live mutation recorder. Hosts authenticate callers and own storage/retention policy.
/// </summary>
public sealed partial class DocxVersionHistory
{
    private readonly IHistoryBlobStore _blobs;
    private readonly IHistoryHeadStore _heads;
    private readonly HistoryRecordStore _records;
    private readonly HistoryRequestJournalStore _requests;
    private readonly DocxSnapshotStore _snapshots;
    private readonly int _maxSnapshotBytes;
    private readonly int _maxRecordBytes;
    private readonly PackageManifestOptions? _packageOptions;

    public DocxVersionHistory(IHistoryBlobStore blobs, IHistoryHeadStore heads,
        int maxSnapshotBytes = HistoryBlobIO.DefaultMaxBlobBytes,
        int maxRecordBytes = HistoryRecordStore.DefaultMaxRecordBytes, PackageManifestOptions? packageOptions = null)
    {
        ArgumentNullException.ThrowIfNull(blobs);
        ArgumentNullException.ThrowIfNull(heads);
        _blobs = blobs;
        _heads = heads;
        _records = new HistoryRecordStore(blobs, maxRecordBytes);
        _requests = new HistoryRequestJournalStore(blobs);
        _snapshots = new DocxSnapshotStore(blobs, maxSnapshotBytes, packageOptions);
        _maxSnapshotBytes = maxSnapshotBytes;
        _maxRecordBytes = maxRecordBytes;
        _packageOptions = packageOptions;
    }

    /// <summary>Read one consistent publication and validate its document, sequence and tip relationships.</summary>
    public async ValueTask<DocxHistoryView?> ReadAsync(string documentId, CancellationToken cancellationToken = default)
    {
        HistoryHeadCodec.Key(documentId);
        cancellationToken.ThrowIfCancellationRequested();
        var head = await _heads.ReadAsync(documentId, cancellationToken).ConfigureAwait(false);
        if (head is null) return null;
        return await ReadHeadViewAsync(documentId, head, cancellationToken).ConfigureAwait(false);
    }

    private async ValueTask<DocxHistoryView> ReadHeadViewAsync(string documentId, HistoryHead head,
        CancellationToken cancellationToken)
    {
        var state = await _records.LoadStateAsync(head.State, cancellationToken).ConfigureAwait(false);
        SameDocument(documentId, state.DocumentId);
        HistoryRequestJournalStore.ValidatePublication(documentId, head, state.Requests);
        Consistent(state.ParentPublication is null || state.ParentPublication.Revision == head.Revision - 1,
            "Publication parent must immediately precede this head.");
        Consistent(state.Sequence < head.Revision && state.Epoch <= state.Sequence, "Invalid head publication position.");
        var version = await GetVersionAsync(documentId, state.Version, cancellationToken).ConfigureAwait(false);
        Consistent(version.Record.Sequence == state.Sequence && version.Record.Snapshot == state.Snapshot,
            "Version tip disagrees with the head state.");
        if (state.Commit is not null)
        {
            var commit = await _records.LoadCommitAsync(state.Commit, cancellationToken).ConfigureAwait(false);
            SameDocument(documentId, commit.DocumentId);
            Consistent(commit.Sequence == state.Sequence && commit.Epoch == state.Epoch
                && commit.After.ContentDigest == state.Snapshot.ContentDigest, "Commit tip disagrees with the head state.");
            var committedVersion = commit.Version == state.Version ? version
                : await GetVersionAsync(documentId, commit.Version, cancellationToken).ConfigureAwait(false);
            Consistent(committedVersion.Record.Sequence == commit.Sequence && committedVersion.Record.Snapshot == commit.After
                && committedVersion.Record.Parent is not null
                && (committedVersion.Record.RestoredFrom is not null) == (commit.Kind == "restore"),
                "Commit disagrees with its recorded version.");
        }
        return new DocxHistoryView(head, state, version);
    }

    /// <summary>
    /// Capture a distinct named version using an exact head expectation (null initializes history).
    /// Changed OPC content appends an import commit with lossless effects; unchanged content adds
    /// metadata only, even when ZIP serialization differs. All blobs precede the one CAS commit point.
    /// </summary>
    public ValueTask<DocxHistoryView> CreateVersionAsync(string documentId, HistoryHead? expected,
        byte[] docxBytes, DocxVersionMetadata metadata, CancellationToken cancellationToken = default) =>
        CreateVersionCoreAsync(documentId, expected, docxBytes, metadata, null, cancellationToken);

    /// <summary>
    /// Idempotent publication. Persist requestId with the original input before submission. An
    /// identical retry returns the original view, even after later writes; changed input fails.
    /// </summary>
    public ValueTask<DocxHistoryView> CreateVersionAsync(string documentId, string requestId, HistoryHead? expected,
        byte[] docxBytes, DocxVersionMetadata metadata, CancellationToken cancellationToken = default) =>
        CreateVersionCoreAsync(documentId, expected, docxBytes, metadata,
            requestId ?? throw new ArgumentNullException(nameof(requestId)), cancellationToken);

    private async ValueTask<DocxHistoryView> CreateVersionCoreAsync(string documentId, HistoryHead? expected,
        byte[] docxBytes, DocxVersionMetadata metadata, string? requestId, CancellationToken cancellationToken)
    {
        HistoryHeadCodec.Key(documentId);
        ArgumentNullException.ThrowIfNull(docxBytes);
        cancellationToken.ThrowIfCancellationRequested();
        if (docxBytes.Length > _maxSnapshotBytes)
            throw new PackageChangeException(PackageChangeError.ResourceLimit, "Snapshot exceeds the byte limit.");
        var captured = docxBytes.ToArray();
        var capturedMetadata = HistoryRecordStore.PrepareMetadata(metadata);
        var request = requestId is null ? null : VersionRequest("create", documentId, expected,
            capturedMetadata, requestId, captured, null);
        var current = await ReadAsync(documentId, cancellationToken).ConfigureAwait(false);
        var repeated = await FindRequestAsync(documentId, current, request, cancellationToken).ConfigureAwait(false);
        if (repeated is not null) return repeated;
        if (current?.Head != expected) throw Stale();
        var prepared = await PrepareCandidateAsync(documentId, current, captured, capturedMetadata, cancellationToken).ConfigureAwait(false);
        return await PublishAsync(documentId, expected, prepared.State, prepared.Version,
            cancellationToken, request, current?.State.Requests).ConfigureAwait(false);
    }

    // Shared immutable preparation for version imports and accepted backend effects. Only the
    // caller publishes; a losing backend writer may reconcile its original intent and try again.
    private async ValueTask<(DocxHistoryStateRecord State, DocxStoredVersion Version)> PrepareCandidateAsync(
        string documentId, DocxHistoryView? current, byte[] captured, DocxVersionMetadata capturedMetadata,
        CancellationToken cancellationToken)
    {
        var snapshot = await _snapshots.CaptureAsync(captured, cancellationToken).ConfigureAwait(false);
        var changed = current is not null && current.State.Snapshot.ContentDigest != snapshot.ContentDigest;
        var sequence = checked((current?.State.Sequence ?? 0) + (changed ? 1 : 0));
        HistoryBlobReference? contribution = null;
        if (changed)
        {
            var before = await _snapshots.ExportAsync(current!.State.Snapshot, cancellationToken).ConfigureAwait(false);
            var changes = PackageChangeSet.Create(before, captured, _packageOptions);
            var manifest = await PackageChangeSetCodec.SaveAsync(changes, _blobs, cancellationToken: cancellationToken).ConfigureAwait(false);
            contribution = await HistoryBlobIO.PutBytesAsync(_blobs, manifest, cancellationToken).ConfigureAwait(false);
        }
        var version = new DocxVersionRecord
        {
            DocumentId = documentId, Metadata = capturedMetadata, Nonce = Guid.NewGuid(),
            Parent = current?.State.Version, RestoredFrom = null, Sequence = sequence, Snapshot = snapshot,
        };
        var versionId = await _records.SaveVersionAsync(version, cancellationToken).ConfigureAwait(false);
        var commitId = current?.State.Commit;
        if (changed)
        {
            commitId = await _records.SaveCommitAsync(new PackageHistoryCommitRecord
            {
                After = snapshot, Before = current!.State.Snapshot, Contribution = contribution,
                DocumentId = documentId, Epoch = current.State.Epoch, Kind = "import", Parent = commitId,
                Sequence = sequence, Version = versionId,
            }, cancellationToken).ConfigureAwait(false);
        }
        var state = new DocxHistoryStateRecord
        {
            Operation = current?.State.Operation, ParentPublication = current?.State.ParentPublication,
            Commit = commitId, DocumentId = documentId, Epoch = current?.State.Epoch ?? 0,
            InitialSnapshot = current?.State.InitialSnapshot ?? snapshot, Sequence = sequence,
            Snapshot = snapshot, Version = versionId,
        };
        return (state, new DocxStoredVersion(versionId, version));
    }

    /// <summary>
    /// Read a retained version by verified reference, including a host-retained branch. The
    /// reference is not authorization or proof of publication; default listing follows the head.
    /// </summary>
    public async ValueTask<DocxStoredVersion> GetVersionAsync(string documentId, HistoryBlobReference versionId,
        CancellationToken cancellationToken = default)
    {
        HistoryHeadCodec.Key(documentId);
        cancellationToken.ThrowIfCancellationRequested();
        var version = await _records.LoadVersionAsync(versionId, cancellationToken).ConfigureAwait(false);
        SameDocument(documentId, version.DocumentId);
        return new DocxStoredVersion(versionId, version);
    }

    /// <summary>
    /// Newest-first immutable-chain pagination. A returned continuation remains stable while
    /// newer versions publish. An explicit cursor may select a host-retained branch of this document.
    /// </summary>
    public async ValueTask<DocxVersionPage> ListVersionsAsync(string documentId, HistoryBlobReference? cursor = null,
        int limit = 25, CancellationToken cancellationToken = default)
    {
        HistoryHeadCodec.Key(documentId);
        cancellationToken.ThrowIfCancellationRequested();
        if (limit is < 1 or > 100) throw new ArgumentOutOfRangeException(nameof(limit));
        cursor ??= (await ReadAsync(documentId, cancellationToken).ConfigureAwait(false))?.State.Version;
        var versions = new List<DocxStoredVersion>();
        while (cursor is not null && versions.Count < limit)
        {
            var version = await GetVersionAsync(documentId, cursor, cancellationToken).ConfigureAwait(false);
            Consistent(versions.Count == 0 || version.Record.Sequence <= versions[^1].Record.Sequence,
                "Version chain runs forward in content sequence.");
            versions.Add(version);
            cursor = version.Record.Parent;
        }
        return new DocxVersionPage(versions.AsReadOnly(), cursor);
    }

    public async ValueTask<byte[]> ExportVersionAsync(string documentId, HistoryBlobReference versionId,
        CancellationToken cancellationToken = default) =>
        await _snapshots.ExportAsync((await GetVersionAsync(documentId, versionId, cancellationToken).ConfigureAwait(false)).Record.Snapshot,
            cancellationToken).ConfigureAwait(false);

    public async ValueTask<DocxDiffComparison> CompareVersionsAsync(string documentId, HistoryBlobReference before,
        HistoryBlobReference after, DocxDiffSettings? settings = null, CancellationToken cancellationToken = default)
    {
        var left = await GetVersionAsync(documentId, before, cancellationToken).ConfigureAwait(false);
        var right = await GetVersionAsync(documentId, after, cancellationToken).ConfigureAwait(false);
        return await _snapshots.CompareAsync(left.Record.Snapshot, right.Record.Snapshot, settings, cancellationToken).ConfigureAwait(false);
    }

    private async ValueTask<DocxHistoryView> PublishAsync(string documentId, HistoryHead? expected,
        DocxHistoryStateRecord state, DocxStoredVersion version, CancellationToken cancellationToken,
        HistoryRequestIdentity? request = null, HistoryRequestJournal? previousRequests = null)
    {
        state = state with
        {
            ParentPublication = state.Operation is null ? null : expected,
            Requests = await _requests.AdvanceAsync(documentId, expected, previousRequests,
                request, cancellationToken).ConfigureAwait(false),
        };
        var stateId = await _records.SaveStateAsync(state, cancellationToken).ConfigureAwait(false);
        var head = await _heads.TryAdvanceAsync(documentId, expected, stateId, cancellationToken).ConfigureAwait(false);
        // Do not throw cancellation after publication: the CAS is the definitive commit point.
        if (head is not null) return new DocxHistoryView(head, state, version);
        // A concurrent identical retry may have won the CAS. Resolve its durable receipt rather
        // than reporting that successfully published request stale. Unknown storage exceptions
        // propagate: the caller retries the same identity to resolve a possibly lost acknowledgement.
        var repeated = await FindRequestAsync(documentId,
            await ReadAsync(documentId, cancellationToken).ConfigureAwait(false), request, cancellationToken).ConfigureAwait(false);
        return repeated ?? throw Stale();
    }

    private static DocxHistoryException Stale() => new(DocxHistoryError.StaleHead, "History head changed; recapture the expected head before publishing.");
    private static void SameDocument(string expected, string actual)
    {
        if (!string.Equals(expected, actual, StringComparison.Ordinal))
            throw new DocxHistoryException(DocxHistoryError.ForeignDocument, "History record belongs to a different document.");
    }
    private static void Consistent(bool condition, string message)
    {
        if (!condition) throw new DocxHistoryException(DocxHistoryError.InvalidHistory, message);
    }
}
