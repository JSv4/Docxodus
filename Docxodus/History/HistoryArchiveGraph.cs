// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Compression;
using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>
/// Typed closure of ONE captured head. Never enumerates a store and never treats arbitrary
/// metadata/digest strings as references. Caches records, not snapshot/payload byte arrays.
/// </summary>
internal sealed class HistoryArchiveGraph
{
    private enum Role { State, Version, Commit, Index, Decision, Input, Contribution, Snapshot, Payload }
    private readonly string _documentId;
    private readonly HistoryHead _head;
    private readonly IHistoryBlobStore _blobs;
    private readonly DocxHistoryArchiveLimits _limits;
    private readonly PackageManifestOptions? _packageOptions;
    private readonly HistoryRecordStore _records;
    private readonly HistoryRequestJournalStore _requests;
    private readonly DocxOperationStore _operations;
    private readonly Dictionary<string, HistoryBlobReference> _inventory = new(StringComparer.Ordinal);
    private readonly HashSet<(HistoryBlobReference, Role)> _visited = new();
    private readonly Queue<(HistoryBlobReference Reference, Role Role)> _pending = new();
    private readonly HashSet<HistoryHead> _heads = new();
    private readonly Dictionary<HistoryBlobReference, DocxHistoryStateRecord> _states = new();
    private readonly Dictionary<HistoryBlobReference, DocxVersionRecord> _versions = new();
    private readonly Dictionary<HistoryBlobReference, PackageHistoryCommitRecord> _commits = new();
    private readonly Dictionary<HistoryBlobReference, HistoryRequestIndexNode> _indexes = new();
    private readonly Dictionary<HistoryBlobReference, DocxOperationRecord> _decisions = new();
    private readonly Dictionary<HistoryBlobReference, DocxOperationInput> _inputs = new();
    private readonly Dictionary<HistoryBlobReference, PackageChangeSetCodec.Manifest> _contributions = new();
    private readonly Dictionary<HistoryBlobReference, VerificationDigest?> _snapshotDigests = new();
    private readonly Dictionary<HistoryBlobReference, long> _snapshotSizes = new();
    private long _totalBytes, _metadataBytes, _validationBytes, _expandedBytes;
    private int _edges;

    private HistoryArchiveGraph(string documentId, HistoryHead head, IHistoryBlobStore blobs,
        DocxHistoryArchiveLimits limits, PackageManifestOptions? packageOptions)
    {
        HistoryHeadCodec.Key(documentId); limits.Validate();
        _documentId = documentId; _head = head; _blobs = blobs; _limits = limits; _packageOptions = packageOptions;
        _records = new HistoryRecordStore(blobs, limits.MaxRecordBytes);
        _requests = new HistoryRequestJournalStore(blobs); _operations = new DocxOperationStore(blobs);
    }

    internal DocxHistoryView View => ViewAt(_head);
    internal IReadOnlyList<HistoryBlobReference> Inventory => Array.AsReadOnly(_inventory.Values
        .OrderBy(r => r.Digest.Value, StringComparer.Ordinal).ToArray());

    internal static async ValueTask<HistoryArchiveGraph> LoadAsync(string documentId, HistoryHead head,
        IHistoryBlobStore blobs, DocxHistoryArchiveLimits? limits = null,
        PackageManifestOptions? packageOptions = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(blobs); ArgumentNullException.ThrowIfNull(head);
        var graph = new HistoryArchiveGraph(documentId, head, blobs, limits ?? new(), packageOptions);
        graph.Head(head);
        await graph.VisitAsync(cancellationToken).ConfigureAwait(false);
        graph.Validate(cancellationToken);
        await graph.ValidateEffectsAsync(cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return graph;
    }

    private void Spend()
    {
        if (_edges++ >= _limits.MaxEdges) throw new DocxHistoryException(DocxHistoryError.TraversalLimit,
            "Archive graph edge budget reached.");
    }

    private void Add(HistoryBlobReference? reference, Role role)
    {
        if (reference is null) return;
        Spend();
        HistoryBlobIO.Validate(reference, role == Role.Snapshot ? Math.Min(_limits.MaxBlobBytes, _limits.MaxSnapshotBytes) : _limits.MaxBlobBytes);
        if (_inventory.TryGetValue(reference.Digest.Value, out var known))
            Require(known == reference, "One digest has conflicting lengths.");
        else
        {
            Budget(_inventory.Count < _limits.MaxBlobs && reference.Length <= _limits.MaxTotalBlobBytes - _totalBytes,
                "Archive blob inventory exceeds its budget.");
            _totalBytes += reference.Length; _inventory.Add(reference.Digest.Value, reference);
        }
        if (!_visited.Add((reference, role))) return;
        if (role is not (Role.Payload or Role.Snapshot))
        {
            Budget(reference.Length <= _limits.MaxMetadataBytes - _metadataBytes, "Archive metadata exceeds its budget.");
            _metadataBytes += reference.Length;
        }
        _pending.Enqueue((reference, role));
    }

    private void Head(HistoryHead? head)
    {
        if (head is null) return;
        Require(head.Revision > 0 && head.Revision <= _head.Revision, "Invalid referenced publication revision.");
        Add(head.State, Role.State); _heads.Add(head);
    }

    private void Snapshot(DocxSnapshotReference snapshot)
    {
        HistoryBlobIO.Validate(new HistoryBlobReference(snapshot.ContentDigest, 0), 0);
        if (_snapshotDigests.TryGetValue(snapshot.Blob, out var known) && known is not null)
            Require(known == snapshot.ContentDigest, "One snapshot has conflicting OPC content digests.");
        _snapshotDigests[snapshot.Blob] = snapshot.ContentDigest;
        Add(snapshot.Blob, Role.Snapshot);
    }

    private void SameDocument(string actual)
    {
        if (_documentId != actual) throw new DocxHistoryException(DocxHistoryError.ForeignDocument,
            "Archive graph crosses document identity.");
    }

    private async ValueTask VisitAsync(CancellationToken ct)
    {
        while (_pending.TryDequeue(out var item))
        {
            ct.ThrowIfCancellationRequested();
            var r = item.Reference;
            ChargeRead(r.Length);
            switch (item.Role)
            {
                case Role.State:
                    var s = await _records.LoadStateAsync(r, ct).ConfigureAwait(false);
                    SameDocument(s.DocumentId); _states.Add(r, s);
                    Head(s.ParentPublication); Add(s.Version, Role.Version); Add(s.Commit, Role.Commit);
                    Add(s.Operation, Role.Decision); Add(s.Requests?.Index, Role.Index);
                    Snapshot(s.InitialSnapshot); Snapshot(s.Snapshot);
                    break;
                case Role.Version:
                    var v = await _records.LoadVersionAsync(r, ct).ConfigureAwait(false);
                    SameDocument(v.DocumentId); _versions.Add(r, v);
                    Add(v.Parent, Role.Version); Add(v.RestoredFrom, Role.Version); Snapshot(v.Snapshot);
                    break;
                case Role.Commit:
                    var c = await _records.LoadCommitAsync(r, ct).ConfigureAwait(false);
                    SameDocument(c.DocumentId); _commits.Add(r, c);
                    Add(c.Parent, Role.Commit); Add(c.Version, Role.Version); Add(c.Contribution, Role.Contribution);
                    Snapshot(c.Before); Snapshot(c.After);
                    break;
                case Role.Index:
                    var n = await _requests.LoadAsync(_documentId, r, ct).ConfigureAwait(false);
                    _indexes.Add(r, n); Add(n.Zero, Role.Index); Add(n.One, Role.Index); Head(n.Receipt?.Head);
                    break;
                case Role.Decision:
                    var d = await _operations.LoadDecisionAsync(r, ct).ConfigureAwait(false);
                    SameDocument(d.DocumentId); _decisions.Add(r, d);
                    Add(d.Input, Role.Input); Add(d.Parent, Role.Decision); Head(d.Before);
                    Add(d.Version, Role.Version); Add(d.ContentCommit, Role.Commit);
                    Snapshot(d.ProposedSnapshot); Snapshot(d.AfterSnapshot);
                    break;
                case Role.Input:
                    var input = await _operations.LoadInputAsync(r, ct).ConfigureAwait(false);
                    SameDocument(input.DocumentId); _inputs.Add(r, input);
                    Head(input.Request.Base); Add(input.Request.Resolves, Role.Decision);
                    if (input.Candidate is { } candidate)
                    {
                        _snapshotDigests.TryAdd(candidate, null); Add(candidate, Role.Snapshot);
                    }
                    break;
                case Role.Contribution:
                    var manifest = PackageChangeSetCodec.ReadInventory(await HistoryBlobIO.ReadBytesAsync(_blobs, r,
                        _limits.MaxManifestBytes, ct).ConfigureAwait(false), ChangeLimits());
                    _contributions.Add(r, manifest);
                    foreach (var payload in manifest.Payloads) Add(payload, Role.Payload);
                    break;
                case Role.Snapshot:
                    var bytes = await HistoryBlobIO.ReadBytesAsync(_blobs, r, _limits.MaxBlobBytes, ct).ConfigureAwait(false);
                    var inspected = InspectSnapshot(bytes);
                    var digest = inspected.OrderedOpcContentDigest!;
                    Require(_snapshotDigests[r] is null || _snapshotDigests[r] == digest, "Snapshot OPC content digest disagrees.");
                    // Retain the inspected identity, not the potentially large package.
                    _snapshotDigests[r] = digest;
                    _snapshotSizes[r] = inspected.Entries.Sum(e => e.Size);
                    break;
                case Role.Payload:
                    using (var content = await _blobs.OpenReadAsync(r, ct).ConfigureAwait(false))
                    {
                        if (content is null) throw new PackageChangeException(PackageChangeError.PayloadMissing, "Archive payload is missing.");
                        await HistoryBlobIO.CopyVerifiedAsync(r, content, Stream.Null, ct).ConfigureAwait(false);
                    }
                    break;
            }
        }
    }

    private void Validate(CancellationToken ct)
    {
        // Iterative ancestry numbering detects cycles without recursion/stack exhaustion.
        var depths = new Dictionary<HistoryBlobReference, long>();
        foreach (var id in _versions.Keys)
        {
            ct.ThrowIfCancellationRequested();
            var path = new List<HistoryBlobReference>(); var seen = new HashSet<HistoryBlobReference>();
            HistoryBlobReference? cursor = id;
            while (cursor is not null && !depths.ContainsKey(cursor))
            {
                Spend(); Require(seen.Add(cursor), "Cyclic version ancestry.");
                path.Add(cursor); cursor = _versions[cursor].Parent;
            }
            var depth = cursor is null ? 0 : depths[cursor];
            for (var i = path.Count - 1; i >= 0; i--) depths.Add(path[i], ++depth);
        }
        var contentVersions = new Dictionary<HistoryBlobReference, HistoryBlobReference>();
        foreach (var (id, _) in depths.OrderBy(pair => pair.Value))
        {
            var v = _versions[id];
            contentVersions.Add(id, v.Parent is not null && _versions[v.Parent].Sequence == v.Sequence
                ? contentVersions[v.Parent] : id);
        }
        foreach (var v in _versions.Values)
        {
            ct.ThrowIfCancellationRequested();
            if (v.Parent is null) Require(v.Sequence == 0 && v.RestoredFrom is null, "Invalid initial version.");
            else
            {
                var p = _versions[v.Parent];
                Require(v.Sequence >= p.Sequence && v.Sequence - p.Sequence <= 1, "Discontinuous version ancestry.");
                if (v.Sequence == p.Sequence) Require(v.RestoredFrom is null
                    && v.Snapshot.ContentDigest == p.Snapshot.ContentDigest, "Metadata version changed content.");
            }
            if (v.RestoredFrom is not null) Require(_versions[v.RestoredFrom].Snapshot == v.Snapshot, "Restore target disagrees.");
        }
        foreach (var c in _commits.Values)
        {
            ct.ThrowIfCancellationRequested();
            var v = _versions[c.Version];
            Require(v.Sequence == c.Sequence && v.Snapshot == c.After && v.Parent is not null
                && (v.RestoredFrom is not null) == (c.Kind == "restore"), "Commit version disagrees.");
            var previousVersion = _versions[v.Parent!];
            Require(previousVersion.Sequence == c.Sequence - 1 && previousVersion.Snapshot == c.Before, "Commit before-version disagrees.");
            var previous = c.Parent is null ? null : _commits[c.Parent];
            Require((previous?.Sequence ?? 0) == c.Sequence - 1
                && (previous?.Epoch ?? 0) == c.Epoch - (c.Kind == "restore" ? 1 : 0)
                && (previous is null || previous.After.ContentDigest == c.Before.ContentDigest), "Discontinuous commit ancestry.");
            Require(previous is null || previous.Version == contentVersions[v.Parent!], "Commit ancestry follows a different version branch.");
            if (c.Contribution is not null)
            {
                var effects = _contributions[c.Contribution];
                Require(effects.Before == c.Before.ContentDigest && effects.After == c.After.ContentDigest, "Contribution endpoints disagree.");
            }
        }
        // Memoized max receipt revision validates every root context without repeatedly expanding
        // the persistent radix DAG. Every incoming branch edge is still checked individually.
        var indexMax = new Dictionary<HistoryBlobReference, long>();
        var indexCount = new Dictionary<HistoryBlobReference, long>();
        foreach (var (id, node) in _indexes.OrderByDescending(pair => pair.Value.Bit))
        {
            ct.ThrowIfCancellationRequested();
            if (node.Receipt is { } receipt)
            {
                Require(_states[receipt.Head.State].Requests?.Current == receipt.Request, "Receipt does not bind its original publication.");
                indexMax.Add(id, receipt.Head.Revision);
                indexCount.Add(id, 1);
            }
            else
            {
                HistoryRequestJournalStore.ValidateArchiveChild(node, _indexes[node.Zero!], false);
                HistoryRequestJournalStore.ValidateArchiveChild(node, _indexes[node.One!], true);
                indexMax.Add(id, Math.Max(indexMax[node.Zero!], indexMax[node.One!]));
                indexCount.Add(id, indexCount[node.Zero!] + indexCount[node.One!]);
            }
        }
        // Every leaf is unique by the validated radix prefix, points at one earlier inline
        // identity, and is revision-bounded. Cardinality then proves each journal retains ALL
        // known earlier identities without expanding the shared tree once per publication.
        var expectedReceipts = new Dictionary<HistoryHead, long>();
        var requestIds = new HashSet<string>(StringComparer.Ordinal);
        long receiptCount = 0;
        foreach (var head in _heads.OrderBy(h => h.Revision))
        {
            ct.ThrowIfCancellationRequested(); expectedReceipts.Add(head, receiptCount);
            if (_states[head.State].Requests?.Current is { } identity)
            {
                Require(requestIds.Add(identity.Id), "A request identity was published more than once.");
                receiptCount++;
            }
        }
        foreach (var head in _heads)
        {
            ct.ThrowIfCancellationRequested();
            var s = _states[head.State]; var v = _versions[s.Version];
            HistoryRequestJournalStore.ValidatePublication(_documentId, head, s.Requests);
            Require(s.Sequence < head.Revision && s.Epoch <= s.Sequence
                && v.Snapshot == s.Snapshot && v.Sequence == s.Sequence, "Head/version position disagrees.");
            if (s.Requests?.Index is { } index) Require(indexMax[index] < head.Revision, "Journal indexes a future receipt.");
            Require((s.Requests?.Index is { } root ? indexCount[root] : 0) == expectedReceipts[head],
                "Journal dropped earlier retained retry identities.");
            if (s.Commit is { } commit)
            {
                var c = _commits[commit];
                Require(c.Sequence == s.Sequence && c.Epoch == s.Epoch && c.After.ContentDigest == s.Snapshot.ContentDigest,
                    "Head/commit position disagrees.");
                Require(c.Version == contentVersions[s.Version], "Head commit follows a different version branch.");
            }
            else Require(s.Sequence == 0 && s.Epoch == 0 && s.InitialSnapshot.ContentDigest == s.Snapshot.ContentDigest,
                "Empty content history disagrees with initial snapshot.");
            if (s.ParentPublication is { } parent)
            {
                Require(parent.Revision == head.Revision - 1, "Publication parent revision disagrees.");
                DocxVersionHistory.ValidatePublicationEdge(ViewAt(head), ViewAt(parent));
                var before = _states[parent.State];
                if (s.Sequence != before.Sequence)
                {
                    var c = _commits[s.Commit!];
                    Require(c.Parent == before.Commit && c.Before == before.Snapshot,
                        "Content publication does not extend its parent's commit.");
                }
                if (s.Operation != _states[parent.State].Operation) ValidateDecisionPublication(head, s.Operation!);
            }
            else Require(s.Operation is null && depths[s.Version] == head.Revision, "Legacy publication/version depth disagrees.");
        }
        ValidatePublishedHeads(depths, ct);
        // Every referenced head must share the root checkpoint, including retained branches.
        Require(_states.Values.All(s => s.InitialSnapshot == _states[_head.State].InitialSnapshot), "History has multiple initial checkpoints.");
        foreach (var c in _commits.Values.Where(c => c.Parent is null))
            Require(c.Before.ContentDigest == _states[_head.State].InitialSnapshot.ContentDigest, "Commit root is not the initial checkpoint.");
        var resolved = new HashSet<HistoryBlobReference>();
        foreach (var (id, d) in _decisions)
        {
            ct.ThrowIfCancellationRequested();
            var input = _inputs[d.Input];
            Require(DocxOperationStore.Reference(DocxOperationStore.EncodeInput(input)) == d.Input, "Noncanonical operation input.");
            Require(input.Request.Base.Revision <= d.Before.Revision && (input.Candidate is null || input.Candidate == d.ProposedSnapshot.Blob),
                "Operation base or preserved candidate disagrees.");
            Require(d.Parent == _states[d.Before.State].Operation, "Decision parent disagrees with its before publication.");
            if (d.Parent is not null) Require(_decisions[d.Parent].Revision < d.Revision, "Decision ancestry is not decreasing.");
            if (input.Request.Resolves is { } resolves) Require(_decisions[resolves].Status == "conflict"
                && _decisions[resolves].Revision <= input.Request.Base.Revision, "Resolution target is not an earlier conflict.");
            if (input.Request.Resolves is { } acceptedResolution && d.Status == "accepted")
                Require(resolved.Add(acceptedResolution), "A conflict was resolved more than once.");
            Require(input.Request.Kind != "discard" || d.ContentCommit is null, "Discard cannot have content effects.");
            Require(input.Request.Kind != "text" || d.ContentCommit is null || d.AppliedText is not null,
                "Accepted text effects require their recorded text map.");
            if (d.AppliedText is { } applied) Require(input.Request.Kind == "text" && input.Request.Text is { } original
                && applied == original with { Offset = applied.Offset }, "Mapped text changed the original intent.");
        }
    }

    private void ValidatePublishedHeads(Dictionary<HistoryBlobReference, long> depths, CancellationToken ct)
    {
        // V3 has exact publication links; V1/V2 proves ancestry through one version per
        // publication. Receipt and operation-base heads must belong to this pinned chain,
        // while a restored-from *version* may deliberately retain a separate branch.
        var published = new HashSet<HistoryHead>();
        var decisions = new HashSet<HistoryBlobReference>();
        var anchor = _head;
        while (_states[anchor.State].ParentPublication is { } parent)
        {
            ct.ThrowIfCancellationRequested(); Spend();
            Require(published.Add(anchor), "Cyclic publication ancestry.");
            if (_states[anchor.State].Operation != _states[parent.State].Operation)
                decisions.Add(_states[anchor.State].Operation!);
            anchor = parent;
        }
        published.Add(anchor);
        var legacyVersions = new HashSet<HistoryBlobReference>();
        HistoryBlobReference? cursor = _states[anchor.State].Version;
        while (cursor is not null)
        {
            ct.ThrowIfCancellationRequested(); Spend();
            legacyVersions.Add(cursor);
            var version = _versions[cursor];
            if (version.Parent is null) Require(version.Snapshot == _states[_head.State].InitialSnapshot,
                "Published version root is not the initial checkpoint.");
            cursor = version.Parent;
        }
        var revisions = new Dictionary<long, HistoryHead>();
        foreach (var head in _heads)
        {
            ct.ThrowIfCancellationRequested();
            Require(!revisions.TryGetValue(head.Revision, out var samePosition) || samePosition == head,
                "Referenced publications fork at the same revision.");
            revisions[head.Revision] = head;
            if (published.Contains(head)) continue;
            var state = _states[head.State];
            Require(state.Operation is null && head.Revision < anchor.Revision && legacyVersions.Contains(state.Version)
                && anchor.Revision - head.Revision == depths[_states[anchor.State].Version] - depths[state.Version],
                "Referenced publication is not on the captured history chain.");
        }
        Require(_decisions.Keys.All(decisions.Contains), "Referenced decision was not published on the captured history chain.");
    }

    private void ValidateDecisionPublication(HistoryHead head, HistoryBlobReference id)
    {
        var s = _states[head.State]; var d = _decisions[id]; var input = _inputs[d.Input];
        var before = _states[d.Before.State];
        Require(d.Revision == head.Revision && s.ParentPublication == d.Before && s.Snapshot == d.AfterSnapshot
            && s.Version == d.Version && s.Requests?.Current == new HistoryRequestIdentity(input.Request.RequestId, d.Input.Digest),
            "Decision disagrees with its publication or receipt.");
        if (d.ContentCommit is null) Require(s.Version == before.Version && s.Snapshot == before.Snapshot,
            "Decision without effects changed the document.");
        else
        {
            var c = _commits[d.ContentCommit];
            Require(d.Status == "accepted" && s.Commit == d.ContentCommit && s.Sequence == before.Sequence + 1
                && s.Epoch == before.Epoch && c.Kind == "import" && c.Parent == before.Commit && c.Before == before.Snapshot
                && c.After == s.Snapshot && c.Version == s.Version, "Decision content effects disagree.");
        }
    }

    private async ValueTask ValidateEffectsAsync(CancellationToken ct)
    {
        // Derive the actual entry delta from the exact endpoints and compare the complete
        // recorded footprint. Payload bytes were independently hash-verified during traversal.
        // This proves forward/inverse replay without allocating a reconstructed ZIP from
        // untrusted effects. Only one pair of snapshots is held at a time.
        foreach (var c in _commits.Values.Where(c => c.Contribution is not null))
        {
            ct.ThrowIfCancellationRequested(); Spend();
            var inventory = _contributions[c.Contribution!];
            var before = await ReadInspectedSnapshotAsync(c.Before, ct).ConfigureAwait(false);
            var after = await ReadInspectedSnapshotAsync(c.After, ct).ConfigureAwait(false);
            Require(EntryDelta(before, after).SequenceEqual(inventory.Changes.OrderBy(c => c.Uri, StringComparer.Ordinal)),
                "Recorded contribution is not the exact endpoint entry delta.");
        }
        foreach (var d in _decisions.Values.Where(d => d.AppliedText is not null))
        {
            ct.ThrowIfCancellationRequested(); Spend();
            var beforeRef = _states[d.Before.State].Snapshot;
            var before = await ReadSnapshotAsync(beforeRef, ct).ConfigureAwait(false);
            ReserveTextWork(beforeRef, d.AppliedText!);
            var after = DocxOperationPackage.ApplyText(before, d.AppliedText!, _packageOptions);
            Require(InspectSnapshot(after).OrderedOpcContentDigest == d.AfterSnapshot.ContentDigest,
                "Recorded text map does not describe the accepted effects.");
        }
        foreach (var d in _decisions.Values)
        {
            ct.ThrowIfCancellationRequested(); Spend();
            var request = _inputs[d.Input].Request; var baseline = _states[request.Base.State].Snapshot;
            if (request.Kind == "discard") Require(d.ProposedSnapshot == baseline, "Discard proposal is not its base snapshot.");
            if (request.Kind != "text") continue;
            var before = await ReadSnapshotAsync(baseline, ct).ConfigureAwait(false);
            ReserveTextWork(baseline, request.Text!);
            var proposed = DocxOperationPackage.ApplyText(before, request.Text!, _packageOptions);
            Require(InspectSnapshot(proposed).OrderedOpcContentDigest == d.ProposedSnapshot.ContentDigest,
                "Preserved proposal does not represent the original text intent.");
        }
    }

    private async ValueTask<byte[]> ReadSnapshotAsync(DocxSnapshotReference snapshot, CancellationToken ct)
        => (await ReadInspectedSnapshotAsync(snapshot, ct).ConfigureAwait(false)).Bytes;

    private async ValueTask<(byte[] Bytes, PackageManifest Manifest)> ReadInspectedSnapshotAsync(
        DocxSnapshotReference snapshot, CancellationToken ct)
    {
        ChargeRead(snapshot.Blob.Length);
        var bytes = await HistoryBlobIO.ReadBytesAsync(_blobs, snapshot.Blob,
            Math.Min(_limits.MaxBlobBytes, _limits.MaxSnapshotBytes), ct).ConfigureAwait(false);
        var manifest = InspectSnapshot(bytes);
        Require(manifest.OrderedOpcContentDigest == snapshot.ContentDigest, "Snapshot OPC identity changed.");
        return (bytes, manifest);
    }

    private static IEnumerable<PackageEntryChange> EntryDelta((byte[] Bytes, PackageManifest Manifest) before,
        (byte[] Bytes, PackageManifest Manifest) after)
    {
        // Entry names are ZIP metadata. Do not use PackageChangeSet.Create: it would retain
        // all changed uncompressed payloads before checking the hostile contribution inventory.
        var leftNames = EntryNames(before.Bytes); var rightNames = EntryNames(after.Bytes);
        var left = before.Manifest.Entries.ToDictionary(e => e.Uri, StringComparer.Ordinal);
        var right = after.Manifest.Entries.ToDictionary(e => e.Uri, StringComparer.Ordinal);
        foreach (var uri in left.Keys.Union(right.Keys, StringComparer.Ordinal).OrderBy(uri => uri, StringComparer.Ordinal))
        {
            left.TryGetValue(uri, out var a); right.TryGetValue(uri, out var b);
            if (a?.RawBytesDigest == b?.RawBytesDigest) continue;
            yield return new PackageEntryChange(uri, a is null ? null : leftNames[uri], a?.RawBytesDigest,
                b is null ? null : rightNames[uri], b?.RawBytesDigest);
        }
    }

    private static Dictionary<string, string> EntryNames(byte[] bytes)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        return zip.Entries.ToDictionary(e =>
        {
            Require(PackageManifestGenerator.TryCanonicalizeEntryName(e.FullName, out var uri), "Invalid OPC entry name.");
            return uri;
        }, e => e.FullName, StringComparer.Ordinal);
    }

    private PackageManifest InspectSnapshot(byte[] bytes)
    {
        // The inspector enforces this remaining allowance BEFORE inflating package entries.
        // Eight passes conservatively covers inspection's XML/facts reads and hashing.
        var remaining = (_limits.MaxExpandedBytes - _expandedBytes) / 8;
        Budget(remaining > 0, "Archive package-expansion budget reached.");
        var options = _packageOptions ?? new PackageManifestOptions();
        var manifest = new DocxSnapshotStore(_blobs, Math.Min(_limits.MaxBlobBytes, _limits.MaxSnapshotBytes), options with
        {
            MaxTotalUncompressedBytes = Math.Min(options.MaxTotalUncompressedBytes, remaining),
            MaxEntryUncompressedBytes = Math.Min(options.MaxEntryUncompressedBytes, remaining),
        }).Inspect(bytes);
        ChargeExpanded(8 * manifest.Entries.Sum(e => e.Size));
        return manifest;
    }

    private void ReserveTextWork(DocxSnapshotReference before, DocxTextSplice text) =>
        // Includes input inspection, entry copies, XML escaping/serialization and output
        // construction. Output inspection separately reserves its own remaining allowance.
        ChargeExpanded(8 * (6 * _snapshotSizes[before.Blob] + 6L * text.Insert.Length + 1024));

    private void ChargeExpanded(long bytes)
    {
        Budget(bytes <= _limits.MaxExpandedBytes - _expandedBytes, "Archive package-expansion budget reached.");
        _expandedBytes += bytes;
    }

    private PackageChangeLimits ChangeLimits() => new()
    {
        MaxManifestBytes = _limits.MaxManifestBytes, MaxPayloadBytes = _limits.MaxBlobBytes,
        MaxTotalPayloadBytes = Math.Min(_limits.MaxTotalBlobBytes, 512L * 1024 * 1024),
    };

    private void ChargeRead(long bytes)
    {
        Budget(bytes <= _limits.MaxValidationBytes - _validationBytes, "Archive graph-blob validation exceeds its byte budget.");
        _validationBytes += bytes;
    }

    private DocxHistoryView ViewAt(HistoryHead head)
    {
        var state = _states[head.State];
        return new(head, state, new(state.Version, _versions[state.Version]));
    }
    private static void Require(bool condition, string message)
    { if (!condition) throw new DocxHistoryException(DocxHistoryError.InvalidHistory, message); }
    private static void Budget(bool condition, string message)
    { if (!condition) throw new PackageChangeException(PackageChangeError.ResourceLimit, message); }
}
