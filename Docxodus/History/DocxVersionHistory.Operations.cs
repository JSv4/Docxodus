#nullable enable

namespace Docxodus.History;

public sealed partial class DocxVersionHistory
{
    /// <summary>
    /// Reconcile a durable backend intent against accepted history and atomically publish effects
    /// OR a recoverable conflict. Concurrent CAS losers re-evaluate the ORIGINAL input. No GUI,
    /// transport, outbox, subscription, or optimistic live session is installed. Initialize with
    /// CreateVersionAsync first. Each bounded ancestry/decision scan is limited independently.
    /// </summary>
    public async ValueTask<DocxOperationResult> SubmitOperationAsync(string documentId, DocxOperationRequest request,
        byte[]? candidateDocx = null, int maxAttempts = 16, int maxEntriesToScan = 10_000,
        CancellationToken cancellationToken = default)
    {
        if (maxAttempts is < 1 or > 1024) throw new ArgumentOutOfRangeException(nameof(maxAttempts));
        if (maxEntriesToScan <= 0) throw new ArgumentOutOfRangeException(nameof(maxEntriesToScan));
        cancellationToken.ThrowIfCancellationRequested();
        if (candidateDocx?.Length > _maxSnapshotBytes)
            throw new PackageChangeException(PackageChangeError.ResourceLimit, "Candidate exceeds the snapshot byte limit.");
        var captured = candidateDocx?.ToArray();
        var input = DocxOperationStore.Capture(documentId, request, captured);
        var inputBytes = DocxOperationStore.EncodeInput(input);
        var inputId = DocxOperationStore.Reference(inputBytes);
        var identity = new HistoryRequestIdentity(input.Request.RequestId, inputId.Digest);
        var store = new DocxOperationStore(_blobs);
        var current = await ReadAsync(documentId, cancellationToken).ConfigureAwait(false)
            ?? throw new DocxHistoryException(DocxHistoryError.HistoryUnavailable, "Initialize document history before submitting operations.");
        var duplicate = await FindRequestAsync(documentId, current, identity, cancellationToken).ConfigureAwait(false);
        if (duplicate is not null) return await OperationResultAsync(documentId, duplicate, inputId, cancellationToken).ConfigureAwait(false);

        var baseline = await ReadHeadViewAsync(documentId, input.Request.Base, cancellationToken).ConfigureAwait(false);
        // Prove the submitted base belongs to this publication chain before preserving new input.
        _ = await ReadChangesBetweenAsync(documentId, current, baseline.Head, maxEntriesToScan, cancellationToken).ConfigureAwait(false);
        var baseBytes = await _snapshots.ExportAsync(baseline.State.Snapshot, cancellationToken).ConfigureAwait(false);
        var proposedBytes = input.Request.Kind switch
        {
            "text" => DocxOperationPackage.ApplyText(baseBytes, input.Request.Text!, _packageOptions),
            "package" => captured!,
            _ => baseBytes,
        };
        var proposed = await _snapshots.CaptureAsync(proposedBytes, cancellationToken).ConfigureAwait(false);
        Consistent(input.Candidate is null || input.Candidate == proposed.Blob, "Captured candidate disagrees with its input fingerprint.");
        _ = await HistoryBlobIO.PutBytesAsync(_blobs, inputBytes, cancellationToken).ConfigureAwait(false);
        var contribution = input.Request.Kind == "package" ? PackageChangeSet.Create(baseBytes, proposedBytes, _packageOptions) : null;

        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (attempt > 0)
            {
                current = await ReadAsync(documentId, cancellationToken).ConfigureAwait(false)
                    ?? throw new DocxHistoryException(DocxHistoryError.InvalidHistory, "Published document history disappeared.");
                duplicate = await FindRequestAsync(documentId, current, identity, cancellationToken).ConfigureAwait(false);
                if (duplicate is not null) return await OperationResultAsync(documentId, duplicate, inputId, cancellationToken).ConfigureAwait(false);
            }
            var updates = await ReadChangesBetweenAsync(documentId, current, baseline.Head, maxEntriesToScan, cancellationToken).ConfigureAwait(false);
            var decisions = await ReadOperationChainAsync(documentId, current,
                input.Request.Resolves is null ? baseline.State.Operation : null, maxEntriesToScan, cancellationToken).ConfigureAwait(false);
            string? conflict = null;
            if (input.Request.Resolves is { } resolves)
            {
                var target = decisions.SingleOrDefault(d => d.Id == resolves);
                Consistent(target is not null && target.Record.Status == "conflict" && target.Record.Revision <= baseline.Head.Revision,
                    "Resolution must name an earlier conflict observed at the submitted base.");
                if (decisions.Any(d => d.Record.Status == "accepted" && d.Input.Request.Resolves == resolves)) conflict = "AlreadyResolved";
            }
            if (current.State.Epoch != baseline.State.Epoch) conflict ??= "EpochChanged";
            var currentBytes = await _snapshots.ExportAsync(current.State.Snapshot, cancellationToken).ConfigureAwait(false);
            conflict ??= DocxOperationPackage.CheckReads(baseBytes, currentBytes, input.Request.ReadParts, _packageOptions);
            byte[]? accepted = null;
            var mapped = input.Request.Text;
            if (conflict is null && mapped is not null)
            {
                var byCommit = decisions.Where(d => d.Record.ContentCommit is not null)
                    .ToDictionary(d => d.Record.ContentCommit!);
                foreach (var entry in updates.Entries)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Consistent(entry.Commit.Contribution is not null, "A same-epoch text tail cannot contain a reset.");
                    var manifest = await HistoryBlobIO.ReadBytesAsync(_blobs, entry.Commit.Contribution!,
                        new PackageChangeLimits().MaxManifestBytes, cancellationToken).ConfigureAwait(false);
                    var changes = await PackageChangeSetCodec.LoadAsync(manifest, _blobs, cancellationToken: cancellationToken).ConfigureAwait(false);
                    Consistent(changes.BeforeDigest == entry.Commit.Before.ContentDigest && changes.AfterDigest == entry.Commit.After.ContentDigest,
                        "Contribution endpoints disagree with their accepted commit.");
                    var beforeBytes = await _snapshots.ExportAsync(entry.Commit.Before, cancellationToken).ConfigureAwait(false);
                    _ = changes.Apply(beforeBytes, _packageOptions); // Verify the footprint, not just a claimed list of changed URIs.
                    if (changes.Changes.Any(c => DocxOperationPackage.IsTopology(c.Uri) || input.Request.ReadParts.Contains(c.Uri)))
                    { conflict = "ReadDependencyChanged"; break; }
                    if (!changes.Changes.Any(c => c.Uri == mapped.PartUri)) continue;
                    if (!byCommit.TryGetValue(entry.Id, out var decision) || decision.Record.AppliedText is null)
                    { conflict = "UnknownTextChange"; break; }
                    VerifyRecordedText(decision, entry.Commit, beforeBytes);
                    mapped = DocxOperationPackage.Map(mapped, decision.Record.AppliedText);
                    if (mapped is null) { conflict = "OverlappingText"; break; }
                }
                if (conflict is null) accepted = DocxOperationPackage.ApplyText(currentBytes, mapped!, _packageOptions);
            }
            else if (conflict is null && contribution is not null)
            {
                (accepted, conflict) = DocxOperationPackage.Merge(currentBytes, contribution, _packageOptions);
            }

            var state = current.State;
            var version = current.Version;
            var changed = conflict is null && accepted is not null
                && PackageChangeSet.Create(currentBytes, accepted, _packageOptions).Changes.Count != 0;
            if (changed)
                (state, version) = await PrepareCandidateAsync(documentId, current, accepted!, input.Request.Metadata, cancellationToken).ConfigureAwait(false);
            var record = new DocxOperationRecord
            {
                DocumentId = documentId, Input = inputId, Parent = current.State.Operation,
                Revision = checked(current.Head.Revision + 1), Before = current.Head,
                ProposedSnapshot = proposed, AfterSnapshot = state.Snapshot, Version = version.Id,
                ContentCommit = changed ? state.Commit : null, Status = conflict is null ? "accepted" : "conflict",
                Conflict = conflict, AppliedText = changed && input.Request.Kind == "text" ? mapped : null,
            };
            var decisionId = await store.SaveDecisionAsync(record, cancellationToken).ConfigureAwait(false);
            try
            {
                var published = await PublishAsync(documentId, current.Head, state with { Operation = decisionId }, version,
                    cancellationToken, identity, current.State.Requests).ConfigureAwait(false);
                // No reads/cancellation after OUR successful commit. Another identical winner's
                // original result is separately verified instead of returning this abandoned draft.
                if (published.State.Operation == decisionId)
                    return new DocxOperationResult(published, new DocxStoredOperation(decisionId, record, input));
                return await OperationResultAsync(documentId, published, inputId, cancellationToken).ConfigureAwait(false);
            }
            catch (DocxHistoryException error) when (error.Code == DocxHistoryError.StaleHead) { }
        }
        throw new DocxHistoryException(DocxHistoryError.Contention, "Publication contention limit reached; retry the same captured request.");
    }

    /// <summary>Ascending durable outcomes since an accepted head; null scans all retained decisions.</summary>
    public async ValueTask<DocxOperationUpdate> ReadOperationsSinceAsync(string documentId, HistoryHead? after,
        int maxEntriesToScan = 10_000, CancellationToken cancellationToken = default)
    {
        if (maxEntriesToScan <= 0) throw new ArgumentOutOfRangeException(nameof(maxEntriesToScan));
        var current = await ReadAsync(documentId, cancellationToken).ConfigureAwait(false)
            ?? throw new DocxHistoryException(DocxHistoryError.HistoryUnavailable, "Document history is unavailable.");
        _ = await ReadChangesBetweenAsync(documentId, current, after, maxEntriesToScan, cancellationToken).ConfigureAwait(false);
        var stop = after is null ? null : (await ReadHeadViewAsync(documentId, after, cancellationToken).ConfigureAwait(false)).State.Operation;
        var entries = await ReadOperationChainAsync(documentId, current, stop, maxEntriesToScan, cancellationToken).ConfigureAwait(false);
        entries.Reverse();
        return new DocxOperationUpdate(after, current, entries.AsReadOnly());
    }

    /// <summary>Read a retained immutable outcome. Reference access is host-authorized, not publication proof.</summary>
    public async ValueTask<DocxStoredOperation> GetOperationAsync(string documentId, HistoryBlobReference operationId,
        CancellationToken cancellationToken = default)
    {
        HistoryHeadCodec.Key(documentId);
        var store = new DocxOperationStore(_blobs);
        var record = await store.LoadDecisionAsync(operationId, cancellationToken).ConfigureAwait(false);
        SameDocument(documentId, record.DocumentId);
        var input = await store.LoadInputAsync(record.Input, cancellationToken).ConfigureAwait(false);
        SameDocument(documentId, input.DocumentId);
        Consistent(input.Request.Base.Revision <= record.Before.Revision, "Operation base is newer than its before publication.");
        Consistent(DocxOperationStore.Reference(DocxOperationStore.EncodeInput(input)) == record.Input,
            "Operation input is not its canonical fingerprint.");
        Consistent(input.Candidate is null || input.Candidate == record.ProposedSnapshot.Blob,
            "Preserved package proposal does not match original input.");
        if (record.AppliedText is { } applied)
            Consistent(input.Request.Kind == "text" && input.Request.Text is { } original
                && applied == original with { Offset = applied.Offset }, "Mapped text changed the original intent.");
        return new DocxStoredOperation(operationId, record, input);
    }

    /// <summary>Exact preserved proposal, including work that conflicted; does not change current history.</summary>
    public async ValueTask<byte[]> ExportOperationProposalAsync(string documentId, HistoryBlobReference operationId,
        CancellationToken cancellationToken = default) => await _snapshots.ExportAsync(
            (await GetOperationAsync(documentId, operationId, cancellationToken).ConfigureAwait(false)).Record.ProposedSnapshot,
            cancellationToken).ConfigureAwait(false);

    private async ValueTask<DocxOperationResult> OperationResultAsync(string documentId, DocxHistoryView publication,
        HistoryBlobReference inputId, CancellationToken cancellationToken)
    {
        Consistent(publication.State.Operation is not null, "Request receipt does not identify an operation publication.");
        var operation = await GetOperationAsync(documentId, publication.State.Operation!, cancellationToken).ConfigureAwait(false);
        Consistent(operation.Record.Input == inputId, "Request receipt identifies a different operation input.");
        await ValidateOperationPublicationAsync(documentId, publication, operation, cancellationToken).ConfigureAwait(false);
        return new DocxOperationResult(publication, operation);
    }

    private async ValueTask<List<DocxStoredOperation>> ReadOperationChainAsync(string documentId, DocxHistoryView current,
        HistoryBlobReference? stop, int maximum, CancellationToken cancellationToken)
    {
        var entries = new List<DocxStoredOperation>();
        var publication = current;
        var scanned = 0;
        // Correlate decisions with exact publication ancestry in one bounded walk. Restarting
        // a radix receipt lookup per decision would multiply this budget by up to 257 reads.
        while (publication.State.Operation != stop)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Consistent(publication.State.Operation is not null && publication.State.ParentPublication is not null,
                "Accepted operation log is not an ancestor of the new head.");
            if (scanned++ >= maximum) throw new DocxHistoryException(DocxHistoryError.TraversalLimit, "Operation publication scan budget reached.");
            var parent = await ReadHeadViewAsync(documentId, publication.State.ParentPublication!, cancellationToken).ConfigureAwait(false);
            ValidatePublicationEdge(publication, parent);
            if (publication.State.Operation != parent.State.Operation)
            {
                var operation = await GetOperationAsync(documentId, publication.State.Operation!, cancellationToken).ConfigureAwait(false);
                await ValidateOperationPublicationAsync(documentId, publication, operation, cancellationToken, parent).ConfigureAwait(false);
                entries.Add(operation);
            }
            publication = parent;
        }
        return entries;
    }

    private async ValueTask ValidateOperationPublicationAsync(string documentId, DocxHistoryView publication,
        DocxStoredOperation operation, CancellationToken cancellationToken, DocxHistoryView? before = null)
    {
        var record = operation.Record;
        Consistent(publication.Head.Revision == record.Revision && publication.State.ParentPublication == record.Before
            && publication.State.Operation == operation.Id && publication.State.Snapshot == record.AfterSnapshot
            && publication.State.Version == record.Version, "Decision disagrees with its original publication.");
        Consistent(publication.State.Requests?.Current == new HistoryRequestIdentity(operation.Input.Request.RequestId, record.Input.Digest),
            "Decision is not bound to its original inline request receipt.");
        before ??= await ReadHeadViewAsync(documentId, record.Before, cancellationToken).ConfigureAwait(false);
        ValidatePublicationEdge(publication, before);
        Consistent(record.Parent == before.State.Operation, "Decision does not extend the preceding operation tip.");
        if (record.ContentCommit is null)
            Consistent(publication.State.Version == before.State.Version && publication.State.Snapshot == before.State.Snapshot,
                "Decision without effects changed the document.");
        else
        {
            Consistent(record.Status == "accepted" && publication.State.Commit == record.ContentCommit
                && publication.State.Sequence == before.State.Sequence + 1 && publication.State.Epoch == before.State.Epoch,
                "Accepted effects disagree with the document commit.");
            var commit = await _records.LoadCommitAsync(record.ContentCommit, cancellationToken).ConfigureAwait(false);
            Consistent(commit.Kind == "import" && commit.Parent == before.State.Commit && commit.Before == before.State.Snapshot
                && commit.After == publication.State.Snapshot && commit.Version == publication.Version.Id,
                "Accepted content commit does not extend its before publication.");
        }
    }

    private void VerifyRecordedText(DocxStoredOperation operation, PackageHistoryCommitRecord commit, byte[] before)
    {
        var applied = DocxOperationPackage.ApplyText(before, operation.Record.AppliedText!, _packageOptions);
        Consistent(PackageChangeSet.Create(before, applied, _packageOptions).AfterDigest == commit.After.ContentDigest,
            "Recorded text map does not describe the accepted package effects.");
    }
}
