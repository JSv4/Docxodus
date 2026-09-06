// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Docxodus.History;

/// <summary>The imported captured view, not a later reread. AlreadyPresent means exact-head idempotence.</summary>
public sealed record DocxHistoryImportResult(DocxHistoryArchiveInfo Archive, DocxHistoryView View, bool AlreadyPresent);

public sealed partial class DocxVersionHistory
{
    /// <summary>
    /// Validate an owned copy before any destination writes, then copy immutable blobs and initialize
    /// the original head. Different existing heads fail ImportConflict without being overwritten.
    /// </summary>
    public ValueTask<DocxHistoryImportResult> ImportHistoryArchiveAsync(byte[] bytes,
        DocxHistoryArchiveLimits? limits = null, CancellationToken cancellationToken = default)
    {
        var initializer = Initializer();
        return ImportAsync(DocxHistoryArchive.OpenAsync(bytes, ImportLimits(limits), _packageOptions, cancellationToken),
            initializer, cancellationToken);
    }

    /// <summary>
    /// Input must remain unchanged, readable and seekable until completion. Leaves it open by default.
    /// Failed copies may leave harmless immutable blobs but no new head. Retry the same archive after
    /// uncertain storage failures; an already-advanced destination produces ImportConflict.
    /// </summary>
    public ValueTask<DocxHistoryImportResult> ImportHistoryArchiveAsync(Stream input, bool leaveOpen = true,
        DocxHistoryArchiveLimits? limits = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        IHistoryHeadInitializer initializer; DocxHistoryArchiveLimits bounded;
        // leaveOpen: false hands this call the stream, exactly as the archive reader's own failure
        // path does. Rejecting the store or the limits happens before the reader takes over, so
        // dispose here too rather than leaking a stream the caller has already given up.
        try { initializer = Initializer(); bounded = ImportLimits(limits); }
        catch { if (!leaveOpen) input.Dispose(); throw; }
        return ImportAsync(DocxHistoryArchive.OpenAsync(input, leaveOpen, bounded, _packageOptions, cancellationToken),
            initializer, cancellationToken);
    }

    private IHistoryHeadInitializer Initializer() => _heads as IHistoryHeadInitializer
        ?? throw new DocxHistoryException(DocxHistoryError.InitializationUnsupported,
            "Writable archive import requires absent-only exact-head initialization.");

    private DocxHistoryArchiveLimits ImportLimits(DocxHistoryArchiveLimits? limits)
    {
        limits ??= new(); limits.Validate();
        return limits with
        {
            MaxSnapshotBytes = Math.Min(limits.MaxSnapshotBytes, _maxSnapshotBytes),
            MaxRecordBytes = Math.Min(limits.MaxRecordBytes, _maxRecordBytes),
        };
    }

    private async ValueTask<DocxHistoryImportResult> ImportAsync(ValueTask<DocxHistoryArchive> opening,
        IHistoryHeadInitializer initializer, CancellationToken ct)
    {
        DocxHistoryArchiveInfo info; DocxHistoryView view;
        using (var archive = await opening.ConfigureAwait(false))
        {
            info = archive.Info; view = archive.View;
            var existing = await _heads.ReadAsync(info.DocumentId, ct).ConfigureAwait(false);
            if (existing is not null && existing != info.Head) throw ImportConflict();
            // Repeat puts even on exact-head retry: missing bytes can be repaired; corrupt immutable
            // collisions fail under the blob-store contract. Never report success on incomplete data.
            foreach (var reference in archive.Inventory)
            {
                ct.ThrowIfCancellationRequested();
                using var content = await archive.Blobs.OpenReadAsync(reference, ct).ConfigureAwait(false);
                if (content is null) throw new PackageChangeException(PackageChangeError.PayloadMissing, "Import blob is missing.");
                await _blobs.PutAsync(reference, content, ct).ConfigureAwait(false);
            }
        } // Dispose reader/entry leases (and owned input) before publication, never after our commit.
        var result = await initializer.TryInitializeAsync(info.DocumentId, info.Head, ct).ConfigureAwait(false);
        if (result.Head != info.Head) throw ImportConflict();
        // No reads or cancellation after this commit point.
        return new(info, view, !result.Initialized);
    }

    private static DocxHistoryException ImportConflict() => new(DocxHistoryError.ImportConflict,
        "A different history already exists for this document. Import never overwrites or remaps its identity.");
}
