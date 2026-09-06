// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Docxodus.History;

/// <summary>Document-scoped read capabilities. Reading never publishes or changes a live editor.</summary>
public class DocxHistoryReader
{
    internal DocxHistoryReader(DocxVersionHistory history, string documentId)
    { Core = history; DocumentId = documentId; }

    internal DocxVersionHistory Core { get; }
    public string DocumentId { get; }

    public ValueTask<DocxHistoryView?> ReadAsync(CancellationToken cancellationToken = default) =>
        Core.ReadAsync(DocumentId, cancellationToken);

    public ValueTask<DocxVersionPage> ListVersionsAsync(HistoryBlobReference? cursor = null, int limit = 25,
        CancellationToken cancellationToken = default) => Core.ListVersionsAsync(DocumentId, cursor, limit, cancellationToken);

    public ValueTask<DocxStoredVersion> GetVersionAsync(HistoryBlobReference versionId,
        CancellationToken cancellationToken = default) => Core.GetVersionAsync(DocumentId, versionId, cancellationToken);

    /// <summary>Exact selected snapshot, or the latest version captured once. No history is embedded.</summary>
    public async ValueTask<byte[]> ExportDocxAsync(HistoryBlobReference? versionId = null,
        CancellationToken cancellationToken = default)
    {
        versionId ??= (await ReadAsync(cancellationToken).ConfigureAwait(false))?.Version.Id
            ?? throw new DocxHistoryException(DocxHistoryError.HistoryUnavailable, "Document history is unavailable.");
        return await Core.ExportVersionAsync(DocumentId, versionId, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Compare any retained pair using the existing DocxDiff engine and its revision policy.</summary>
    public ValueTask<DocxDiffComparison> CompareVersionsAsync(HistoryBlobReference before, HistoryBlobReference after,
        DocxDiffSettings? settings = null, CancellationToken cancellationToken = default) =>
        Core.CompareVersionsAsync(DocumentId, before, after, settings, cancellationToken);

    public ValueTask<byte[]> MaterializeAsync(long sequence, int maxEntriesToScan = 10_000,
        CancellationToken cancellationToken = default) => Core.MaterializeAsync(DocumentId, sequence, maxEntriesToScan, cancellationToken);

    public ValueTask<byte[]> ReplayAsync(long sequence, int maxEntriesToScan = 10_000,
        CancellationToken cancellationToken = default) => Core.ReplayAsync(DocumentId, sequence, maxEntriesToScan, cancellationToken);

    public ValueTask<long> ResolveSequenceAtTimeAsync(DateTimeOffset cutoff, int maxEntriesToScan = 10_000,
        CancellationToken cancellationToken = default) => Core.ResolveSequenceAtTimeAsync(DocumentId, cutoff, maxEntriesToScan, cancellationToken);

    public ValueTask<DocxHistoryUpdate> ReadChangesSinceAsync(HistoryHead? after, int maxEntriesToScan = 10_000,
        CancellationToken cancellationToken = default) => Core.ReadChangesSinceAsync(DocumentId, after, maxEntriesToScan, cancellationToken);

    public ValueTask<DocxOperationUpdate> ReadOperationsSinceAsync(HistoryHead? after, int maxEntriesToScan = 10_000,
        CancellationToken cancellationToken = default) => Core.ReadOperationsSinceAsync(DocumentId, after, maxEntriesToScan, cancellationToken);

    public ValueTask<DocxStoredOperation> GetOperationAsync(HistoryBlobReference operationId,
        CancellationToken cancellationToken = default) => Core.GetOperationAsync(DocumentId, operationId, cancellationToken);

    public ValueTask<byte[]> ExportOperationProposalAsync(HistoryBlobReference operationId,
        CancellationToken cancellationToken = default) => Core.ExportOperationProposalAsync(DocumentId, operationId, cancellationToken);

    public ValueTask<DocxHistoryArchiveInfo> ExportHistoryArchiveAsync(Stream output, DocxHistoryArchiveLimits? limits = null,
        CancellationToken cancellationToken = default) => Core.ExportHistoryArchiveAsync(DocumentId, output, limits, cancellationToken);
}
