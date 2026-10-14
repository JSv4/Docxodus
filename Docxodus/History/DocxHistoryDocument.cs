// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Docxodus.History;

/// <summary>
/// Document-scoped reads and explicit checkpoint/restore publication. The host owns the stores,
/// stable document ID and request inputs. Creating this facade does not create a version or editor.
/// </summary>
public sealed class DocxHistoryDocument : DocxHistoryReader
{
    internal DocxHistoryDocument(DocxVersionHistory history, string documentId) : base(history, documentId) { }

    /// <summary>Persist requestId and the exact inputs before calling; retry them unchanged after lost acknowledgement.</summary>
    public ValueTask<DocxHistoryView> CreateVersionAsync(string requestId, HistoryHead? expected, byte[] docxBytes,
        DocxVersionMetadata metadata, CancellationToken cancellationToken = default) =>
        Core.CreateVersionAsync(DocumentId, requestId, expected, docxBytes, metadata, cancellationToken);

    /// <summary>Publish a new version; never rewinds/deletes history or changes a live editor.</summary>
    public ValueTask<DocxHistoryView> RestoreVersionAsync(string requestId, HistoryHead expected,
        HistoryBlobReference versionId, DocxVersionMetadata metadata, CancellationToken cancellationToken = default) =>
        Core.RestoreVersionAsync(DocumentId, requestId, expected, versionId, metadata, cancellationToken);
}
