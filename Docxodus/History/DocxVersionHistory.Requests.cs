// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text.Json;
using Docxodus.Verification;

namespace Docxodus.History;

public sealed partial class DocxVersionHistory
{
    private async ValueTask<DocxHistoryView?> FindRequestAsync(string documentId, DocxHistoryView? current,
        HistoryRequestIdentity? request, CancellationToken cancellationToken)
    {
        if (request is null) return null;
        var original = await _requests.FindAsync(documentId, current?.Head, current?.State.Requests,
            request, cancellationToken).ConfigureAwait(false);
        if (original is null) return null;
        var view = original == current!.Head ? current
            : await ReadHeadViewAsync(documentId, original, cancellationToken).ConfigureAwait(false);
        Consistent(original.Revision <= current.Head.Revision && view.State.Requests?.Current == request,
            "Request receipt does not match its original publication.");
        return view;
    }

    // Explicit, versioned, length-delimited canonical input. Metadata has already been copied and
    // sorted before any await. Exact raw bytes distinguish repacks; timestamps retain their offset
    // and precision. Nonces/generated results are intentionally NOT part of caller input identity.
    private static HistoryRequestIdentity VersionRequest(string operation, string documentId, HistoryHead? expected,
        DocxVersionMetadata metadata, string requestId, byte[]? docx, HistoryBlobReference? target)
    {
        HistoryRequestJournalStore.ValidateId(requestId);
        if (expected is not null)
        {
            if (expected.Revision <= 0) throw new ArgumentOutOfRangeException(nameof(expected));
            HistoryBlobIO.Validate(expected.State, int.MaxValue);
        }
        if (target is not null) HistoryBlobIO.Validate(target, int.MaxValue);
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject();
            writer.WriteString("schema", "docxodus.version-request/v1");
            writer.WriteString("operation", operation);
            writer.WriteString("documentId", documentId);
            writer.WritePropertyName("expectedHead");
            if (expected is null) writer.WriteNullValue();
            else
            {
                writer.WriteStartObject(); writer.WriteNumber("revision", expected.Revision);
                Reference("state", expected.State); writer.WriteEndObject();
            }
            Reference("target", target);
            if (docx is null) writer.WriteNull("snapshot");
            else Reference("snapshot", new HistoryBlobReference(new VerificationDigest
            { Algorithm = "SHA-256", Value = Convert.ToHexString(SHA256.HashData(docx)).ToLowerInvariant() }, docx.Length));
            writer.WriteString("author", metadata.Author);
            writer.WriteString("createdAt", metadata.CreatedAt.ToString("O", System.Globalization.CultureInfo.InvariantCulture));
            writer.WriteString("label", metadata.Label); writer.WriteString("message", metadata.Message);
            writer.WriteStartObject("applicationMetadata");
            foreach (var pair in metadata.ApplicationMetadata) writer.WriteString(pair.Key, pair.Value);
            writer.WriteEndObject(); writer.WriteEndObject();

            void Reference(string name, HistoryBlobReference? reference)
            {
                writer.WritePropertyName(name);
                if (reference is null) { writer.WriteNullValue(); return; }
                writer.WriteStartObject(); writer.WriteString("sha256", reference.Digest.Value);
                writer.WriteNumber("length", reference.Length); writer.WriteEndObject();
            }
        }
        return new HistoryRequestIdentity(requestId, new VerificationDigest
        { Algorithm = "SHA-256", Value = Convert.ToHexString(SHA256.HashData(buffer.ToArray())).ToLowerInvariant() });
    }
}
