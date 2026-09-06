// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.History;

/// <summary>
/// Filesystem reference store with verified, flushed temporary writes and atomic no-overwrite
/// publication. Independent instances may share a root. The host must own and protect that root;
/// this is not a sandbox against external file modification or symlinks. Readers verify content.
/// File durability across power loss depends on the host filesystem's rename/directory semantics.
/// </summary>
public sealed class FileHistoryBlobStore : IHistoryBlobStore
{
    private readonly string _directory;
    private readonly int _maxBlobBytes;

    public FileHistoryBlobStore(string directory, int maxBlobBytes = HistoryBlobIO.DefaultMaxBlobBytes)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(directory);
        _directory = Path.GetFullPath(directory);
        _maxBlobBytes = HistoryBlobIO.ValidateLimit(maxBlobBytes);
    }

    public async ValueTask PutAsync(HistoryBlobReference reference, Stream content,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(content);
        HistoryBlobIO.Validate(reference, _maxBlobBytes);
        cancellationToken.ThrowIfCancellationRequested();
        Directory.CreateDirectory(_directory);
        var destination = BlobPath(reference);
        var temporary = Path.Combine(_directory, $".{Guid.NewGuid():N}.tmp");
        var ownsTemporary = false;
        try
        {
            await using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None,
                64 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan))
            {
                ownsTemporary = true;
                await HistoryBlobIO.CopyVerifiedAsync(reference, content, file, cancellationToken).ConfigureAwait(false);
                await file.FlushAsync(cancellationToken).ConfigureAwait(false);
                file.Flush(flushToDisk: true);
            }
            cancellationToken.ThrowIfCancellationRequested();
            try { File.Move(temporary, destination, overwrite: false); }
            catch (IOException) when (File.Exists(destination))
            {
                // Concurrent/idempotent publication is successful only if the existing bytes
                // are valid. Never repair or overwrite externally corrupted history implicitly.
                using var existing = OpenFile(destination);
                await HistoryBlobIO.CopyVerifiedAsync(reference, existing, Stream.Null, cancellationToken).ConfigureAwait(false);
            }
        }
        finally
        {
            // Only this write's private temporary file is eligible for cleanup. A process crash
            // may leave an unreferenced .tmp file; reclamation is an explicit host responsibility.
            try { if (ownsTemporary) File.Delete(temporary); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference,
        CancellationToken cancellationToken = default)
    {
        HistoryBlobIO.Validate(reference, _maxBlobBytes);
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            var file = OpenFile(BlobPath(reference));
            if (file.Length == reference.Length) return ValueTask.FromResult<Stream?>(file);
            file.Dispose();
            throw new PackageChangeException(PackageChangeError.PayloadMismatch, "Stored blob length does not match its reference.");
        }
        catch (FileNotFoundException) { return ValueTask.FromResult<Stream?>(null); }
        catch (DirectoryNotFoundException) { return ValueTask.FromResult<Stream?>(null); }
    }

    private string BlobPath(HistoryBlobReference reference) => Path.Combine(_directory, reference.Digest.Value + ".blob");

    private static FileStream OpenFile(string path) => new(path, FileMode.Open, FileAccess.Read,
        FileShare.Read, 64 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan);
}
