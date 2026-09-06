// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.History;

/// <summary>
/// Local-filesystem heads with cross-instance/process exclusive file sharing and atomic replacement.
/// Requires a protected host-owned directory and a filesystem supporting these local locking/rename
/// semantics; not a distributed/network-filesystem coordinator. Locks release on process exit.
/// Persistent .lock files must never be deleted while the store is in use. Power-loss durability
/// depends on host filesystem directory guarantees, as with FileHistoryBlobStore.
/// </summary>
public sealed class FileHistoryHeadStore : IHistoryHeadStore
{
    private readonly string _directory;

    public FileHistoryHeadStore(string directory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(directory);
        _directory = Path.GetFullPath(directory);
    }

    public async ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken = default)
    {
        var path = HeadPath(HistoryHeadCodec.Key(documentId));
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete,
                4096, FileOptions.Asynchronous);
            if (file.Length > HistoryHeadCodec.MaxBytes) throw new InvalidDataException("History head exceeds the byte limit.");
            var bytes = new byte[(int)file.Length];
            await file.ReadExactlyAsync(bytes, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return HistoryHeadCodec.Decode(bytes);
        }
        catch (FileNotFoundException) { return null; }
        catch (DirectoryNotFoundException) { return null; }
    }

    public async ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected,
        HistoryBlobReference state, CancellationToken cancellationToken = default)
    {
        var key = HistoryHeadCodec.Key(documentId);
        var next = HistoryHeadCodec.Next(expected, state);
        cancellationToken.ThrowIfCancellationRequested();
        Directory.CreateDirectory(_directory);
        using var gate = await AcquireAsync(Path.Combine(_directory, key + ".lock"), cancellationToken).ConfigureAwait(false);
        if (await ReadAsync(documentId, cancellationToken).ConfigureAwait(false) != expected) return null;
        var temporary = Path.Combine(_directory, $".{Guid.NewGuid():N}.tmp");
        var ownsTemporary = false;
        try
        {
            await using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None,
                4096, FileOptions.Asynchronous))
            {
                ownsTemporary = true;
                await file.WriteAsync(HistoryHeadCodec.Encode(next), cancellationToken).ConfigureAwait(false);
                await file.FlushAsync(cancellationToken).ConfigureAwait(false);
                file.Flush(flushToDisk: true);
            }
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(temporary, HeadPath(key), overwrite: true);
            return next;
        }
        finally
        {
            try { if (ownsTemporary) File.Delete(temporary); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private string HeadPath(string key) => Path.Combine(_directory, key + ".head");

    private static async ValueTask<FileStream> AcquireAsync(string path, CancellationToken cancellationToken)
    {
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var gate = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
                try
                {
                    // Some runtimes/configurations/filesystems silently disable advisory locks.
                    // Fail closed rather than publishing a head with no actual writer exclusion.
                    try
                    {
                        using var probe = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                        throw new PlatformNotSupportedException("The filesystem/runtime does not enforce exclusive file sharing.");
                    }
                    catch (IOException error) when (IsLockContention(error)) { }
                    return gate;
                }
                catch { gate.Dispose(); throw; }
            }
            // Only sharing violations are retryable; permissions and other I/O failures propagate.
            catch (IOException error) when (IsLockContention(error))
            {
                await Task.Delay(10, cancellationToken).ConfigureAwait(false);
            }
        }
    }

    private static bool IsLockContention(IOException error) =>
        OperatingSystem.IsWindows() ? (error.HResult & 0xffff) == 32
        : OperatingSystem.IsLinux() || OperatingSystem.IsAndroid() ? error.HResult == 11
        : OperatingSystem.IsMacOS() || OperatingSystem.IsFreeBSD() ? error.HResult == 35
        : false;
}
