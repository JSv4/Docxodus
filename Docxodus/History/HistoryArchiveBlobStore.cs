// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Compression;

namespace Docxodus.History;

/// <summary>Serializes ZIP entry reads; each returned stream owns the gate until disposed.</summary>
internal sealed class HistoryArchiveBlobStore(ZipArchive zip, Stream input, bool leaveOpen,
    IReadOnlyDictionary<string, ZipArchiveEntry> entries, int maxBlobBytes) : IHistoryBlobStore, IDisposable
{
    private readonly SemaphoreSlim _gate = new(1);
    private bool _disposed;

    public ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken = default) =>
        throw new NotSupportedException("History archives are read-only.");

    public async ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default)
    {
        HistoryBlobIO.Validate(reference, maxBlobBytes);
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (!entries.TryGetValue(HistoryArchiveManifest.BlobName(reference), out var entry))
            { _gate.Release(); return null; }
            if (entry.Length != reference.Length)
                throw new PackageChangeException(PackageChangeError.PayloadMismatch, "Archive blob length disagrees with its reference.");
            return new EntryLease(entry.Open(), _gate);
        }
        catch { _gate.Release(); throw; }
    }

    public void Dispose()
    {
        if (_disposed) return;
        if (!_gate.Wait(0)) throw new InvalidOperationException("Await active archive reads before disposing the archive.");
        try
        {
            if (_disposed) return;
            _disposed = true;
            try { zip.Dispose(); }
            finally { if (!leaveOpen) input.Dispose(); }
        }
        finally { _gate.Release(); }
    }

    private sealed class EntryLease(Stream inner, SemaphoreSlim gate) : Stream
    {
        private int _closed;
        public override bool CanRead => inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => inner.Length;
        public override long Position { get => inner.Position; set => throw new NotSupportedException(); }
        public override void Flush() => throw new NotSupportedException();
        public override int Read(byte[] buffer, int offset, int count) => inner.Read(buffer, offset, count);
        public override int Read(Span<byte> buffer) => inner.Read(buffer);
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default) =>
            inner.ReadAsync(buffer, cancellationToken);
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            inner.ReadAsync(buffer, offset, count, cancellationToken);
        protected override void Dispose(bool disposing)
        {
            if (disposing && Interlocked.Exchange(ref _closed, 1) == 0)
                try { inner.Dispose(); } finally { gate.Release(); }
            base.Dispose(disposing);
        }
        public override async ValueTask DisposeAsync()
        {
            if (Interlocked.Exchange(ref _closed, 1) != 0) return;
            try { await inner.DisposeAsync().ConfigureAwait(false); }
            finally { gate.Release(); }
            GC.SuppressFinalize(this);
        }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
