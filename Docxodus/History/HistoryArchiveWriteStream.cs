// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Docxodus.History;

/// <summary>Forward-only archive output budget. Never owns or closes the destination.</summary>
internal sealed class HistoryArchiveWriteStream(Stream destination, long maximum) : Stream
{
    private long _written;
    private MemoryStream? _directory;
    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => destination.CanWrite;
    public override long Length => _written;
    public override long Position { get => _written; set => throw new NotSupportedException(); }
    public override void Flush() => destination.Flush();
    public override Task FlushAsync(CancellationToken cancellationToken) => destination.FlushAsync(cancellationToken);
    public override void Write(byte[] buffer, int offset, int count) => Write(buffer.AsSpan(offset, count));
    public override void Write(ReadOnlySpan<byte> buffer)
    {
        Check(buffer.Length); (_directory ?? destination).Write(buffer); _written += buffer.Length;
    }
    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
    {
        Check(buffer.Length); await (_directory ?? destination).WriteAsync(buffer, cancellationToken).ConfigureAwait(false); _written += buffer.Length;
    }
    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
        WriteAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
    /// <summary>Keep what is written from here on, the central directory a <see cref="System.IO.Compression.ZipArchive"/>
    /// writes when it closes, until <see cref="WriteCentralDirectoryAsync"/> sends it on.</summary>
    public void HoldCentralDirectory() => _directory = new MemoryStream();

    /// <summary>Send the held central directory on, recording every entry as made on Unix so the
    /// archive's bytes do not depend on the host that wrote it (issue #1659).</summary>
    public async ValueTask WriteCentralDirectoryAsync(CancellationToken cancellationToken)
    {
        var directory = _directory!.ToArray();
        _directory = null;
        ZipPackageOutputNormalizer.StampUnixHost(directory, _written - directory.Length);
        await destination.WriteAsync(directory, cancellationToken).ConfigureAwait(false);
    }

    private void Check(int count) => HistoryArchiveManifest.Budget(count <= maximum - _written, "History archive output exceeds its byte limit.");
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
}
