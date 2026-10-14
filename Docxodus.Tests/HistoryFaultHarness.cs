#nullable enable

using Docxodus.History;

namespace Docxodus.Tests;

/// <summary>Real reference adapters beneath deterministic faults; shared by history/backend fuzzers.</summary>
internal sealed class HistoryFaultHarness : IHistoryBlobStore, IHistoryHeadInitializer, IDisposable
{
    internal enum Fault { None, BeforeBlob, AfterBlob, BeforeHead, AfterHead, CancelBeforeHead, CancelAfterHead }
    private readonly IHistoryBlobStore _blobs;
    private readonly IHistoryHeadStore _heads;
    private readonly string? _directory;
    private Fault _fault;
    private int _at;
    internal int Writes, Publications;
    internal string? DamagedDigest;
    internal bool Missing;
    internal CancellationTokenSource? Cancellation;
    private readonly AsyncLocal<string?> _contender = new();
    private PublicationGate? _publicationGate;

    internal void HoldCompetingPublications(string firstTag) => _publicationGate = new PublicationGate(firstTag);
    internal void ReleaseCompetingPublications() => _publicationGate = null;
    internal async Task<T> AsContenderAsync<T>(string tag, Func<Task<T>> action)
    {
        var previous = _contender.Value; _contender.Value = tag;
        try { return await action(); }
        finally { _contender.Value = previous; }
    }

    private sealed class PublicationGate(string firstTag)
    {
        private readonly TaskCompletionSource _bothArrived = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource _firstFinished = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _arrived;
        internal async Task EnterAsync(string? tag, CancellationToken cancellationToken)
        {
            if (tag is null) throw new InvalidOperationException("A gated publication must identify its contender.");
            if (Interlocked.Increment(ref _arrived) == 2) _bothArrived.TrySetResult();
            await _bothArrived.Task.WaitAsync(TimeSpan.FromSeconds(30), cancellationToken);
            if (tag != firstTag) await _firstFinished.Task.WaitAsync(TimeSpan.FromSeconds(30), cancellationToken);
        }
        internal void Exit(string? tag) { if (tag == firstTag) _firstFinished.TrySetResult(); }
    }

    internal HistoryFaultHarness(bool filesystem = false)
    {
        if (filesystem)
        {
            _directory = Directory.CreateTempSubdirectory("history-fault-fuzz-").FullName;
            _blobs = new FileHistoryBlobStore(Path.Combine(_directory, "blobs"));
            _heads = new FileHistoryHeadStore(Path.Combine(_directory, "heads"));
        }
        else { _blobs = new MemoryHistoryBlobStore(); _heads = new MemoryHistoryHeadStore(); }
    }

    internal void Arm(Fault fault, int at = 1)
    { _fault = fault; _at = at; Writes = 0; Publications = 0; }

    public async ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken = default)
    {
        var at = Interlocked.Increment(ref Writes);
        if (_fault == Fault.BeforeBlob && at == _at) throw new IOException("Injected pre-blob fault " + at);
        await _blobs.PutAsync(reference, content, cancellationToken);
        if (_fault == Fault.AfterBlob && at == _at) throw new IOException("Injected post-blob fault " + at);
    }

    public async ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default)
    {
        if (reference.Digest.Value != DamagedDigest) return await _blobs.OpenReadAsync(reference, cancellationToken);
        if (Missing) return null;
        using var input = await _blobs.OpenReadAsync(reference, cancellationToken);
        using var buffer = new MemoryStream(); await input!.CopyToAsync(buffer, cancellationToken);
        var bytes = buffer.ToArray(); bytes[bytes.Length / 2] ^= 1;
        return new MemoryStream(bytes, writable: false);
    }

    public ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken = default) =>
        _heads.ReadAsync(documentId, cancellationToken);

    public async ValueTask<HistoryHeadInitializationResult> TryInitializeAsync(string documentId, HistoryHead head,
        CancellationToken cancellationToken = default)
    {
        Interlocked.Increment(ref Publications);
        if (_fault == Fault.BeforeHead) throw new IOException("Injected failure before initialization.");
        if (_fault == Fault.CancelBeforeHead) Cancellation!.Cancel();
        var result = await ((IHistoryHeadInitializer)_heads).TryInitializeAsync(documentId, head, cancellationToken);
        if (result.Initialized && _fault == Fault.AfterHead) throw new IOException("Injected lost initialization acknowledgement.");
        if (result.Initialized && _fault == Fault.CancelAfterHead) Cancellation!.Cancel();
        return result;
    }

    public async ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected,
        HistoryBlobReference state, CancellationToken cancellationToken = default)
    {
        Interlocked.Increment(ref Publications);
        var gate = _publicationGate;
        if (gate is not null) await gate.EnterAsync(_contender.Value, cancellationToken);
        try
        {
            if (_fault == Fault.BeforeHead) throw new IOException("Injected failure before head CAS.");
            if (_fault == Fault.CancelBeforeHead) Cancellation!.Cancel();
            var result = await _heads.TryAdvanceAsync(documentId, expected, state, cancellationToken);
            if (result is not null && _fault == Fault.AfterHead) throw new IOException("Injected lost head acknowledgement.");
            if (result is not null && _fault == Fault.CancelAfterHead) Cancellation!.Cancel();
            return result;
        }
        finally { gate?.Exit(_contender.Value); }
    }

    public void Dispose()
    {
        Cancellation?.Dispose();
        if (_directory is not null) Directory.Delete(_directory, recursive: true);
    }
}
