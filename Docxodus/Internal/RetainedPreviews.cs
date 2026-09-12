// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Security.Cryptography;

namespace Docxodus.Internal;

/// <summary>
/// One successful isolated preview kept for a later guarded commit (issue #760): the shadow's
/// final package checkpoint, the live state it was predicted from, and the receipt it produced.
/// <see cref="Result"/> is null when the transport composed the receipt itself (the browser
/// client); the commit then returns a bare envelope the client merges with its own receipt.
/// </summary>
internal sealed record RetainedPreview(
    MutationPreviewRetention Retention,
    DocxSession.DocumentSnapshot Snapshot,
    string PackageHash,
    TrackedChangeMode TrackedChanges,
    string? RevisionAuthor,
    MutationBatchResult? Result)
{
    internal long ApproximateBytes => Snapshot.ApproximateBytes;
}

/// <summary>
/// Bounded, per-session store of previews retained for commit. A retained preview is a whole
/// package, so retention is bounded three ways: a count, a byte budget, and a time to live.
/// Eviction is oldest-first; a preview that alone exceeds the byte budget is refused rather
/// than evicting everything else. The store never touches the live package or its history —
/// it only holds bytes a commit may later restore.
/// </summary>
internal sealed class RetainedPreviews
{
    public const int DefaultCapacity = 8;
    public const long DefaultByteBudget = 64L * 1024 * 1024;
    public static readonly TimeSpan DefaultTimeToLive = TimeSpan.FromMinutes(15);

    private readonly int _capacity;
    private readonly long _byteBudget;
    private readonly TimeSpan _timeToLive;
    private readonly Func<DateTimeOffset> _utcNow;
    private readonly object _gate = new();

    // Insertion order, oldest first; eviction and purge walk from the front.
    private readonly List<RetainedPreview> _entries = new();
    private long _retainedBytes;

    public RetainedPreviews(
        int capacity = DefaultCapacity,
        long byteBudget = DefaultByteBudget,
        TimeSpan? timeToLive = null,
        Func<DateTimeOffset>? utcNow = null)
    {
        if (capacity < 1) throw new ArgumentOutOfRangeException(nameof(capacity));
        if (byteBudget < 1) throw new ArgumentOutOfRangeException(nameof(byteBudget));
        _capacity = capacity;
        _byteBudget = byteBudget;
        _timeToLive = timeToLive ?? DefaultTimeToLive;
        if (_timeToLive <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(timeToLive));
        _utcNow = utcNow ?? (static () => DateTimeOffset.UtcNow);
    }

    public int Capacity => _capacity;

    public long ByteBudget => _byteBudget;

    public TimeSpan TimeToLive => _timeToLive;

    public int Count
    {
        get
        {
            lock (_gate)
            {
                Purge();
                return _entries.Count;
            }
        }
    }

    public long RetainedBytes
    {
        get
        {
            lock (_gate)
            {
                Purge();
                return _retainedBytes;
            }
        }
    }

    public DateTimeOffset ExpiryFromNow() => _utcNow() + _timeToLive;

    public static string NewPreviewId() =>
        "pv-" + Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant();

    /// <summary>
    /// Retain a preview, evicting the oldest entries until the count and byte budget allow it.
    /// Returns false, retaining nothing, when the preview alone exceeds the byte budget.
    /// </summary>
    public bool Add(RetainedPreview preview)
    {
        ArgumentNullException.ThrowIfNull(preview);
        lock (_gate)
        {
            Purge();
            var bytes = preview.ApproximateBytes;
            if (bytes > _byteBudget) return false;
            while (_entries.Count > 0
                && (_entries.Count >= _capacity || _retainedBytes + bytes > _byteBudget))
            {
                EvictAt(0);
            }
            _entries.Add(preview);
            _retainedBytes += bytes;
            return true;
        }
    }

    /// <summary>Look a preview up without consuming it. Expired entries are never returned.</summary>
    public bool TryGet(string previewId, out RetainedPreview preview)
    {
        lock (_gate)
        {
            Purge();
            var index = IndexOf(previewId);
            preview = index < 0 ? null! : _entries[index];
            return index >= 0;
        }
    }

    public bool Remove(string previewId)
    {
        lock (_gate)
        {
            var index = IndexOf(previewId);
            if (index < 0) return false;
            EvictAt(index);
            return true;
        }
    }

    public void Clear()
    {
        lock (_gate)
        {
            _entries.Clear();
            _retainedBytes = 0;
        }
    }

    private int IndexOf(string previewId)
    {
        for (int i = 0; i < _entries.Count; i++)
            if (string.Equals(_entries[i].Retention.PreviewId, previewId, StringComparison.Ordinal))
                return i;
        return -1;
    }

    private void Purge()
    {
        var now = _utcNow();
        for (int i = _entries.Count - 1; i >= 0; i--)
            if (_entries[i].Retention.ExpiresAt <= now) EvictAt(i);
    }

    private void EvictAt(int index)
    {
        _retainedBytes -= _entries[index].ApproximateBytes;
        _entries.RemoveAt(index);
    }
}
