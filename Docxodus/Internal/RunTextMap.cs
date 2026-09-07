// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// Flat-text view of a paragraph (or other block element) with an offset map
/// back to the originating <c>&lt;w:r&gt;</c> runs. Lets callers run regex/text
/// searches across run boundaries and resolve each match to the runs it spans.
///
/// Mirrors the text-map+offset-map pattern in <see cref="ExternalAnnotationProjector"/>
/// but operates on OOXML run elements (one entry per run) instead of XHTML text nodes.
/// </summary>
internal static class RunTextMap
{
    /// <summary>One run's contribution to the flat text.</summary>
    public readonly record struct RunSegment(
        XElement Run,
        int StartOffsetInBlock,
        int Length)
    {
        public int EndOffsetInBlock => StartOffsetInBlock + Length;
    }

    /// <summary>Flat text + the segment list, in document order.</summary>
    public readonly record struct Map(string FlatText, IReadOnlyList<RunSegment> Segments);

    /// <summary>
    /// Builds the flat text + segment map for a block element. Walks runs via the
    /// same <see cref="DocxSession.InlineRuns"/> pass the session uses, so the
    /// flat text matches what every other session op operates on.
    /// </summary>
    public static Map Build(XElement blockElement)
    {
        var segments = new List<RunSegment>();
        var offset = 0;

        foreach (var run in DocxSession.InlineRuns(blockElement))
        {
            var runText = DocxSession.RunText(run);
            if (runText.Length == 0) continue;
            segments.Add(new RunSegment(run, offset, runText.Length));
            offset += runText.Length;
        }

        var flat = string.Concat(segments.Select(s => DocxSession.RunText(s.Run)));
        return new Map(flat, segments);
    }

    /// <summary>
    /// Given a character range <c>[start, start+length)</c> in the block's flat text,
    /// returns the run segments that overlap it, each with the offset+length WITHIN
    /// the run's own text that participates in the range. The list is in document order
    /// and is never empty when the range is in bounds.
    /// </summary>
    public static List<(RunSegment Segment, int OffsetInRun, int Length)> ResolveRange(
        Map map,
        int start,
        int length)
    {
        var result = new List<(RunSegment, int, int)>();
        if (length <= 0) return result;
        var end = start + length;

        foreach (var seg in map.Segments)
        {
            if (seg.EndOffsetInBlock <= start) continue;
            if (seg.StartOffsetInBlock >= end) break;

            var overlapStart = System.Math.Max(start, seg.StartOffsetInBlock);
            var overlapEnd = System.Math.Min(end, seg.EndOffsetInBlock);
            var offsetInRun = overlapStart - seg.StartOffsetInBlock;
            var overlapLen = overlapEnd - overlapStart;
            result.Add((seg, offsetInRun, overlapLen));
        }

        return result;
    }
}
