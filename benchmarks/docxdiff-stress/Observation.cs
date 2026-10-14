// One observation of one product on one document: everything the harness can see about that call,
// not just what the call returned.
//
// The differential used to digest the return value and nothing else, which is why the #616
// regression -- a fast path that skipped the compatibility pre-flight -- passed all 8,136 digests
// while genuinely changing behaviour. For two byte-identical documents "no revisions" is the right
// answer before and after; the return value never moved. The defect was on a channel nothing
// watched.
//
// Making the unit a record with one field per channel is the structural fix. Adding a channel is
// one field on one type, applied to every product and every document at once, rather than a new
// sink[...] family that someone has to remember to wire into each mode.

namespace Docxodus.Stress;

/// <summary>Everything observed about one product call. Every field is a digest or a short literal,
/// so a mismatch names the channel that moved rather than just saying "the digest changed".</summary>
internal readonly record struct Observation
{
    /// <summary>The return value's digest, or <c>FAIL &lt;Type&gt;: &lt;message&gt;</c>. A thrown
    /// exception is an outcome like any other: a build that throws something different, or stops
    /// throwing, is a regression.</summary>
    public required string Result { get; init; }

    /// <summary>The compatibility pre-flight's feature ids for this call, or <c>none</c>. Captured
    /// from the same call that produced <see cref="Result"/>, so a product that stops running the
    /// pre-flight records <c>none</c> where it used to record a report.</summary>
    public required string Warnings { get; init; }

    /// <summary>Whether the call left its inputs byte-for-byte unchanged: <c>clean</c>, or which
    /// side's bytes moved. <c>IrReader.Read</c> and <c>PreAccept</c> both promise this in prose and
    /// nothing verified it.</summary>
    public required string InputMutation { get; init; }

    /// <summary>Whether the product agrees with itself when the engine is reached a different way:
    /// <c>stable</c>, or a description of the disagreement. Guards the shared memoized snapshot
    /// <c>DocxDiffComparison</c> introduced.</summary>
    public required string OrderVariance { get; init; }

    public static Observation NotRun(string reason) => new()
    {
        Result = reason,
        Warnings = "n/a",
        InputMutation = "n/a",
        OrderVariance = "n/a",
    };

    /// <summary>Field-by-field difference against a recorded observation, or an empty list.</summary>
    public IReadOnlyList<string> DifferencesFrom(Observation expected)
    {
        var diffs = new List<string>();
        void Check(string channel, string want, string got)
        {
            if (!string.Equals(want, got, StringComparison.Ordinal))
                diffs.Add($"{channel}: expected {want} / actual {got}");
        }

        Check("result", expected.Result, Result);
        Check("warnings", expected.Warnings, Warnings);
        Check("inputMutation", expected.InputMutation, InputMutation);
        Check("orderVariance", expected.OrderVariance, OrderVariance);
        return diffs;
    }
}
