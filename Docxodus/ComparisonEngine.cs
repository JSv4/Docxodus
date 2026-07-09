#nullable enable

namespace Docxodus;

/// <summary>
/// Selects which comparison engine <see cref="DocxCompare.Compare"/> — and the CLI / WASM / npm
/// surfaces that route through it — use to redline two DOCX documents.
///
/// <para>The integer values are part of the contract: the WASM and npm surfaces marshal the
/// selector as an <c>int</c>, and <c>0</c> (the default / unset value) MUST map to
/// <see cref="WmlComparer"/> so that omitting the selector reproduces today's behavior exactly.</para>
///
/// <para><see cref="WmlComparer"/> is the blessed default. <see cref="DocxDiff"/> is the newer IR diff
/// engine — a production-candidate that becomes the default only after the Word manual-verification
/// checklist clears and a burn-in period (decision D4). This selector is the seam that lets that flip
/// happen in one line; M-B does NOT flip it.</para>
/// </summary>
public enum ComparisonEngine
{
    /// <summary>The default, blessed <see cref="Docxodus.WmlComparer"/> engine.</summary>
    WmlComparer = 0,

    /// <summary>The <see cref="Docxodus.DocxDiff"/> IR diff engine (production-candidate; opt-in).</summary>
    DocxDiff = 1,
}
