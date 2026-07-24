#nullable enable

namespace Docxodus;

/// <summary>
/// Selects which comparison engine <see cref="DocxCompare.Compare"/> — and the CLI / WASM / npm
/// surfaces that route through it — use to redline two DOCX documents.
///
/// <para>The integer values are part of the contract: the WASM and npm surfaces marshal the
/// selector as an <c>int</c>, and <c>0</c> remains mapped to
/// <see cref="WmlComparer"/> for wire compatibility. Public surfaces that omit the selector choose
/// <see cref="DocxDiff"/> explicitly.</para>
///
/// <para><see cref="DocxDiff"/> is the default comparison engine. <see cref="WmlComparer"/> remains
/// available through an explicit selector for callers that require its historical behavior.</para>
/// </summary>
public enum ComparisonEngine
{
    /// <summary>The legacy <see cref="Docxodus.WmlComparer"/> engine (retained at wire value 0).</summary>
    WmlComparer = 0,

    /// <summary>The default <see cref="Docxodus.DocxDiff"/> IR diff engine.</summary>
    DocxDiff = 1,
}
