#nullable enable

using System;

namespace Docxodus.Ir;

/// <summary>
/// How <see cref="IrReader"/> treats tracked revisions (`w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`/
/// `w:rPrChange`/`w:pPrChange`) before reading the body (spec §5, rule N13).
/// </summary>
internal enum RevisionView
{
    /// <summary>Accept all revisions (insertions kept, deletions removed) before reading.</summary>
    Accept,

    /// <summary>Reject all revisions (insertions removed, deletions restored) before reading.</summary>
    Reject,

    /// <summary>Throw a <see cref="DocxodusException"/> if any revision markup is present.</summary>
    FailIfPresent,
}

/// <summary>
/// Which document scopes the reader walks. All flags are honored: <see cref="Body"/> reads
/// <c>w:body</c>; <see cref="HeadersFooters"/> reads the header/footer parts (scopes
/// <c>hdr1</c>/<c>ftr1</c>…); <see cref="Notes"/> reads footnotes + endnotes (scopes <c>fn</c>/
/// <c>en</c>); <see cref="Comments"/> reads the comments part (scope <c>cmt</c>) and records N15
/// comment-range targets during the body walk. The body is always read because
/// <see cref="IrDocument.Body"/> is required; an unselected non-body scope is emitted as an empty
/// store/list.
/// </summary>
[Flags]
internal enum IrScopes
{
    Body = 1,
    HeadersFooters = 2,
    Notes = 4,
    Comments = 8,
    All = Body | HeadersFooters | Notes | Comments,
}

/// <summary>Options controlling an <see cref="IrReader.Read"/> pass.</summary>
internal sealed class IrReaderOptions
{
    /// <summary>How tracked revisions are normalized before reading (default <see cref="RevisionView.Accept"/>).</summary>
    public RevisionView RevisionView { get; init; } = RevisionView.Accept;

    /// <summary>
    /// Which scopes to read. Defaults to <see cref="IrScopes.All"/>; in M1.1 only
    /// <see cref="IrScopes.Body"/> is honored and the remaining flags are accepted and ignored.
    /// </summary>
    public IrScopes Scopes { get; init; } = IrScopes.All;
}
