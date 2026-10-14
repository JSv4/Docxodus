#nullable enable

using System;

namespace Docxodus.Ir;

/// <summary>
/// The kind component of an IR anchor. The token strings (see <see cref="IrAnchor.KindToken"/>)
/// are the markdown-projection anchor kinds produced by <c>WmlToMarkdownConverter.KindFor</c>,
/// extended with <see cref="Img"/>/<see cref="Drw"/>/<see cref="Sdt"/>/<see cref="Unk"/> for
/// IR-internal use. <see cref="Col"/> is projection/index metadata for a table's
/// <c>w:gridCol</c>; it is part of the closed anchor vocabulary even though it is not an IR block.
/// </summary>
internal enum IrAnchorKind
{
    P,
    H,
    Li,
    Tbl,
    Tr,
    Tc,
    Col,
    Cmt,
    Fn,
    En,
    Img,
    Drw,
    Sec,
    Sdt,
    Unk,
}

/// <summary>
/// Stable identity of an addressable IR node. Renders as <c>kind:scope:unid</c> — the same
/// grammar as the markdown projection's public <c>Anchor.Id</c> (unid is a 32-char hex Unid).
/// </summary>
internal readonly record struct IrAnchor(IrAnchorKind Kind, string Scope, string Unid)
{
    public override string ToString() => $"{KindToken(Kind)}:{Scope}:{Unid}";

    /// <summary>Map a kind to its lowercase token (matches the markdown projection vocabulary).</summary>
    public static string KindToken(IrAnchorKind kind) => kind switch
    {
        IrAnchorKind.P => "p",
        IrAnchorKind.H => "h",
        IrAnchorKind.Li => "li",
        IrAnchorKind.Tbl => "tbl",
        IrAnchorKind.Tr => "tr",
        IrAnchorKind.Tc => "tc",
        IrAnchorKind.Col => "col",
        IrAnchorKind.Cmt => "cmt",
        IrAnchorKind.Fn => "fn",
        IrAnchorKind.En => "en",
        IrAnchorKind.Img => "img",
        IrAnchorKind.Drw => "drw",
        IrAnchorKind.Sec => "sec",
        IrAnchorKind.Sdt => "sdt",
        IrAnchorKind.Unk => "unk",
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unknown anchor kind."),
    };

    /// <summary>Inverse of <see cref="KindToken"/>. Throws on an unrecognized token.</summary>
    public static IrAnchorKind KindFromToken(string token) => token switch
    {
        "p" => IrAnchorKind.P,
        "h" => IrAnchorKind.H,
        "li" => IrAnchorKind.Li,
        "tbl" => IrAnchorKind.Tbl,
        "tr" => IrAnchorKind.Tr,
        "tc" => IrAnchorKind.Tc,
        "col" => IrAnchorKind.Col,
        "cmt" => IrAnchorKind.Cmt,
        "fn" => IrAnchorKind.Fn,
        "en" => IrAnchorKind.En,
        "img" => IrAnchorKind.Img,
        "drw" => IrAnchorKind.Drw,
        "sec" => IrAnchorKind.Sec,
        "sdt" => IrAnchorKind.Sdt,
        "unk" => IrAnchorKind.Unk,
        _ => throw new ArgumentException($"Unknown anchor kind token: '{token}'.", nameof(token)),
    };
}
