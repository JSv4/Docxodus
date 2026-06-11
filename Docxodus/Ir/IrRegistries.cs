#nullable enable

namespace Docxodus.Ir;

/// <summary>
/// Resolved style registry. Empty placeholder for M1.1; populated when style/cascade resolution
/// lands in M1.3.
/// </summary>
internal sealed record IrStyleRegistry
{
    public static readonly IrStyleRegistry Empty = new();
}

/// <summary>
/// Resolved numbering registry. Empty placeholder for M1.1; populated when numbering resolution
/// lands in M1.3.
/// </summary>
internal sealed record IrNumberingRegistry
{
    public static readonly IrNumberingRegistry Empty = new();
}

/// <summary>
/// Resolved theme fonts. Empty placeholder for M1.1; populated when theme-font resolution lands
/// in M1.3.
/// </summary>
internal sealed record IrThemeFonts
{
    public static readonly IrThemeFonts Empty = new();
}
