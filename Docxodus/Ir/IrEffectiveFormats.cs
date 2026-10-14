#nullable enable

using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace Docxodus.Ir;

/// <summary>
/// Non-destructive effective-format resolution (§5.2): given a paragraph (or a run's direct
/// <see cref="IrRunFormat"/>), compute the <em>effective</em> paragraph/run format by cascading
/// document defaults → the style chain → direct properties, without mutating the IR. The resolver
/// is built once per <see cref="IrDocument"/> and reads its style registry, theme fonts, and the
/// reused <see cref="IrReader.MapParaFormat"/>/<see cref="IrReader.MapRunFormat"/> mappers — so a
/// style layer's cloned <c>w:pPr</c>/<c>w:rPr</c> is mapped through the exact same code path as a
/// paragraph's direct props (one mapping, never duplicated).
/// </summary>
/// <remarks>
/// <para><b>Cascade (paragraph).</b> docDefaults <c>w:pPr</c> → the paragraph-style chain (the
/// paragraph's <c>pStyle</c> — or, when absent, the registry's default paragraph style — then its
/// <c>basedOn</c> ancestors), applied ROOT-FIRST so a derived style overrides its base → the
/// paragraph's direct <see cref="IrParagraph.Format"/>. Each layer merges per-field: a later
/// layer's NON-NULL field wins; a null field inherits from the layer below.</para>
///
/// <para><b>Cascade (run).</b> docDefaults <c>w:rPr</c> → the paragraph-style chain's <c>w:rPr</c>
/// (paragraph styles carry run props for their runs) → the character-style chain's <c>w:rPr</c>
/// (only when <see cref="IrRunFormat.StyleId"/> names a <c>type="character"</c> style) → the run's
/// direct fields. Same non-null-wins merge.</para>
///
/// <para><b>Toggle properties.</b> Bold/Italic/Strike/Caps/etc. follow LAST-WRITER-WINS at this
/// fidelity tier: the topmost layer that sets the toggle decides it. This diverges from OOXML's
/// real toggle semantics, where a toggle property XOR-aggregates across style layers (e.g. bold in
/// the base style and bold in a derived style cancel to non-bold). Matching Word exactly here would
/// require tracking per-layer toggle contributions rather than a single merged value.
/// TODO(M1.4+): aggregate toggle properties with OOXML XOR semantics across style layers instead of
/// last-writer-wins; today a derived style re-asserting bold cannot un-bold an already-bold base.</para>
///
/// <para><b>Theme fonts.</b> When an rPr layer carries <c>w:rFonts/@w:asciiTheme</c> (majorHAnsi,
/// majorAscii, minorHAnsi, minorAscii, …), it is resolved at THAT layer through
/// <see cref="IrThemeFonts"/> — major* → <see cref="IrThemeFonts.MajorAscii"/>, minor* →
/// <see cref="IrThemeFonts.MinorAscii"/> — into the layer's <see cref="IrRunFormat.FontAscii"/>.
/// Theme resolution belongs to effective formats only (§5.2): the direct reader stores the literal
/// <c>@w:ascii</c> and never the theme indirection. A literal <c>@w:ascii</c> on the same layer
/// beats the theme (it is mapped into <see cref="IrRunFormat.FontAscii"/> directly, so the theme is
/// only consulted when that field is null on the layer). Theme resolution applies to the layers
/// the resolver maps from XElements — docDefaults and the style chains — because those still hold
/// the raw <c>w:rFonts</c>. The DIRECT run layer arrives as an already-mapped
/// <see cref="IrRunFormat"/> whose literal <c>@w:ascii</c> the reader captured but whose
/// <c>@w:asciiTheme</c> indirection it dropped (the IR does not model it); a direct run that
/// expresses its font ONLY via a theme token therefore contributes no direct font and inherits the
/// resolved theme face from the layers below. TODO(M1.4+): thread the direct rPr through so a
/// theme-only direct run resolves at its own layer too.</para>
///
/// <para><b>UnmodeledDigest.</b> Effective resolution covers the MODELED fields only. The effective
/// record's <see cref="IrParaFormat.UnmodeledDigest"/>/<see cref="IrRunFormat.UnmodeledDigest"/> is
/// carried straight from the DIRECT record — it is NOT a cascade of style-layer unmodeled digests.
/// Callers needing unmodeled-property cascade must consult the style XElements themselves.</para>
///
/// <para><b>Caching / thread-safety.</b> Per-instance memo caches key on styleId and hold the
/// style-CHAIN contribution (docDefaults ⊕ the basedOn walk), NOT the merged-with-direct result —
/// so two paragraphs of the same style share the chain computation but still merge their own direct
/// props. Access is guarded by a single lock: resolution is read-mostly and the critical section is
/// a dictionary lookup plus an occasional chain build, so a simple lock is ample; lock-free or
/// striped locking would be premature.</para>
/// </remarks>
internal sealed class IrEffectiveFormats
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const int MaxStyleDepth = 16;

    private readonly IrStyleRegistry _styles;
    private readonly IrThemeFonts _theme;

    // Memoized style-CHAIN contributions (docDefaults ⊕ basedOn walk), keyed by the leaf styleId.
    // "" is the key for "no style" (docDefaults only). These hold the chain result BEFORE merging
    // the paragraph/run's own direct props, so they are shared across all blocks of a style.
    private readonly Dictionary<string, IrParaFormat> _paraChainCache = new(StringComparer.Ordinal);
    private readonly Dictionary<string, IrRunFormat> _runParaChainCache = new(StringComparer.Ordinal);
    private readonly Dictionary<string, IrRunFormat> _runCharChainCache = new(StringComparer.Ordinal);
    private readonly object _lock = new();

    public IrEffectiveFormats(IrDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        _styles = document.Styles;
        _theme = document.ThemeFonts;
    }

    // --- paragraph --------------------------------------------------------

    /// <summary>
    /// Resolve the effective paragraph format for <paramref name="p"/>: docDefaults pPr →
    /// paragraph-style chain → the paragraph's direct <see cref="IrParagraph.Format"/>. The
    /// effective <see cref="IrParaFormat.UnmodeledDigest"/> is the direct format's digest.
    /// </summary>
    public IrParaFormat ResolveParagraph(IrParagraph p)
    {
        ArgumentNullException.ThrowIfNull(p);

        var direct = p.Format;
        var styleId = direct.StyleId ?? _styles.DefaultParagraphStyleId;
        var chain = ParaStyleChain(styleId ?? "");
        return MergePara(chain, direct, direct.UnmodeledDigest);
    }

    private IrParaFormat ParaStyleChain(string styleId)
    {
        lock (_lock)
        {
            if (_paraChainCache.TryGetValue(styleId, out var cached))
                return cached;

            // Layer 0: docDefaults pPr.
            var acc = IrReader.MapParaFormat(_styles.DocDefaultsPPr);

            // Layers 1..N: the basedOn chain, applied ROOT-FIRST (derived overrides base). Collect
            // the chain leaf→root, then fold in reverse.
            foreach (var style in StyleChainRootFirst(styleId))
            {
                var layer = IrReader.MapParaFormat(style.PPr);
                acc = MergePara(acc, layer, acc.UnmodeledDigest);
            }

            _paraChainCache[styleId] = acc;
            return acc;
        }
    }

    // --- run --------------------------------------------------------------

    /// <summary>
    /// Resolve the effective run format for a run whose direct props are <paramref name="direct"/>,
    /// living in paragraph <paramref name="p"/>: docDefaults rPr → the paragraph-style chain's rPr →
    /// the character-style chain's rPr (when <paramref name="direct"/>.<see cref="IrRunFormat.StyleId"/>
    /// names a character style) → the direct fields. The effective
    /// <see cref="IrRunFormat.UnmodeledDigest"/> is the direct format's digest.
    /// </summary>
    public IrRunFormat ResolveRun(IrParagraph p, IrRunFormat direct)
    {
        ArgumentNullException.ThrowIfNull(p);
        ArgumentNullException.ThrowIfNull(direct);

        var paraStyleId = p.Format.StyleId ?? _styles.DefaultParagraphStyleId;
        var acc = RunParaStyleChain(paraStyleId ?? "");

        // Character-style chain (only when the run's rStyle names a character-typed style).
        if (direct.StyleId is { } charStyleId && IsCharacterStyle(charStyleId))
        {
            var charChain = RunCharStyleChain(charStyleId);
            acc = MergeRun(acc, charChain, acc.UnmodeledDigest);
        }

        return MergeRun(acc, direct, direct.UnmodeledDigest);
    }

    private IrRunFormat RunParaStyleChain(string styleId)
    {
        lock (_lock)
        {
            if (_runParaChainCache.TryGetValue(styleId, out var cached))
                return cached;

            var acc = MapRunWithTheme(_styles.DocDefaultsRPr);
            foreach (var style in StyleChainRootFirst(styleId))
            {
                var layer = MapRunWithTheme(style.RPr);
                acc = MergeRun(acc, layer, acc.UnmodeledDigest);
            }

            _runParaChainCache[styleId] = acc;
            return acc;
        }
    }

    private IrRunFormat RunCharStyleChain(string styleId)
    {
        lock (_lock)
        {
            if (_runCharChainCache.TryGetValue(styleId, out var cached))
                return cached;

            // Character-style chain does NOT re-apply docDefaults (that is the paragraph chain's
            // layer 0); it is just the rStyle's basedOn walk, root-first.
            IrRunFormat acc = new() { UnmodeledDigest = MapRunWithTheme(null).UnmodeledDigest };
            foreach (var style in StyleChainRootFirst(styleId))
            {
                var layer = MapRunWithTheme(style.RPr);
                acc = MergeRun(acc, layer, acc.UnmodeledDigest);
            }

            _runCharChainCache[styleId] = acc;
            return acc;
        }
    }

    // --- style-chain walk -------------------------------------------------

    /// <summary>
    /// The style chain for <paramref name="styleId"/> in ROOT-FIRST order (base ancestor first,
    /// leaf style last) so folding the layers left-to-right lets a derived style override its base.
    /// Cycle-guarded and depth-capped at <see cref="MaxStyleDepth"/> (same discipline as the
    /// reader's <c>ResolveStyleNumPr</c>). An empty or unknown id yields no styles.
    /// </summary>
    private List<IrStyle> StyleChainRootFirst(string styleId)
    {
        var leafToRoot = new List<IrStyle>();
        if (string.IsNullOrEmpty(styleId))
            return leafToRoot;

        var visited = new HashSet<string>(StringComparer.Ordinal);
        string? current = styleId;
        for (int i = 0; i < MaxStyleDepth && current is not null; i++)
        {
            if (!visited.Add(current))
                break; // cycle guard: a style basedOn itself (directly or transitively) stops here.
            if (!_styles.Styles.TryGetValue(current, out var style))
                break; // broken basedOn reference: stop, keep what we have.
            leafToRoot.Add(style);
            current = style.BasedOn;
        }

        leafToRoot.Reverse(); // → root-first.
        return leafToRoot;
    }

    private bool IsCharacterStyle(string styleId) =>
        _styles.Styles.TryGetValue(styleId, out var s) && s.Type == "character";

    // --- theme-resolving run mapper ---------------------------------------

    /// <summary>
    /// Map a layer's <c>w:rPr</c> via the reused <see cref="IrReader.MapRunFormat"/>, then resolve a
    /// theme-font indirection (<c>w:rFonts/@w:asciiTheme</c>) into <see cref="IrRunFormat.FontAscii"/>
    /// when that field is still null (a literal <c>@w:ascii</c> on the layer wins). major* themes
    /// resolve to <see cref="IrThemeFonts.MajorAscii"/>, minor* to <see cref="IrThemeFonts.MinorAscii"/>.
    /// </summary>
    private IrRunFormat MapRunWithTheme(XElement? rPr)
    {
        var mapped = IrReader.MapRunFormat(rPr);
        if (rPr is null || mapped.FontAscii is not null)
            return mapped;

        var themeAttr = (string?)rPr.Element(W + "rFonts")?.Attribute(W + "asciiTheme");
        var resolved = ResolveThemeFont(themeAttr);
        return resolved is null ? mapped : mapped with { FontAscii = resolved };
    }

    private string? ResolveThemeFont(string? asciiTheme)
    {
        if (asciiTheme is null)
            return null;
        // OOXML theme tokens: majorHAnsi/majorAscii/majorBidi/majorEastAsia → major scheme face;
        // minorHAnsi/minorAscii/… → minor scheme face. We only model the Latin/ASCII face.
        if (asciiTheme.StartsWith("major", StringComparison.Ordinal))
            return _theme.MajorAscii;
        if (asciiTheme.StartsWith("minor", StringComparison.Ordinal))
            return _theme.MinorAscii;
        return null;
    }

    // --- per-field merges (later non-null wins) ---------------------------

    /// <summary>
    /// Merge a <paramref name="lower"/> paragraph layer with a higher <paramref name="upper"/> one:
    /// each field takes the upper layer's value when non-null, else inherits the lower. Toggles are
    /// last-writer-wins (see the class remarks). The merged record carries
    /// <paramref name="unmodeledDigest"/> verbatim.
    /// </summary>
    private static IrParaFormat MergePara(IrParaFormat lower, IrParaFormat upper, IrHash unmodeledDigest) =>
        new()
        {
            StyleId = upper.StyleId ?? lower.StyleId,
            Justification = upper.Justification ?? lower.Justification,
            IndentLeftTwips = upper.IndentLeftTwips ?? lower.IndentLeftTwips,
            IndentRightTwips = upper.IndentRightTwips ?? lower.IndentRightTwips,
            IndentFirstLineTwips = upper.IndentFirstLineTwips ?? lower.IndentFirstLineTwips,
            SpacingBeforeTwips = upper.SpacingBeforeTwips ?? lower.SpacingBeforeTwips,
            SpacingAfterTwips = upper.SpacingAfterTwips ?? lower.SpacingAfterTwips,
            LineSpacing = upper.LineSpacing ?? lower.LineSpacing,
            OutlineLevel = upper.OutlineLevel ?? lower.OutlineLevel,
            KeepNext = upper.KeepNext ?? lower.KeepNext,
            KeepLines = upper.KeepLines ?? lower.KeepLines,
            PageBreakBefore = upper.PageBreakBefore ?? lower.PageBreakBefore,
            Shading = upper.Shading ?? lower.Shading,
            UnmodeledDigest = unmodeledDigest,
        };

    /// <summary>
    /// Merge a <paramref name="lower"/> run layer with a higher <paramref name="upper"/> one: each
    /// field takes the upper layer's value when non-null, else inherits the lower. Toggles are
    /// last-writer-wins (see the class remarks). The merged record carries
    /// <paramref name="unmodeledDigest"/> verbatim.
    /// </summary>
    private static IrRunFormat MergeRun(IrRunFormat lower, IrRunFormat upper, IrHash unmodeledDigest) =>
        new()
        {
            StyleId = upper.StyleId ?? lower.StyleId,
            Bold = upper.Bold ?? lower.Bold,
            Italic = upper.Italic ?? lower.Italic,
            Underline = upper.Underline ?? lower.Underline,
            Strike = upper.Strike ?? lower.Strike,
            DoubleStrike = upper.DoubleStrike ?? lower.DoubleStrike,
            VertAlign = upper.VertAlign ?? lower.VertAlign,
            FontAscii = upper.FontAscii ?? lower.FontAscii,
            SizeHalfPoints = upper.SizeHalfPoints ?? lower.SizeHalfPoints,
            ColorHex = upper.ColorHex ?? lower.ColorHex,
            Highlight = upper.Highlight ?? lower.Highlight,
            Shading = upper.Shading ?? lower.Shading,
            Caps = upper.Caps ?? lower.Caps,
            SmallCaps = upper.SmallCaps ?? lower.SmallCaps,
            Vanish = upper.Vanish ?? lower.Vanish,
            UnmodeledDigest = unmodeledDigest,
        };
}
