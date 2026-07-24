#nullable enable

using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Backfills the canonical <c>word/settings.xml</c> children that Word's compare output SYNTHESIZES for
/// every document even when the source carries an empty settings stub (verified against Word's compare
/// output). Sibling of <see cref="WordCompareFontTableBackfill"/> — the theme, docDefaults, and fontTable
/// backfills are the analogous single-owner passes; this completes the set for <c>settings.xml</c>.
/// <para><b>The load-bearing element is <c>compat/compatibilityMode</c>.</b> <c>compatibilityMode</c>
/// selects LibreOffice's layout-engine emulation (Word 2007/2010/2013+): Word's redline carries one and,
/// when our output does not, the two documents lay out under DIFFERENT engines and a rendered redline
/// diverges from Word's compare output even on a byte-identical body. The other three children —
/// <c>characterSpacingControl</c>/<c>themeFontLang</c>/<c>clrSchemeMapping</c> — are inert against
/// LibreOffice's defaults and are emitted purely for parity with what Word writes.</para>
/// <para><b>compatibilityMode value — an articulable rule, not a per-document constant.</b> Word keeps the
/// ORIGINAL document's mode when it has one, otherwise the revised document's, otherwise <c>12</c> (Word's
/// default for an unmarked .docx). Since our output is cloned from the LEFT/original, a mode the left already
/// carries is left untouched (the <c>hasMode</c> short-circuit); only when the left has none do we consult the
/// right, then fall back to <c>12</c>. Matches Word's compare output in the common cases; the remainder are
/// cases where Word overrides the source mode unpredictably. A genuine mode-15 document is therefore never
/// downgraded — correct production behavior, not just a test artifact.</para>
/// </summary>
internal static class WordCompareSettingsBackfill
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string CompatUri = "http://schemas.microsoft.com/office/word";

    /// <summary>Ensure the output's settings part carries Word's canonical compare settings, deriving
    /// <c>compatibilityMode</c> from the left (already on <paramref name="main"/>) then
    /// <paramref name="rightMain"/> then <c>12</c>.</summary>
    public static void Backfill(MainDocumentPart main, MainDocumentPart? rightMain)
    {
        var settingsPart = main.DocumentSettingsPart;
        if (settingsPart is null)
            return;
        var root = settingsPart.GetXDocument().Root;
        if (root is null || root.Name != W + "settings")
            return;

        var changed = false;
        // compat/compatibilityMode — the load-bearing one (see class remarks).
        changed |= EnsureCompatibilityMode(root, rightMain);
        // Inert-but-faithful canonical settings Word always writes.
        changed |= WordprocessingMLUtil.EnsureSettingsChildInOrder(root,
            new XElement(W + "characterSpacingControl", new XAttribute(W + "val", "doNotCompress")));
        changed |= WordprocessingMLUtil.EnsureSettingsChildInOrder(root,
            new XElement(W + "themeFontLang", new XAttribute(W + "val", "en-US")));
        changed |= WordprocessingMLUtil.EnsureSettingsChildInOrder(root, BuildClrSchemeMapping());

        if (changed)
            settingsPart.PutXDocument();
    }

    private static bool EnsureCompatibilityMode(XElement root, MainDocumentPart? rightMain)
    {
        var compat = root.Element(W + "compat");
        // Left (cloned onto `main`) already declares a mode → keep it verbatim (never downgrade).
        if (HasCompatSetting(compat, "compatibilityMode"))
            return false;

        var mode = ReadCompatibilityMode(rightMain) ?? "12";
        var hyphenationPresent = HasCompatSetting(compat, "useWord2013TrackBottomHyphenation");

        if (compat is null)
        {
            compat = new XElement(W + "compat");
            WordprocessingMLUtil.EnsureSettingsChildInOrder(root, compat);
        }
        compat.Add(CompatSetting("compatibilityMode", mode));
        if (!hyphenationPresent)
            compat.Add(CompatSetting("useWord2013TrackBottomHyphenation", "1"));
        return true;
    }

    private static bool HasCompatSetting(XElement? compat, string name) =>
        compat?.Elements(W + "compatSetting")
            .Any(cs => (string?)cs.Attribute(W + "name") == name) ?? false;

    private static XElement CompatSetting(string name, string val) =>
        new XElement(W + "compatSetting",
            new XAttribute(W + "name", name),
            new XAttribute(W + "uri", CompatUri),
            new XAttribute(W + "val", val));

    private static string? ReadCompatibilityMode(MainDocumentPart? mainPart)
    {
        var root = mainPart?.DocumentSettingsPart?.GetXDocument().Root;
        return root?.Element(W + "compat")?.Elements(W + "compatSetting")
            .FirstOrDefault(cs => (string?)cs.Attribute(W + "name") == "compatibilityMode")
            ?.Attribute(W + "val")?.Value;
    }

    /// <summary>The standard 1:1 theme→document color-slot mapping Word writes.</summary>
    private static XElement BuildClrSchemeMapping()
    {
        var el = new XElement(W + "clrSchemeMapping");
        el.Add(new XAttribute(W + "bg1", "light1"));
        el.Add(new XAttribute(W + "t1", "dark1"));
        el.Add(new XAttribute(W + "bg2", "light2"));
        el.Add(new XAttribute(W + "t2", "dark2"));
        el.Add(new XAttribute(W + "accent1", "accent1"));
        el.Add(new XAttribute(W + "accent2", "accent2"));
        el.Add(new XAttribute(W + "accent3", "accent3"));
        el.Add(new XAttribute(W + "accent4", "accent4"));
        el.Add(new XAttribute(W + "accent5", "accent5"));
        el.Add(new XAttribute(W + "accent6", "accent6"));
        el.Add(new XAttribute(W + "hyperlink", "hyperlink"));
        el.Add(new XAttribute(W + "followedHyperlink", "followedHyperlink"));
        return el;
    }
}
