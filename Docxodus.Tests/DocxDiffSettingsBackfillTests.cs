#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Word's compare output SYNTHESIZES canonical <c>word/settings.xml</c> children for every document even
/// when the source carries an empty settings stub (verified against Word's compare output). The only
/// rendering-relevant one is <c>compat/compatibilityMode</c>, which selects LibreOffice's layout-engine
/// emulation — an output missing it lays out under a different engine than Word's redline, so a rendered
/// redline diverges from Word's compare output. See <c>WordCompareSettingsBackfill</c>.
/// <para>compatibilityMode rule (matches Word's compare output in the common cases): keep the ORIGINAL
/// (left) document's value when present, otherwise the revised (right) document's, otherwise <c>12</c>
/// (Word's default for an unmarked .docx). A genuine mode-15 document is therefore never downgraded.</para>
/// </summary>
public class DocxDiffSettingsBackfillTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string CompatUri = "http://schemas.microsoft.com/office/word";

    /// <summary>Build a minimal doc whose settings part carries <paramref name="settingsInner"/> verbatim
    /// (empty stub when null) — the shape real documents ship with.</summary>
    private static WmlDocument BuildDoc(string text, string? settingsInner)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            using var w = new StreamWriter(settingsPart.GetStream(FileMode.Create), new System.Text.UTF8Encoding(false));
            w.Write($"<w:settings xmlns:w=\"{W.NamespaceName}\">{settingsInner ?? string.Empty}</w:settings>");
        }
        return new WmlDocument("d.docx", stream.ToArray());
    }

    private static string CompatBlock(string mode) =>
        $"<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"{CompatUri}\" w:val=\"{mode}\"/></w:compat>";

    private static XElement OutputSettings(WmlDocument redline)
    {
        using var ms = new MemoryStream(redline.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        using var reader = new StreamReader(doc.MainDocumentPart!.DocumentSettingsPart!.GetStream());
        return XDocument.Parse(reader.ReadToEnd()).Root!;
    }

    private static string? CompatMode(XElement settings) =>
        settings.Element(W + "compat")?.Elements(W + "compatSetting")
            .FirstOrDefault(cs => (string?)cs.Attribute(W + "name") == "compatibilityMode")
            ?.Attribute(W + "val")?.Value;

    [Fact]
    public void EmptySettings_BothSides_BackfillsCanonicalDefaults()
    {
        var left = BuildDoc("Old text here.", settingsInner: null);
        var right = BuildDoc("New replacement words entirely.", settingsInner: null);

        var settings = OutputSettings(DocxDiff.Compare(left, right));

        // compat/compatibilityMode = 12 (Word's default for an unmarked .docx) — the load-bearing setting.
        Assert.Equal("12", CompatMode(settings));
        Assert.Contains(settings.Element(W + "compat")!.Elements(W + "compatSetting"),
            cs => (string?)cs.Attribute(W + "name") == "useWord2013TrackBottomHyphenation");
        // The inert-but-faithful canonical settings Word always writes.
        Assert.Equal("doNotCompress", (string?)settings.Element(W + "characterSpacingControl")?.Attribute(W + "val"));
        Assert.Equal("en-US", (string?)settings.Element(W + "themeFontLang")?.Attribute(W + "val"));
        Assert.Equal("light1", (string?)settings.Element(W + "clrSchemeMapping")?.Attribute(W + "bg1"));
    }

    [Fact]
    public void LeftHasMode15_IsPreserved_NeverDowngraded()
    {
        var left = BuildDoc("Old text here.", settingsInner: CompatBlock("15"));
        var right = BuildDoc("New replacement words entirely.", settingsInner: null);

        var settings = OutputSettings(DocxDiff.Compare(left, right));

        // The original's genuine mode-15 survives (output clones the left); it is not overwritten with 12.
        Assert.Equal("15", CompatMode(settings));
        // ... and only ONE compatibilityMode is present (no duplicate injected).
        Assert.Single(settings.Element(W + "compat")!.Elements(W + "compatSetting"),
            cs => (string?)cs.Attribute(W + "name") == "compatibilityMode");
    }

    [Fact]
    public void LeftEmpty_RightHasMode14_AdoptsRightMode()
    {
        var left = BuildDoc("Old text here.", settingsInner: null);
        var right = BuildDoc("New replacement words entirely.", settingsInner: CompatBlock("14"));

        var settings = OutputSettings(DocxDiff.Compare(left, right));

        // base(empty)||next(14)||12 -> 14.
        Assert.Equal("14", CompatMode(settings));
    }

    [Fact]
    public void BackfilledSettings_AreSchemaOrdered_AndValid()
    {
        var left = BuildDoc("Old text here.", settingsInner: null);
        var right = BuildDoc("New replacement words entirely.", settingsInner: null);

        var redline = DocxDiff.Compare(left, right);
        var settings = OutputSettings(redline);
        var names = settings.Elements().Select(e => e.Name.LocalName).ToList();

        // CT_Settings schema order among the children we touch: characterSpacingControl < compat <
        // themeFontLang < clrSchemeMapping (defaultTabStop, if present, precedes them all).
        int csc = names.IndexOf("characterSpacingControl");
        int compat = names.IndexOf("compat");
        int tfl = names.IndexOf("themeFontLang");
        int clr = names.IndexOf("clrSchemeMapping");
        Assert.True(csc >= 0 && csc < compat, $"characterSpacingControl not before compat: [{string.Join(", ", names)}]");
        Assert.True(compat < tfl, $"compat not before themeFontLang: [{string.Join(", ", names)}]");
        Assert.True(tfl < clr, $"themeFontLang not before clrSchemeMapping: [{string.Join(", ", names)}]");

        // The whole output validates against the Office schema (no invalid settings shape).
        using var ms = new MemoryStream(redline.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc)
            .Where(e => e.Path?.XPath?.Contains("settings") == true)
            .Select(e => $"{e.Id}: {e.Description}")
            .ToList();
        Assert.Empty(errors);
    }
}
