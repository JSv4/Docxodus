#nullable enable
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffCompatibilityTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string V = "urn:schemas-microsoft-com:vml";
    private const string O = "urn:schemas-microsoft-com:office:office";

    // ---- wired-in pre-flight behavior (DocxDiff.Compare with the compatibility settings) -------

    private static WmlDocument Math() =>
        new("m.docx", Docx("<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>"));
    private static WmlDocument Plain(string text) =>
        new("p.docx", Docx($"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"));

    [Fact]
    public void Compare_DefaultSettings_DoesNotScanOrThrow()
    {
        // default settings: the pre-flight never runs — an under-tested input does NOT throw.
        var ex = Record.Exception(() => DocxDiff.Compare(Math(), Plain("plain text")));
        Assert.Null(ex);
    }

    [Fact]
    public void Compare_OnCompatibilityWarning_FiresWithReport()
    {
        DocxDiffCompatibilityReport? captured = null;
        DocxDiff.Compare(Math(), Plain("plain"),
            new DocxDiffSettings { OnCompatibilityWarning = r => captured = r });
        Assert.NotNull(captured);
        Assert.Contains(captured!.Warnings, w => w.Feature.Id == "math");
    }

    [Fact]
    public void Compare_ThrowOnCompatibilityWarning_Throws_WithReport()
    {
        var ex = Assert.Throws<DocxDiffCompatibilityException>(() =>
            DocxDiff.Compare(Math(), Plain("plain"),
                new DocxDiffSettings { ThrowOnCompatibilityWarning = true }));
        Assert.Contains(ex.Report.Warnings, w => w.Feature.Id == "math");
    }

    [Fact]
    public void Compare_CleanInputs_DoNotFireOrThrow_EvenWhenEngaged()
    {
        bool fired = false;
        var ex = Record.Exception(() => DocxDiff.Compare(Plain("alpha"), Plain("alpha beta"),
            new DocxDiffSettings { OnCompatibilityWarning = _ => fired = true, ThrowOnCompatibilityWarning = true }));
        Assert.Null(ex);
        Assert.False(fired);
    }

    [Fact]
    public void Compare_FeatureOnlyInRight_IsStillCaught()
    {
        var right = new WmlDocument("r.docx", Docx(
            "<w:p><w:r><w:t>plain</w:t></w:r></w:p>" +
            "<w:p><w:sdt><w:sdtContent><w:r><w:t>cc</w:t></w:r></w:sdtContent></w:sdt></w:p>"));
        DocxDiffCompatibilityReport? captured = null;
        DocxDiff.Compare(Plain("plain"), right,
            new DocxDiffSettings { OnCompatibilityWarning = r => captured = r });
        Assert.NotNull(captured);
        Assert.Contains(captured!.Warnings, w => w.Feature.Id == "contentControls");
    }

    /// <summary>A minimal valid DOCX whose body is <paramref name="bodyInner"/>. The document element declares the
    /// w/m/v/o namespaces so any construct in bodyInner parses under the right namespace. extraParts can add a
    /// footnotes/header/etc. part for multi-part tests.</summary>
    private static byte[] Docx(string bodyInner, Action<MainDocumentPart>? extraParts = null)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            extraParts?.Invoke(main);
            using var s = main.GetStream(FileMode.Create);
            using var wr = new StreamWriter(s);
            wr.Write($"<w:document xmlns:w=\"{W}\" xmlns:m=\"{M}\" xmlns:v=\"{V}\" xmlns:o=\"{O}\"><w:body>" +
                     $"{bodyInner}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return ms.ToArray();
    }

    [Fact]
    public void Catalog_HasUniqueIds_AndIsNonEmpty()
    {
        var ids = DocxDiffCompatibility.Catalog.Select(f => f.Id).ToList();
        Assert.NotEmpty(ids);
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Fact]
    public void Inspect_DetectsMath_AsWarning()
    {
        var bytes = Docx("<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>");
        var report = DocxDiffCompatibility.Inspect(bytes);
        var math = report.Warnings.SingleOrDefault(w => w.Feature.Id == "math");
        Assert.NotNull(math);
        Assert.True(math!.Occurrences >= 1);
    }

    [Fact]
    public void Inspect_CleanDocument_HasNoWarnings()
    {
        var bytes = Docx("<w:p><w:r><w:t>Plain paragraph, nothing exotic.</w:t></w:r></w:p>");
        var report = DocxDiffCompatibility.Inspect(bytes);
        Assert.False(report.HasWarnings);
        Assert.Empty(report.Warnings);
    }

    [Fact]
    public void Inspect_CoveredFeature_GoesToCoveredPresent_NotWarnings()
    {
        // a bookmark + a REF field => the Covered "bookmarksCrossRefs" feature
        var bytes = Docx(
            "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"_Ref_T\"/><w:r><w:t>T</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>" +
            "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:r><w:instrText xml:space=\"preserve\"> REF _Ref_T \\h </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
        var report = DocxDiffCompatibility.Inspect(bytes);
        Assert.DoesNotContain(report.Warnings, w => w.Feature.Id == "bookmarksCrossRefs");
        Assert.Contains(report.CoveredPresent, f => f.Id == "bookmarksCrossRefs");
    }

    [Theory]
    [InlineData("contentControls", "<w:p><w:sdt><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt></w:p>")]
    [InlineData("drawingml", "<w:p><w:r><w:drawing/></w:r></w:p>")]
    [InlineData("textboxes", "<w:p><w:r><w:pict><v:textbox/></w:pict></w:r></w:p>")]
    [InlineData("rtlComplexScript", "<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>")]
    [InlineData("oleEmbeddedObjects", "<w:p><w:r><w:object><o:OLEObject/></w:object></w:r></w:p>")]
    [InlineData("comments", "<w:p><w:r><w:t>x</w:t></w:r><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>")]
    [InlineData("complexFields", "<w:p><w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h </w:instrText></w:r></w:p>")]
    [InlineData("revisionsInInput", "<w:p><w:ins w:id=\"1\" w:author=\"a\"><w:r><w:t>x</w:t></w:r></w:ins></w:p>")]
    [InlineData("footnotesEndnotes", "<w:p><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>")]
    public void Inspect_DetectsFeature_PresentInBody(string featureId, string bodyInner)
    {
        var report = DocxDiffCompatibility.Inspect(Docx(bodyInner));
        var feature = DocxDiffCompatibility.Catalog.Single(f => f.Id == featureId);
        if (feature.Coverage == DocxDiffCoverage.Covered)
            Assert.Contains(report.CoveredPresent, f => f.Id == featureId);
        else
            Assert.Contains(report.Warnings, w => w.Feature.Id == featureId && w.Occurrences >= 1);
    }

    [Fact]
    public void Inspect_RefFieldAlone_IsNotComplexField()
    {
        // a bare REF field (no bookmark) must NOT be flagged as a complex field
        var report = DocxDiffCompatibility.Inspect(Docx(
            "<w:p><w:r><w:instrText xml:space=\"preserve\"> REF _Ref1 \\h </w:instrText></w:r></w:p>"));
        Assert.DoesNotContain(report.Warnings, w => w.Feature.Id == "complexFields");
    }

    [Fact]
    public void Inspect_DetectsFeature_InFootnotePart()
    {
        // math living ONLY in a footnote is still detected (multi-part scan)
        var bytes = Docx(
            "<w:p><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>",
            main =>
            {
                var fn = main.AddNewPart<FootnotesPart>();
                using var s = fn.GetStream(FileMode.Create);
                using var wr = new StreamWriter(s);
                wr.Write($"<w:footnotes xmlns:w=\"{W}\" xmlns:m=\"{M}\">" +
                         "<w:footnote w:id=\"1\"><w:p><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></w:p></w:footnote>" +
                         "</w:footnotes>");
            });
        var report = DocxDiffCompatibility.Inspect(bytes);
        Assert.Contains(report.Warnings, w => w.Feature.Id == "math");
    }

    [Fact]
    public void DocxDiff_InspectCompatibility_Shim_MatchesDirectInspect()
    {
        var bytes = Docx("<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>");
        var viaShim = DocxDiff.InspectCompatibility(new WmlDocument("d.docx", bytes));
        Assert.Contains(viaShim.Warnings, w => w.Feature.Id == "math");
    }

    [Fact]
    public void Summarize_NoWarnings_IsClear()
    {
        var report = DocxDiffCompatibility.Inspect(Docx("<w:p><w:r><w:t>plain</w:t></w:r></w:p>"));
        Assert.False(report.HasWarnings);
        Assert.Contains("no under-tested features", report.Summarize());
    }

    [Fact]
    public void Summarize_ListsEveryWarning()
    {
        var report = DocxDiffCompatibility.Inspect(Docx(
            "<w:p><w:sdt><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt></w:p>" +
            "<w:p><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></w:p>"));
        var summary = report.Summarize();
        Assert.Contains("Content controls", summary);
        Assert.Contains("Office Math", summary);
        // one header line + one line per warning
        Assert.Equal(report.Warnings.Count, summary.Split('\n').Count(l => l.TrimStart().StartsWith("[")));
    }
}
