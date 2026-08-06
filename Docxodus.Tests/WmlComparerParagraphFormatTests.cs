#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for <see cref="WmlComparerSettings.DetectParagraphFormatChanges"/> — paragraph-level
/// format tracking (alignment, indent, spacing, style, list membership) as native
/// <c>w:pPrChange</c> markup.
/// </summary>
public class WmlComparerParagraphFormatTests
{
    private static WmlDocument Doc(params XElement[] bodyChildren)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.PutXDocument(new XDocument(
                new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XElement(W.body, bodyChildren))));

            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(
                new XDocument(
                    new XElement(W.styles,
                        new XAttribute(XNamespace.Xmlns + "w", W.w),
                        Style("Heading1", "heading 1"),
                        Style("Quote", "Quote"))));

            main.AddNewPart<NumberingDefinitionsPart>().PutXDocument(
                new XDocument(
                    new XElement(W.numbering,
                        new XAttribute(XNamespace.Xmlns + "w", W.w),
                        AbstractNum(1, "bullet"),
                        AbstractNum(2, "decimal"),
                        Num(1, 1),
                        Num(2, 2))));

            main.AddNewPart<DocumentSettingsPart>().PutXDocument(
                new XDocument(new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w))));
        }
        return new WmlDocument("test.docx", ms.ToArray());
    }

    private static XElement Style(string id, string name) =>
        new(W.style,
            new XAttribute(W.type, "paragraph"),
            new XAttribute(W.styleId, id),
            new XElement(W.name, new XAttribute(W.val, name)));

    private static XElement AbstractNum(int id, string format) =>
        new(W.abstractNum,
            new XAttribute(W.abstractNumId, id),
            Enumerable.Range(0, 3).Select(lvl => new XElement(W.lvl,
                new XAttribute(W.ilvl, lvl),
                new XElement(W.start, new XAttribute(W.val, "1")),
                new XElement(W.numFmt, new XAttribute(W.val, format)),
                new XElement(W.lvlText, new XAttribute(W.val, format == "bullet" ? "·" : "%1.")),
                new XElement(W.lvlJc, new XAttribute(W.val, "left")))));

    private static XElement Num(int numId, int abstractNumId) =>
        new(W.num,
            new XAttribute(W.numId, numId),
            new XElement(W.abstractNumId, new XAttribute(W.val, abstractNumId)));

    private static XElement Para(string text, XElement? pPr = null) =>
        pPr == null
            ? new XElement(W.p, new XElement(W.r, new XElement(W.t, text)))
            : new XElement(W.p, pPr, new XElement(W.r, new XElement(W.t, text)));

    private static XElement PPr(params XElement[] children) => new(W.pPr, children);

    private static XElement Jc(string val) => new(W.jc, new XAttribute(W.val, val));

    private static XElement PStyle(string id) => new(W.pStyle, new XAttribute(W.val, id));

    private static XElement NumPr(int numId, int ilvl) =>
        new(W.numPr,
            new XElement(W.ilvl, new XAttribute(W.val, ilvl)),
            new XElement(W.numId, new XAttribute(W.val, numId)));

    private static XElement SectPr() =>
        new(W.sectPr,
            new XElement(W.pgSz, new XAttribute(W._w, "12240"), new XAttribute(W.h, "15840")));

    private static WmlComparerSettings Settings(bool detectParagraphFormat) => new()
    {
        DetectParagraphFormatChanges = detectParagraphFormat,
        DateTimeForRevisions = "2000-01-01T00:00:00Z",
    };

    private static XElement Body(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    /// <summary>The result of comparing left against right with paragraph format tracking on.</summary>
    private static WmlDocument CompareOn(WmlDocument left, WmlDocument right)
    {
        var result = WmlComparer.Compare(left, right, Settings(true));
        AssertSchemaValid(result);
        return result;
    }

    private static void AssertSchemaValid(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);

        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(wDoc)
            .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema &&
                        !OxPt.WcTests.ExpectedErrors.Contains(e.Description))
            .Select(e => e.Description)
            .ToList();
        Assert.Empty(errors);
    }

    private static List<XElement> PPrChanges(WmlDocument doc) =>
        Body(doc).Descendants(W.pPrChange).ToList();

    /// <summary>The properties a pPrChange archives, as local names in document order.</summary>
    private static List<string> ArchivedProps(XElement pPrChange) =>
        pPrChange.Element(W.pPr)!.Elements().Select(e => e.Name.LocalName).ToList();

    /// <summary>The live properties of the paragraph owning a pPrChange, excluding the marker.</summary>
    private static List<string> LiveProps(XElement pPrChange) =>
        pPrChange.Parent!.Elements()
            .Where(e => e.Name != W.pPrChange)
            .Select(e => e.Name.LocalName)
            .ToList();

    // ---------------------------------------------------------------------------------------
    // Detection
    // ---------------------------------------------------------------------------------------

    /// <summary>PF001 — alignment change is tracked, and does not become a delete/insert.</summary>
    [Fact]
    public void PF001_AlignmentChange_EmitsPPrChange()
    {
        var result = CompareOn(
            Doc(Para("Hello")),
            Doc(Para("Hello", PPr(Jc("center")))));

        var change = Assert.Single(PPrChanges(result));
        Assert.Equal(new[] { "jc" }, LiveProps(change));
        Assert.Empty(ArchivedProps(change));

        var body = Body(result);
        Assert.Empty(body.Descendants(W.ins));
        Assert.Empty(body.Descendants(W.del));
    }

    /// <summary>
    /// PF002 — the guard that matters most: off by default, so an existing caller who upgrades
    /// sees byte-for-byte what they saw before.
    /// </summary>
    [Fact]
    public void PF002_OffByDefault_EmitsNothing()
    {
        Assert.False(new WmlComparerSettings().DetectParagraphFormatChanges);

        var left = Doc(Para("Hello"));
        var right = Doc(Para("Hello", PPr(Jc("center"))));

        var off = WmlComparer.Compare(left, right, Settings(false));
        Assert.Empty(PPrChanges(off));
        Assert.Equal(0, WmlComparer.GetRevisions(off, Settings(false)).Count);

        // Not vacuous: the same pair does produce a change with the option on.
        Assert.Single(PPrChanges(CompareOn(left, right)));
    }

    /// <summary>PF003 — a paragraph style change.</summary>
    [Fact]
    public void PF003_StyleChange_EmitsPPrChange()
    {
        var result = CompareOn(
            Doc(Para("Hello", PPr(PStyle("Quote")))),
            Doc(Para("Hello", PPr(PStyle("Heading1")))));

        var change = Assert.Single(PPrChanges(result));
        Assert.Equal("Heading1", (string)change.Parent!.Element(W.pStyle)!.Attribute(W.val)!);
        Assert.Equal("Quote", (string)change.Element(W.pPr)!.Element(W.pStyle)!.Attribute(W.val)!);
    }

    /// <summary>PF004 — indent and spacing, the other two common paragraph properties.</summary>
    [Fact]
    public void PF004_IndentAndSpacingChange_EmitsPPrChange()
    {
        var result = CompareOn(
            Doc(Para("Hello")),
            Doc(Para("Hello", PPr(
                new XElement(W.ind, new XAttribute(W.left, "720")),
                new XElement(W.spacing, new XAttribute(W.before, "240"))))));

        // CT_PPrBase order, which the comparer restores: w:spacing precedes w:ind.
        var change = Assert.Single(PPrChanges(result));
        Assert.Equal(new[] { "spacing", "ind" }, LiveProps(change));
    }

    /// <summary>
    /// PF004b — a source paragraph whose properties are stored out of schema order. The archived
    /// copy must come out ordered, or Word rejects the w:pPrChange.
    /// </summary>
    [Fact]
    public void PF004b_OutOfOrderSourceProperties_AreArchivedInSchemaOrder()
    {
        var result = CompareOn(
            Doc(Para("Hello", PPr(
                new XElement(W.ind, new XAttribute(W.left, "720")),
                new XElement(W.spacing, new XAttribute(W.before, "240")),
                PStyle("Quote")))),
            Doc(Para("Hello", PPr(Jc("center")))));

        var change = Assert.Single(PPrChanges(result));
        Assert.Equal(new[] { "pStyle", "spacing", "ind" }, ArchivedProps(change));
    }

    // ---------------------------------------------------------------------------------------
    // Lists — w:numPr is an ordinary pPr child, so list changes ride the same path
    // ---------------------------------------------------------------------------------------

    /// <summary>PF005 — a paragraph made into a list item.</summary>
    [Fact]
    public void PF005_ListMembershipAdded_EmitsPPrChange()
    {
        var result = CompareOn(
            Doc(Para("Hello")),
            Doc(Para("Hello", PPr(NumPr(1, 0)))));

        var change = Assert.Single(PPrChanges(result));
        Assert.Equal(new[] { "numPr" }, LiveProps(change));
        Assert.Empty(ArchivedProps(change));
    }

    /// <summary>PF006 — a list item demoted to plain text: the archive holds the old numbering.</summary>
    [Fact]
    public void PF006_ListMembershipRemoved_ArchivesOldNumbering()
    {
        var result = CompareOn(
            Doc(Para("Hello", PPr(NumPr(1, 0)))),
            Doc(Para("Hello")));

        var change = Assert.Single(PPrChanges(result));
        Assert.Empty(LiveProps(change));
        Assert.Equal(new[] { "numPr" }, ArchivedProps(change));
        Assert.Equal("1", (string)change.Descendants(W.numId).Single().Attribute(W.val)!);
    }

    /// <summary>PF007 — bullet list to numbered list, and an indent-level change.</summary>
    [Theory]
    [InlineData(1, 0, 2, 0)] // bullet -> decimal
    [InlineData(1, 0, 1, 1)] // level 0 -> level 1
    public void PF007_ListFormatAndLevelChanges_AreTracked(int leftNum, int leftLvl, int rightNum, int rightLvl)
    {
        var result = CompareOn(
            Doc(Para("Hello", PPr(NumPr(leftNum, leftLvl)))),
            Doc(Para("Hello", PPr(NumPr(rightNum, rightLvl)))));

        var change = Assert.Single(PPrChanges(result));
        Assert.Equal(leftNum.ToString(), (string)change.Element(W.pPr)!.Descendants(W.numId).Single().Attribute(W.val)!);
        Assert.Equal(leftLvl.ToString(), (string)change.Element(W.pPr)!.Descendants(W.ilvl).Single().Attribute(W.val)!);
    }

    // ---------------------------------------------------------------------------------------
    // False-positive guards — these carry more weight than the positive cases
    // ---------------------------------------------------------------------------------------

    /// <summary>PF008 — identical documents produce nothing.</summary>
    [Fact]
    public void PF008_IdenticalDocuments_ProduceNoChange()
    {
        var result = CompareOn(
            Doc(Para("Hello", PPr(Jc("center"), NumPr(1, 0)))),
            Doc(Para("Hello", PPr(Jc("center"), NumPr(1, 0)))));

        Assert.Empty(PPrChanges(result));
        Assert.Equal(0, WmlComparer.GetRevisions(result, Settings(true)).Count);
    }

    /// <summary>
    /// PF008b — the same properties written in a different source order are not a change.
    /// <see cref="XNode.DeepEquals"/> is order-sensitive, so the reduction has to sort.
    /// </summary>
    [Fact]
    public void PF008b_SamePropertiesInADifferentOrder_ProduceNoChange()
    {
        XElement Ind() => new(W.ind, new XAttribute(W.left, "720"));
        XElement Spacing() => new(W.spacing, new XAttribute(W.before, "240"));

        var result = CompareOn(
            Doc(Para("Hello", PPr(Ind(), Spacing()))),
            Doc(Para("Hello", PPr(Spacing(), Ind()))));

        Assert.Empty(PPrChanges(result));
    }

    /// <summary>
    /// PF008c — the same applies inside a property: a w:numPr holding w:numId before w:ilvl is the
    /// same numbering, so the sort reaches every level.
    /// <para>The fixture is deliberately malformed — CT_NumPr is a sequence, so only a non-Word
    /// generator writes it this way, and unlike a top-level w:pPr the writer's element ordering does
    /// not reach inside w:numPr, leaving the output as invalid as the input. Hence no schema
    /// assertion here: what is pinned is that malformed input yields no phantom revision.</para>
    /// </summary>
    [Fact]
    public void PF008c_SameNumberingInADifferentChildOrder_ProducesNoChange()
    {
        XElement Ilvl() => new(W.ilvl, new XAttribute(W.val, "0"));
        XElement NumId() => new(W.numId, new XAttribute(W.val, "1"));

        var result = WmlComparer.Compare(
            Doc(Para("Hello", PPr(new XElement(W.numPr, Ilvl(), NumId())))),
            Doc(Para("Hello", PPr(new XElement(W.numPr, NumId(), Ilvl())))),
            Settings(true));

        Assert.Empty(PPrChanges(result));
    }

    /// <summary>
    /// PF009 — revision-save ids differ between two saves of the same paragraph. They are not a
    /// formatting change and must not be reported as one.
    /// </summary>
    [Fact]
    public void PF009_RsidOnlyDifference_ProducesNoChange()
    {
        var left = Doc(new XElement(W.p,
            new XAttribute(W.rsidR, "00AA00AA"),
            new XElement(W.pPr, new XAttribute(W.rsidRDefault, "00AA00AA"), Jc("center")),
            new XElement(W.r, new XElement(W.t, "Hello"))));
        var right = Doc(new XElement(W.p,
            new XAttribute(W.rsidR, "00BB00BB"),
            new XElement(W.pPr, new XAttribute(W.rsidRDefault, "00BB00BB"), Jc("center")),
            new XElement(W.r, new XElement(W.t, "Hello"))));

        Assert.Empty(PPrChanges(CompareOn(left, right)));
    }

    /// <summary>
    /// PF010 — a changed section is a section change (w:sectPrChange), explicitly out of scope. It
    /// must not masquerade as a paragraph format change.
    /// <para>Note this passes for a second reason as well: WmlComparer discards a mid-document
    /// inline w:sectPr altogether, with or without this option. That is pre-existing behaviour, not
    /// something the option introduces — see PF012, which tests the exclusion directly because this
    /// route cannot reach it.</para>
    /// </summary>
    [Fact]
    public void PF010_SectionOnlyDifference_ProducesNoParagraphChange()
    {
        var left = Doc(
            Para("First", PPr(SectPr())), Para("Second"),
            new XElement(W.sectPr, new XElement(W.pgSz,
                new XAttribute(W._w, "12240"), new XAttribute(W.h, "15840"))));
        var right = Doc(
            Para("First", PPr(new XElement(W.sectPr, new XElement(W.pgSz,
                new XAttribute(W._w, "15840"), new XAttribute(W.h, "12240"))))),
            Para("Second"),
            new XElement(W.sectPr, new XElement(W.pgSz,
                new XAttribute(W._w, "12240"), new XAttribute(W.h, "15840"))));

        Assert.Empty(PPrChanges(CompareOn(left, right)));
    }

    /// <summary>
    /// PF011 — the paragraph mark's own run properties are out of scope (Word tracks those as
    /// w:pPr/w:rPr/w:rPrChange). Documented, not accidental.
    /// </summary>
    [Fact]
    public void PF011_ParagraphMarkRunPropertiesChange_IsNotTracked()
    {
        var left = Doc(Para("Hello", PPr(new XElement(W.rPr, new XElement(W.b)))));
        var right = Doc(Para("Hello", PPr(new XElement(W.rPr, new XElement(W.i)))));

        Assert.Empty(PPrChanges(CompareOn(left, right)));
    }

    // ---------------------------------------------------------------------------------------
    // The sectPr trap: w:pPrChange stores a CT_PPrBase, which excludes w:sectPr and w:rPr
    // ---------------------------------------------------------------------------------------

    /// <summary>
    /// PF012 — what a w:pPrChange may archive, tested on the reduction directly.
    /// <para>A w:pPrChange stores a CT_PPrBase, which excludes w:sectPr, w:rPr and a nested
    /// w:pPrChange; including any of them produces a file Word refuses. This is asserted against
    /// the reduction rather than through Compare because the comparer discards an inline w:sectPr
    /// before the comparison sees it (PF010), so no document pair can drive this path — the guard
    /// exists so that stays true if the comparer ever starts preserving inline sections.</para>
    /// </summary>
    [Fact]
    public void PF012_ArchivedProperties_ExcludeWhatCtPPrBaseForbids()
    {
        var pPr = PPr(
            PStyle("Quote"),
            new XElement(W.rPr, new XElement(W.b)),
            SectPr(),
            new XElement(W.pPrChange,
                new XAttribute(W.id, "9"),
                new XAttribute(W.author, "Someone"),
                new XAttribute(W.date, "1999-01-01T00:00:00Z"),
                PPr(Jc("left"))));

        var reduced = WmlComparer.ReduceToParagraphPropertiesChange(pPr);

        Assert.Equal(new[] { "pStyle" }, reduced.Elements().Select(e => e.Name.LocalName));
    }

    /// <summary>
    /// PF012b — the reduction ignores revision-save ids and internal bookkeeping at every level, so
    /// they neither cause a false positive nor leak into the archive.
    /// </summary>
    [Fact]
    public void PF012b_ArchivedProperties_StripBookkeepingAtEveryLevel()
    {
        var pPr = PPr(NumPr(1, 0));
        pPr.SetAttributeValue(W.rsidRDefault, "00AA00AA");
        pPr.Element(W.numPr)!.SetAttributeValue(PtOpenXml.Unid, "deadbeef");
        pPr.Element(W.numPr)!.Element(W.numId)!.SetAttributeValue(PtOpenXml.Unid, "cafebabe");

        var reduced = WmlComparer.ReduceToParagraphPropertiesChange(pPr);

        Assert.Empty(reduced.DescendantsAndSelf().Attributes()
            .Where(a => a.Name.Namespace == PtOpenXml.pt ||
                        a.Name.LocalName.StartsWith("rsid", System.StringComparison.OrdinalIgnoreCase)));
        Assert.Equal("1", (string)reduced.Descendants(W.numId).Single().Attribute(W.val)!);
    }

    // ---------------------------------------------------------------------------------------
    // Round trip
    // ---------------------------------------------------------------------------------------

    /// <summary>
    /// PF013 — the contract: accepting reproduces the revised document's properties, rejecting
    /// restores the original's. Rejecting is what the feature buys — today the original is lost.
    /// </summary>
    [Fact]
    public void PF013_AcceptAndReject_RoundTripParagraphProperties()
    {
        var result = CompareOn(
            Doc(Para("Hello", PPr(Jc("both")))),
            Doc(Para("Hello", PPr(Jc("center"), NumPr(2, 0)))));

        var accepted = Body(RevisionProcessor.AcceptRevisions(result)).Descendants(W.p).First();
        Assert.Equal("center", (string)accepted.Element(W.pPr)!.Element(W.jc)!.Attribute(W.val)!);
        Assert.NotNull(accepted.Element(W.pPr)!.Element(W.numPr));
        Assert.Empty(accepted.Descendants(W.pPrChange));

        var rejected = Body(RevisionProcessor.RejectRevisions(result)).Descendants(W.p).First();
        Assert.Equal("both", (string)rejected.Element(W.pPr)!.Element(W.jc)!.Attribute(W.val)!);
        Assert.Null(rejected.Element(W.pPr)!.Element(W.numPr));
        Assert.Empty(rejected.Descendants(W.pPrChange));
    }

    /// <summary>
    /// PF014 — a format change on the last paragraph of a document with a trailing section. The
    /// section survives both the comparison and the reject.
    /// </summary>
    [Fact]
    public void PF014_TrailingSection_SurvivesCompareAndReject()
    {
        var trailing = new XElement(W.sectPr,
            new XElement(W.pgSz, new XAttribute(W._w, "12240"), new XAttribute(W.h, "15840")));

        var result = CompareOn(
            Doc(Para("Only"), new XElement(trailing)),
            Doc(Para("Only", PPr(Jc("center"))), new XElement(trailing)));

        Assert.Single(PPrChanges(result));
        Assert.NotNull(Body(result).Element(W.sectPr));

        var rejected = Body(RevisionProcessor.RejectRevisions(result));
        Assert.Null(rejected.Elements(W.p).First().Element(W.pPr)?.Element(W.jc));
        Assert.NotNull(rejected.Element(W.sectPr));
    }

    // ---------------------------------------------------------------------------------------
    // GetRevisions
    // ---------------------------------------------------------------------------------------

    /// <summary>PF015 — a paragraph format change is reported, with the paragraph's text.</summary>
    [Fact]
    public void PF015_GetRevisions_ReportsParagraphFormatChange()
    {
        var settings = Settings(true);
        var result = CompareOn(
            Doc(Para("Alpha"), Para("Beta")),
            Doc(Para("Alpha"), Para("Beta", PPr(Jc("center")))));

        var revision = Assert.Single(WmlComparer.GetRevisions(result, settings));
        Assert.Equal(WmlComparer.WmlComparerRevisionType.FormatChanged, revision.RevisionType);
        Assert.Equal("Beta", revision.Text);
        Assert.Equal(settings.AuthorForRevisions, revision.Author);
        Assert.Equal(new[] { "alignment" }, revision.FormatChange!.ChangedPropertyNames);
    }

    /// <summary>
    /// PF015b — only the properties that actually differ are reported. A property with no w:val is
    /// compared by its serialized form, so an unrelated w:ind must not read as changed just because
    /// the two copies carry different internal ids.
    /// </summary>
    [Fact]
    public void PF015b_GetRevisions_ReportsOnlyThePropertiesThatDiffer()
    {
        XElement Ind() => new(W.ind, new XAttribute(W.left, "720"));

        var settings = Settings(true);
        var result = CompareOn(
            Doc(Para("Hello", PPr(Ind()))),
            Doc(Para("Hello", PPr(Ind(), Jc("center")))));

        var revision = Assert.Single(WmlComparer.GetRevisions(result, settings));
        Assert.Equal(new[] { "alignment" }, revision.FormatChange!.ChangedPropertyNames);
    }

    /// <summary>
    /// PF015c — reporting is not gated by the setting. `GetRevisions` reads the markup, so a
    /// document that already carries a w:pPrChange — written by Word, not by us — is reported even
    /// with paragraph detection off, exactly as w:rPrChange already was. That keeps a document
    /// produced WITH the option from going silent when read back with default settings.
    /// </summary>
    [Fact]
    public void PF015c_GetRevisions_ReportsExistingParagraphFormatMarkup_RegardlessOfTheSetting()
    {
        var doc = Doc(new XElement(W.p,
            PPr(Jc("center"),
                new XElement(W.pPrChange,
                    new XAttribute(W.id, "1"),
                    new XAttribute(W.author, "Word"),
                    new XAttribute(W.date, "2000-01-01T00:00:00Z"),
                    PPr())),
            new XElement(W.r, new XElement(W.t, "Hello"))));

        var revision = Assert.Single(WmlComparer.GetRevisions(doc, Settings(false)));
        Assert.Equal(WmlComparer.WmlComparerRevisionType.FormatChanged, revision.RevisionType);
        Assert.Equal("Word", revision.Author);
    }

    /// <summary>
    /// PF023 — the archived properties keep the w: prefix. Serializing them for their trip through
    /// a bookkeeping attribute used to emit a default-namespace element that re-declared the
    /// namespace on every child. Covers the run-level w:rPrChange too, which shares the round trip
    /// and is on by default.
    /// </summary>
    [Fact]
    public void PF023_ArchivedPropertiesKeepThePrefix_ForBothScopes()
    {
        var paragraph = CompareOn(
            Doc(Para("Hello", PPr(PStyle("Quote")))),
            Doc(Para("Hello", PPr(PStyle("Heading1")))));

        var runs = WmlComparer.Compare(
            Doc(new XElement(W.p,
                new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), new XElement(W.t, "Hello")))),
            Doc(new XElement(W.p,
                new XElement(W.r, new XElement(W.rPr, new XElement(W.i)), new XElement(W.t, "Hello")))),
            Settings(true));
        AssertSchemaValid(runs);

        foreach (var archived in PPrChanges(paragraph).Select(c => c.Element(W.pPr)!)
                     .Concat(Body(runs).Descendants(W.rPrChange).Select(c => c.Element(W.rPr)!)))
        {
            Assert.NotEmpty(archived.Elements());
            Assert.DoesNotContain(archived.DescendantsAndSelf().Attributes(), a => a.IsNamespaceDeclaration);
        }
    }

    // ---------------------------------------------------------------------------------------
    // Interaction with the existing run-level detection
    // ---------------------------------------------------------------------------------------

    /// <summary>
    /// PF016 — the two detectors are independent: turning paragraph tracking on does not disturb
    /// run tracking, and both can fire on one paragraph.
    /// </summary>
    [Fact]
    public void PF016_RunAndParagraphFormatChanges_Coexist()
    {
        var result = CompareOn(
            Doc(Para("Hello")),
            Doc(new XElement(W.p,
                PPr(Jc("center")),
                new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), new XElement(W.t, "Hello")))));

        var body = Body(result);
        Assert.Single(body.Descendants(W.pPrChange));
        Assert.Single(body.Descendants(W.rPrChange));
        Assert.Equal(2, WmlComparer.GetRevisions(result, Settings(true)).Count);
    }

    /// <summary>
    /// PF017 — run format detection still works with paragraph detection off, and vice versa. The
    /// two settings gate only their own pass.
    /// </summary>
    [Fact]
    public void PF017_SettingsGateOnlyTheirOwnPass()
    {
        var left = Doc(Para("Hello"));
        var right = Doc(new XElement(W.p,
            PPr(Jc("center")),
            new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), new XElement(W.t, "Hello"))));

        var runOnly = WmlComparer.Compare(left, right,
            new WmlComparerSettings { DetectFormatChanges = true, DetectParagraphFormatChanges = false });
        Assert.Empty(Body(runOnly).Descendants(W.pPrChange));
        Assert.Single(Body(runOnly).Descendants(W.rPrChange));

        var paragraphOnly = WmlComparer.Compare(left, right,
            new WmlComparerSettings { DetectFormatChanges = false, DetectParagraphFormatChanges = true });
        Assert.Single(Body(paragraphOnly).Descendants(W.pPrChange));
        Assert.Empty(Body(paragraphOnly).Descendants(W.rPrChange));
    }

    /// <summary>
    /// PF018 — a paragraph that is both reformatted and edited. The text edit is a real
    /// insert/delete; the format change rides alongside it.
    /// </summary>
    [Fact]
    public void PF018_TextEditAndFormatChange_OnTheSameParagraph()
    {
        var result = CompareOn(
            Doc(Para("Hello world")),
            Doc(Para("Hello there", PPr(Jc("center")))));

        var body = Body(result);

        // One paragraph, not two: marking the paragraph mark FormatChanged must not stop it taking
        // the left document's ancestor unids, or its runs and its mark regroup separately.
        Assert.Single(body.Elements(W.p));
        Assert.Single(body.Descendants(W.pPrChange));
        Assert.NotEmpty(body.Descendants(W.ins));
        Assert.NotEmpty(body.Descendants(W.del));
        Assert.Equal("Hello there", body.Descendants(W.t).Select(t => t.Value).StringConcatenate());
    }

    /// <summary>
    /// PF019 — only the reformatted paragraph is touched; its neighbours stay clean.
    /// </summary>
    [Fact]
    public void PF019_OnlyTheChangedParagraphIsMarked()
    {
        var result = CompareOn(
            Doc(Para("Alpha"), Para("Beta"), Para("Gamma")),
            Doc(Para("Alpha"), Para("Beta", PPr(Jc("center"))), Para("Gamma")));

        var paragraphs = Body(result).Elements(W.p).ToList();
        Assert.Equal(3, paragraphs.Count);
        Assert.Empty(paragraphs[0].Descendants(W.pPrChange));
        Assert.Single(paragraphs[1].Descendants(W.pPrChange));
        Assert.Empty(paragraphs[2].Descendants(W.pPrChange));
    }

    /// <summary>
    /// PF020 — the serialized old properties do not ride into the output as an attribute. The live
    /// w:pPr and its children still carry the converter's own pt14 unids, as everywhere else in a
    /// produced document; only the OldPPr hand-off attribute must be gone.
    /// </summary>
    [Fact]
    public void PF020_TheSerializedOldPropertiesDoNotLeakIntoOutput()
    {
        var result = CompareOn(
            Doc(Para("Hello")),
            Doc(Para("Hello", PPr(Jc("center"), NumPr(1, 0)))));

        var change = Assert.Single(PPrChanges(result));
        Assert.DoesNotContain(change.Parent!.Attributes(),
            a => a.Name == PtOpenXml.pt + "OldPPr" || a.Name == PtOpenXml.Status);
        Assert.DoesNotContain(Body(result).DescendantsAndSelf().Attributes(),
            a => a.Name == PtOpenXml.pt + "OldPPr");
    }

    /// <summary>
    /// PF022 — pins a PRE-EXISTING limitation, not a property of this option: WmlComparer detects no
    /// formatting change inside a footnote, at either granularity. A text change in the same note IS
    /// detected, so notes are compared — it is the format passes that do not reach them. The
    /// run-level pass (<see cref="WmlComparerSettings.DetectFormatChanges"/>, on by default) behaves
    /// identically, which is why this is asserted for both.
    /// <para>If a future change makes note-scope format detection work, this test fails and should
    /// be rewritten to assert the markup rather than its absence.</para>
    /// </summary>
    [Fact]
    public void PF022_FormatChangesInsideAFootnote_AreNotDetected_PreExistingLimitation()
    {
        var settings = Settings(true);

        var paragraphChange = WmlComparer.Compare(
            WithFootnote("Body text", "Note text"),
            MutateNoteParagraph(WithFootnote("Body text", "Note text"), p =>
            {
                var pPr = p.Element(W.pPr)!;
                pPr.Add(Jc("center"));
            }),
            settings);

        var runChange = WmlComparer.Compare(
            WithFootnote("Body text", "Note text"),
            MutateNoteParagraph(WithFootnote("Body text", "Note text"), p =>
            {
                foreach (var run in p.Elements(W.r).Where(r => r.Elements(W.t).Any()))
                    run.AddFirst(new XElement(W.rPr, new XElement(W.b)));
            }),
            settings);

        Assert.Empty(NoteMarkup(paragraphChange, W.pPrChange));
        Assert.Empty(NoteMarkup(runChange, W.rPrChange));

        // Not vacuous: the note itself is compared, so the absence above is about formatting only.
        var textChange = WmlComparer.Compare(
            WithFootnote("Body text", "Note text"),
            MutateNoteParagraph(WithFootnote("Body text", "Note text"), p =>
            {
                foreach (var t in p.Descendants(W.t))
                    t.Value = "Different";
            }),
            settings);

        Assert.NotEmpty(NoteMarkup(textChange, W.ins));
        Assert.NotEmpty(NoteMarkup(textChange, W.del));
    }

    private static List<XElement> NoteMarkup(WmlDocument doc, XName name)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.FootnotesPart!.GetXDocument().Descendants(name).ToList();
    }

    /// <summary>A document with one real footnote, built through the session API's scaffold.</summary>
    private static WmlDocument WithFootnote(string bodyText, string noteText)
    {
        var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var body = session.ListBlocks().Body.First(b => b.Kind == "p");

        session.ReplaceText(body.Id, bodyText);
        var inserted = session.InsertFootnote(body.Id, bodyText.Length, noteText);
        Assert.True(inserted.Success, inserted.Error?.Message);

        return new WmlDocument("note.docx", session.Save());
    }

    /// <summary>Edits the real footnote's paragraph, changing nothing else in the document.</summary>
    private static WmlDocument MutateNoteParagraph(WmlDocument doc, System.Action<XElement> mutate)
    {
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        using (var wDoc = WordprocessingDocument.Open(ms, true))
        {
            var notePara = wDoc.MainDocumentPart!.FootnotesPart!.GetXDocument()
                .Root!.Elements(W.footnote)
                .Where(f => (string)f.Attribute(W.type) == null)   // skip the reserved separators
                .Descendants(W.p)
                .First();

            mutate(notePara);
            wDoc.MainDocumentPart.FootnotesPart.PutXDocument();
        }
        return new WmlDocument("note.docx", ms.ToArray());
    }

    /// <summary>PF021 — a change inside a table cell is tracked like any other paragraph.</summary>
    [Fact]
    public void PF021_ParagraphInsideTableCell_IsTracked()
    {
        XElement Table(XElement cellPara) =>
            new(W.tbl,
                new XElement(W.tblPr, new XElement(W.tblW, new XAttribute(W._w, "0"), new XAttribute(W.type, "auto"))),
                new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W._w, "3000"))),
                new XElement(W.tr, new XElement(W.tc, cellPara)));

        var result = CompareOn(
            Doc(Table(Para("Cell")), Para("After")),
            Doc(Table(Para("Cell", PPr(Jc("center")))), Para("After")));

        var change = Assert.Single(PPrChanges(result));
        Assert.NotNull(change.Ancestors(W.tc).FirstOrDefault());
    }
}
