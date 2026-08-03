// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace OxPt
{
    /// <summary>
    /// Tests for list numbering in tracked-changes documents (the redline render).
    ///
    /// Word ties numbering to the paragraph mark and renumbers a tracked-changes
    /// document as if every revision were already accepted: a paragraph whose pilcrow
    /// is marked deleted (<c>w:pPr/w:rPr/w:del</c> — a fully deleted or moved-away
    /// paragraph) still displays the value the counter holds at its position, but does
    /// not advance the counter, so the next live paragraph shows the same number (the
    /// duplicate numbers Word renders in All Markup view). Inserted paragraphs consume
    /// numbers normally. Before this behavior was implemented, a redline of a numbered
    /// list showed one continuous sequence across deleted paragraphs, so every number
    /// after a deletion disagreed with the final document (see the NVCA marquee
    /// screenshot bug).
    /// </summary>
    public class TrackedChangesNumberingTests
    {
        private const string W_NS = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        private static readonly XNamespace Xh = "http://www.w3.org/1999/xhtml";

        #region Helpers

        private static string NumberedParaXml(string text) =>
            $"<w:p {W_NS}><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

        /// <summary>Fully deleted numbered paragraph: deleted pilcrow + deleted content.</summary>
        private static string DeletedParaXml(string text) =>
            $"<w:p {W_NS}><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>" +
            "<w:rPr><w:del w:id=\"101\" w:author=\"Reviewer\" w:date=\"2024-01-01T00:00:00Z\"/></w:rPr></w:pPr>" +
            "<w:del w:id=\"102\" w:author=\"Reviewer\" w:date=\"2024-01-01T00:00:00Z\">" +
            $"<w:r><w:delText>{text}</w:delText></w:r></w:del></w:p>";

        /// <summary>Moved-away numbered paragraph: moveFrom pilcrow + moveFrom content.</summary>
        private static string MovedFromParaXml(string text) =>
            $"<w:p {W_NS}><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>" +
            "<w:rPr><w:moveFrom w:id=\"111\" w:author=\"Reviewer\" w:date=\"2024-01-01T00:00:00Z\"/></w:rPr></w:pPr>" +
            "<w:moveFrom w:id=\"112\" w:author=\"Reviewer\" w:date=\"2024-01-01T00:00:00Z\">" +
            $"<w:r><w:t>{text}</w:t></w:r></w:moveFrom></w:p>";

        /// <summary>Wholly inserted numbered paragraph: inserted pilcrow + inserted content.</summary>
        private static string InsertedParaXml(string text) =>
            $"<w:p {W_NS}><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>" +
            "<w:rPr><w:ins w:id=\"121\" w:author=\"Reviewer\" w:date=\"2024-01-01T00:00:00Z\"/></w:rPr></w:pPr>" +
            "<w:ins w:id=\"122\" w:author=\"Reviewer\" w:date=\"2024-01-01T00:00:00Z\">" +
            $"<w:r><w:t>{text}</w:t></w:r></w:ins></w:p>";

        /// <summary>
        /// A pre-existing paragraph whose pilcrow was inserted: pressing Enter inside an
        /// existing paragraph with track changes on leaves the ins-marked mark on the first
        /// half and its (unchanged) content in place — the FOLLOWING paragraph occupies the
        /// newly created list position.
        /// </summary>
        private static string SplitFirstHalfParaXml(string text) =>
            $"<w:p {W_NS}><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>" +
            "<w:rPr><w:ins w:id=\"131\" w:author=\"Reviewer\" w:date=\"2024-01-01T00:00:00Z\"/></w:rPr></w:pPr>" +
            $"<w:r><w:t>{text}</w:t></w:r></w:p>";

        private static WmlDocument CreateNumberedListDocument(params string[] paragraphXml)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                var body = new Body();
                foreach (var xml in paragraphXml)
                    body.Append(new Paragraph(xml));
                mainPart.Document = new Document(body);

                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(
                    new DocDefaults(
                        new RunPropertiesDefault(
                            new RunPropertiesBaseStyle(
                                new RunFonts { Ascii = "Calibri" },
                                new FontSize { Val = "22" }
                            )
                        ),
                        new ParagraphPropertiesDefault()
                    )
                );

                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();

                var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
                var levelElement = new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." },
                    new LevelJustification { Val = LevelJustificationValues.Left },
                    new PreviousParagraphProperties(
                        new Indentation { Left = "720", Hanging = "360" }
                    )
                ) { LevelIndex = 0 };
                var abstractNum = new AbstractNum(levelElement) { AbstractNumberId = 1 };
                var numberingInstance = new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 };
                numberingPart.Numbering = new Numbering(abstractNum, numberingInstance);

                doc.Save();
            }

            stream.Position = 0;
            return new WmlDocument("test.docx", stream.ToArray());
        }

        private static XElement ConvertWithTrackedChanges(WmlDocument wml)
        {
            using var ms = new MemoryStream();
            ms.Write(wml.DocumentByteArray, 0, wml.DocumentByteArray.Length);
            using var wDoc = WordprocessingDocument.Open(ms, true);
            var settings = new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = true,
                CssClassPrefix = "pt-",
                RenderTrackedChanges = true,
                IncludeRevisionMetadata = true,
                ShowDeletedContent = true,
            };
            return WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
        }

        /// <summary>
        /// The rendered number glyphs (e.g. "1.") in document order, each with whether it
        /// sits inside an <c>&lt;ins&gt;</c> or <c>&lt;del&gt;</c> wrapper.
        /// </summary>
        private static List<(string Number, bool InsideIns, bool InsideDel)> ExtractMarkers(XElement html)
        {
            return html
                .Descendants(Xh + "span")
                .Where(s => (string?)s.Attribute("data-list-marker") == "true")
                .Where(s => Regex.IsMatch(s.Value, @"^\d+\.$"))
                .Select(s => (
                    Number: s.Value,
                    InsideIns: s.Ancestors(Xh + "ins").Any(),
                    InsideDel: s.Ancestors(Xh + "del").Any()))
                .ToList();
        }

        #endregion

        [Fact]
        public void TCN001_DeletedParagraph_DoesNotConsumeANumber()
        {
            var wml = CreateNumberedListDocument(
                NumberedParaXml("Alpha"),
                DeletedParaXml("Bravo"),
                NumberedParaXml("Charlie"),
                NumberedParaXml("Delta"));

            var html = ConvertWithTrackedChanges(wml);
            var markers = ExtractMarkers(html);

            // Word renumbers as if the deletion were accepted: the deleted paragraph
            // displays the counter value without advancing it, so "2." appears twice —
            // once struck on the deleted paragraph, once on the next live paragraph.
            Assert.Equal(new[] { "1.", "2.", "2.", "3." }, markers.Select(m => m.Number));
            Assert.True(markers[1].InsideDel, "the deleted paragraph's number renders struck");
            Assert.False(markers[2].InsideDel);
        }

        [Fact]
        public void TCN002_MovedFromParagraph_DoesNotConsumeANumber()
        {
            var wml = CreateNumberedListDocument(
                NumberedParaXml("Alpha"),
                MovedFromParaXml("Bravo"),
                NumberedParaXml("Charlie"));

            var html = ConvertWithTrackedChanges(wml);
            var markers = ExtractMarkers(html);

            Assert.Equal(new[] { "1.", "2.", "2." }, markers.Select(m => m.Number));
        }

        [Fact]
        public void TCN003_InsertedParagraph_ConsumesANumber_WithoutLeakingInsOntoFollower()
        {
            var wml = CreateNumberedListDocument(
                NumberedParaXml("Alpha"),
                InsertedParaXml("Inserted"),
                NumberedParaXml("Bravo"),
                NumberedParaXml("Charlie"));

            var html = ConvertWithTrackedChanges(wml);
            var markers = ExtractMarkers(html);

            // The inserted paragraph is part of the final document, so it takes "2."
            // and pushes the rest down — and only ITS number renders as an insertion.
            // (Before the fix, the wholly-inserted predecessor also caused Bravo's
            // number to render inside <ins> with the reviewer's attribution.)
            Assert.Equal(new[] { "1.", "2.", "3.", "4." }, markers.Select(m => m.Number));
            Assert.True(markers[1].InsideIns, "the inserted paragraph's own number renders as an insertion");
            Assert.False(markers[2].InsideIns, "an unchanged paragraph after a wholly inserted one keeps a plain number");
            Assert.False(markers[3].InsideIns);
        }

        [Fact]
        public void TCN004_SplitParagraph_MarksTheSplitOffFollowerNumberInserted()
        {
            var wml = CreateNumberedListDocument(
                SplitFirstHalfParaXml("Alpha first half"),
                NumberedParaXml("Alpha second half"),
                NumberedParaXml("Bravo"));

            var html = ConvertWithTrackedChanges(wml);
            var markers = ExtractMarkers(html);

            // Both halves survive acceptance, so both count. The pilcrow inserted into
            // pre-existing content SPLIT the paragraph, making the follower's list
            // position the newly created one — its number renders as an insertion.
            Assert.Equal(new[] { "1.", "2.", "3." }, markers.Select(m => m.Number));
            Assert.True(markers[1].InsideIns, "the split-off follower's number renders as an insertion");
            Assert.False(markers[2].InsideIns);
        }

        [Fact]
        public void TCN005_DocxDiffRedline_RenumbersAsIfAccepted()
        {
            var left = CreateNumberedListDocument(
                NumberedParaXml("Alpha"),
                NumberedParaXml("Bravo"),
                NumberedParaXml("Charlie"));
            var right = CreateNumberedListDocument(
                NumberedParaXml("Alpha"),
                NumberedParaXml("Charlie"));

            var redline = DocxDiff.Compare(left, right, new DocxDiffSettings { AuthorForRevisions = "Reviewer" });

            var html = ConvertWithTrackedChanges(redline);
            var markers = ExtractMarkers(html);

            // The deleted item keeps its original marker. Charlie's automatic renumbering is itself
            // visible: old "3." deleted, new "2." inserted. Showing only the final "2." would make
            // the list-number cascade disappear from the redline.
            Assert.Equal(new[] { "1.", "2.", "3.", "2." }, markers.Select(m => m.Number));
            Assert.True(markers[1].InsideDel);
            Assert.True(markers[2].InsideDel);
            Assert.True(markers[3].InsideIns);
        }

        [Fact]
        public void TCN006_DocxDiffRedline_TracksInsertedItemCascade()
        {
            var left = CreateNumberedListDocument(
                NumberedParaXml("Alpha"),
                NumberedParaXml("Bravo"),
                NumberedParaXml("Charlie"));
            var right = CreateNumberedListDocument(
                NumberedParaXml("Inserted"),
                NumberedParaXml("Alpha"),
                NumberedParaXml("Bravo"),
                NumberedParaXml("Charlie"));

            var redline = DocxDiff.Compare(left, right, new DocxDiffSettings { AuthorForRevisions = "Reviewer" });

            using (var stream = new MemoryStream(redline.DocumentByteArray, writable: false))
            using (var document = WordprocessingDocument.Open(stream, false))
            {
                var originals = document.MainDocumentPart!.GetXDocument()
                    .Descendants(W.numberingChange)
                    .Select(change => (string?)change.Attribute(W.original))
                    .ToArray();
                Assert.Equal(new[] { "1.", "2.", "3." }, originals);
                Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document));
            }

            var html = ConvertWithTrackedChanges(redline);
            var markers = ExtractMarkers(html);

            Assert.Equal(
                new[] { "1.", "1.", "2.", "2.", "3.", "3.", "4." },
                markers.Select(m => m.Number));
            Assert.True(markers[0].InsideIns, "the new first item is inserted");
            Assert.True(markers[1].InsideDel);
            Assert.True(markers[2].InsideIns);
            Assert.True(markers[3].InsideDel);
            Assert.True(markers[4].InsideIns);
            Assert.True(markers[5].InsideDel);
            Assert.True(markers[6].InsideIns);
        }
    }
}
