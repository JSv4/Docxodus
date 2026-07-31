#nullable enable

using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
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
    /// Word Compare's <c>White space</c> and <c>Case changes</c> comparison settings.
    /// </summary>
    public class WmlComparerWordParityOptionTests
    {
        private const string Nbsp = "\u00A0";
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static WmlDocument Doc(params string[] paragraphs) =>
            DocOf(paragraphs.Select(text => new Paragraph(new Run(Txt(text)))).ToArray());

        private static Text Txt(string text) =>
            new Text(text) { Space = SpaceProcessingModeValues.Preserve };

        private static WmlDocument DocOf(params OpenXmlElement[] bodyContent)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(bodyContent.Select(e => e.CloneNode(true))));

                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(
                    new DocDefaults(
                        new RunPropertiesDefault(
                            new RunPropertiesBaseStyle(
                                new RunFonts { Ascii = "Calibri" },
                                new FontSize { Val = "22" })),
                        new ParagraphPropertiesDefault()));

                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();

                doc.Save();
            }

            stream.Position = 0;
            return new WmlDocument("test.docx", stream.ToArray());
        }

        private static int RevisionCount(WmlDocument left, WmlDocument right, WmlComparerSettings settings)
        {
            var compared = WmlComparer.Compare(left, right, settings);
            AssertValid(compared);
            return WmlComparer.GetRevisions(compared, settings).Count;
        }

        private static void AssertValid(WmlDocument doc)
        {
            using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var errors = new OpenXmlValidator().Validate(wDoc).Select(e => e.Description).ToList();
            Assert.Empty(errors);
        }

        private static XElement BodyOf(WmlDocument doc)
        {
            using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
            using var wDoc = WordprocessingDocument.Open(ms, false);
            using var partStream = wDoc.MainDocumentPart!.GetStream();
            return XDocument.Load(partStream).Root!.Element(W + "body")!;
        }

        private static WmlDocument TableDoc(string cellText) => DocOf(
            new DocumentFormat.OpenXml.Wordprocessing.Table(
                new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Dxa }),
                new TableGrid(new GridColumn { Width = "5000" }),
                new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                        new TableCellProperties(new TableCellWidth { Width = "5000", Type = TableWidthUnitValues.Dxa }),
                        new Paragraph(new Run(Txt(cellText)))))),
            new Paragraph(new Run(Txt("after"))));

        private static string TextOf(WmlDocument doc) =>
            string.Concat(BodyOf(doc).Descendants(W + "t").Select(t => t.Value));

        // ---------------------------------------------------------------- White space

        [Fact]
        public void CompareWhitespace_defaults_true_so_behavior_is_unchanged()
        {
            Assert.True(new WmlComparerSettings().CompareWhitespace);
        }

        [Fact]
        public void CompareWhitespace_true_reports_a_doubled_space()
        {
            var settings = new WmlComparerSettings { CompareWhitespace = true };
            Assert.NotEqual(0, RevisionCount(Doc("Hello world."), Doc("Hello  world."), settings));
        }

        [Fact]
        public void CompareWhitespace_false_ignores_a_doubled_space()
        {
            var settings = new WmlComparerSettings { CompareWhitespace = false };
            Assert.Equal(0, RevisionCount(Doc("Hello world."), Doc("Hello  world."), settings));
        }

        [Fact]
        public void CompareWhitespace_false_ignores_a_trailing_space()
        {
            var settings = new WmlComparerSettings { CompareWhitespace = false };
            Assert.Equal(0, RevisionCount(Doc("Hello world."), Doc("Hello world. "), settings));
        }

        [Fact]
        public void CompareWhitespace_false_ignores_a_leading_space()
        {
            var settings = new WmlComparerSettings { CompareWhitespace = false };
            Assert.Equal(0, RevisionCount(Doc("Hello world."), Doc(" Hello world."), settings));
        }

        [Fact]
        public void CompareWhitespace_false_ignores_a_space_becoming_a_nonbreaking_space()
        {
            // Holds independently of ConflateBreakingAndNonbreakingSpaces, which is off here.
            var settings = new WmlComparerSettings
            {
                CompareWhitespace = false,
                ConflateBreakingAndNonbreakingSpaces = false,
            };
            Assert.Equal(0, RevisionCount(Doc("Hello world."), Doc("Hello" + Nbsp + "world."), settings));
        }

        [Fact]
        public void CompareWhitespace_false_still_reports_a_real_text_change()
        {
            var settings = new WmlComparerSettings { CompareWhitespace = false };
            Assert.NotEqual(0, RevisionCount(Doc("Hello  world."), Doc("Hello  there."), settings));
        }

        [Fact]
        public void CompareWhitespace_false_still_reports_an_inserted_word_between_spaces()
        {
            var settings = new WmlComparerSettings { CompareWhitespace = false };
            var revisions = WmlComparer.GetRevisions(
                WmlComparer.Compare(Doc("Hello world."), Doc("Hello  big  world."), settings), settings);

            Assert.Contains(revisions, r =>
                r.RevisionType == WmlComparer.WmlComparerRevisionType.Inserted && r.Text.Contains("big"));
        }

        [Fact]
        public void CompareWhitespace_false_ignores_whitespace_only_paragraph_change()
        {
            var settings = new WmlComparerSettings { CompareWhitespace = false };
            Assert.Equal(0, RevisionCount(
                Doc("First   paragraph.", "Second paragraph."),
                Doc("First paragraph.", "Second   paragraph."),
                settings));
        }

        [Fact]
        public void CompareWhitespace_false_output_carries_canonical_whitespace()
        {
            var settings = new WmlComparerSettings { CompareWhitespace = false };
            var compared = WmlComparer.Compare(Doc("Hello   world. "), Doc("Hello world."), settings);
            AssertValid(compared);
            Assert.Equal("Hello world.", TextOf(compared));
        }

        [Fact]
        public void CompareWhitespace_false_does_not_swallow_the_space_after_a_deleted_run()
        {
            // w:delText vanishes on accept; folding it into the whitespace run ate the following space.
            var left = DocOf(new Paragraph(
                new Run(Txt("Hello")),
                new DeletedRun(new Run(new DeletedText(" ") { Space = SpaceProcessingModeValues.Preserve }))
                {
                    Author = "someone",
                    Id = "1",
                },
                new Run(Txt(" world."))));

            var settings = new WmlComparerSettings { CompareWhitespace = false };
            Assert.Equal(0, RevisionCount(left, Doc("Hello world."), settings));
        }

        [Fact]
        public void CompareWhitespace_false_does_not_swallow_the_space_after_a_deleted_tab()
        {
            var left = DocOf(new Paragraph(
                new Run(Txt("Hello")),
                new DeletedRun(new Run(new TabChar())) { Author = "someone", Id = "1" },
                new Run(Txt(" world."))));

            var settings = new WmlComparerSettings { CompareWhitespace = false };
            Assert.Equal(0, RevisionCount(left, Doc("Hello world."), settings));
        }

        [Fact]
        public void CompareWhitespace_false_ignores_a_space_next_to_a_page_break()
        {
            var left = DocOf(new Paragraph(
                new Run(Txt("a ")),
                new Run(new Break { Type = BreakValues.Page }),
                new Run(Txt("b"))));
            var right = DocOf(new Paragraph(
                new Run(Txt("a")),
                new Run(new Break { Type = BreakValues.Page }),
                new Run(Txt(" b"))));

            var settings = new WmlComparerSettings { CompareWhitespace = false };
            Assert.Equal(0, RevisionCount(left, right, settings));
        }

        [Fact]
        public void CompareWhitespace_false_still_reports_a_tab_replaced_by_a_space()
        {
            // A tab is a whitespace boundary, not an equivalent of a space.
            var left = DocOf(new Paragraph(new Run(Txt("a")), new Run(new TabChar()), new Run(Txt("b"))));

            var settings = new WmlComparerSettings { CompareWhitespace = false };
            Assert.NotEqual(0, RevisionCount(left, Doc("a b"), settings));
        }

        [Fact]
        public void CompareWhitespace_false_applies_to_Consolidate()
        {
            var settings = new WmlComparerSettings { CompareWhitespace = false };
            var revised = new List<WmlRevisedDocumentInfo>
            {
                new WmlRevisedDocumentInfo
                {
                    RevisedDocument = Doc("Hello world."),
                    Revisor = "reviewer",
                    Color = DocxColors.LightBlue,
                },
            };

            var consolidated = WmlComparer.Consolidate(Doc("Hello   world."), revised, settings);
            AssertValid(consolidated);
            Assert.Empty(WmlComparer.GetRevisions(consolidated, settings));
        }

        [Fact]
        public void CompareWhitespace_false_ignores_a_space_next_to_a_tab()
        {
            var left = DocOf(new Paragraph(new Run(Txt("a ")), new Run(new TabChar()), new Run(Txt("b"))));
            var right = DocOf(new Paragraph(new Run(Txt("a")), new Run(new TabChar()), new Run(Txt(" b"))));

            var settings = new WmlComparerSettings { CompareWhitespace = false };
            Assert.Equal(0, RevisionCount(left, right, settings));
        }

        [Fact]
        public void CompareWhitespace_false_keeps_a_page_break()
        {
            var left = DocOf(new Paragraph(
                new Run(Txt("a  b")),
                new Run(new Break { Type = BreakValues.Page }),
                new Run(Txt("c"))));
            var right = DocOf(new Paragraph(
                new Run(Txt("a b")),
                new Run(new Break { Type = BreakValues.Page }),
                new Run(Txt("c"))));

            var settings = new WmlComparerSettings { CompareWhitespace = false };
            var compared = WmlComparer.Compare(left, right, settings);
            AssertValid(compared);
            Assert.Empty(WmlComparer.GetRevisions(compared, settings));
            Assert.Contains(
                BodyOf(compared).Descendants(W + "br"),
                br => (string?)br.Attribute(W + "type") == "page");
        }

        [Fact]
        public void CompareWhitespace_false_ignores_a_doubled_space_inside_a_table_cell()
        {
            var settings = new WmlComparerSettings { CompareWhitespace = false };
            Assert.Equal(0, RevisionCount(TableDoc("in  cell"), TableDoc("in cell"), settings));
        }

        // ---------------------------------------------------------------- Case changes

        [Fact]
        public void CaseInsensitive_defaults_false_so_case_changes_are_reported()
        {
            Assert.False(new WmlComparerSettings().CaseInsensitive);
            Assert.NotEqual(0, RevisionCount(Doc("Hello world."), Doc("HELLO world."), new WmlComparerSettings()));
        }

        [Fact]
        public void CaseInsensitive_true_needs_no_explicit_culture()
        {
            var settings = new WmlComparerSettings { CaseInsensitive = true };
            Assert.Equal(0, RevisionCount(Doc("Hello world."), Doc("HELLO world."), settings));
        }

        [Fact]
        public void CaseInsensitive_true_with_explicit_culture_ignores_case_only_change()
        {
            var settings = new WmlComparerSettings
            {
                CaseInsensitive = true,
                CultureInfo = CultureInfo.InvariantCulture,
            };
            Assert.Equal(0, RevisionCount(Doc("Hello world."), Doc("hello WORLD."), settings));
        }

        [Fact]
        public void CaseInsensitive_true_still_reports_a_real_text_change()
        {
            var settings = new WmlComparerSettings { CaseInsensitive = true };
            Assert.NotEqual(0, RevisionCount(Doc("Hello world."), Doc("HELLO there."), settings));
        }

        // ---------------------------------------------------------------- Both together

        [Fact]
        public void Whitespace_and_case_both_ignored_reports_nothing()
        {
            var settings = new WmlComparerSettings { CompareWhitespace = false, CaseInsensitive = true };
            Assert.Equal(0, RevisionCount(Doc("Hello world."), Doc("HELLO   WORLD. "), settings));
        }
    }
}
