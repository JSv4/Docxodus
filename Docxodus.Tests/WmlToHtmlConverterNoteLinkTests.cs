using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests
{
    /// <summary>
    /// Issue #667: the converter emitted a return-arrow hyperlink for every footnote/endnote
    /// definition it rendered, including definitions no story cites. The href then targeted an id
    /// nothing had emitted, and standalone export reported <c>fragment_target_unavailable</c> for
    /// markup the converter authored itself.
    ///
    /// <para>The invariant under test: every note navigation link the converter emits resolves to
    /// an id that is also in the output.</para>
    /// </summary>
    public class WmlToHtmlConverterNoteLinkTests
    {
        private static readonly XNamespace Xh = "http://www.w3.org/1999/xhtml";

        private static readonly string NvcaPath =
            Path.Combine("../../../../TestFiles/", "NVCA-Model-COI.docx");

        private static XElement Convert(byte[] docx) =>
            WmlToHtmlConverter.ConvertToHtml(
                new WmlDocument("note-links.docx", docx),
                new WmlToHtmlConverterSettings { RenderFootnotesAndEndnotes = true });

        /// <summary>Every <c>#fragment</c> href in the output whose target id is not in the output.</summary>
        private static List<string> DanglingFragmentTargets(XElement html)
        {
            var ids = new HashSet<string>(
                html.Descendants().Select(e => (string?)e.Attribute("id")).OfType<string>(),
                StringComparer.Ordinal);

            return html.Descendants(Xh + "a")
                .Select(a => (string?)a.Attribute("href"))
                .OfType<string>()
                .Where(href => href.StartsWith("#", StringComparison.Ordinal))
                .Select(href => href.Substring(1))
                .Where(target => !ids.Contains(target))
                .Distinct(StringComparer.Ordinal)
                .ToList();
        }

        private static List<XElement> LinksTargeting(XElement html, string hrefPrefix) =>
            html.Descendants(Xh + "a")
                .Where(a => ((string?)a.Attribute("href"))?.StartsWith(hrefPrefix, StringComparison.Ordinal) == true)
                .ToList();

        /// <summary>
        /// Rewrites one note-definition part of the NVCA fixture, so a test can construct the
        /// uncited-footnote and missing-definition shapes from a real document rather than from a
        /// programmatic package that would need every part hand-built.
        /// </summary>
        private static byte[] WithRewrittenNotesPart(string partSuffix, Action<XDocument> rewrite)
        {
            var bytes = File.ReadAllBytes(NvcaPath);
            using var ms = new MemoryStream();
            ms.Write(bytes, 0, bytes.Length);
            using (var wordDoc = WordprocessingDocument.Open(ms, true))
            {
                OpenXmlPart part = partSuffix == "footnotes"
                    ? wordDoc.MainDocumentPart!.FootnotesPart!
                    : wordDoc.MainDocumentPart!.EndnotesPart!;
                var xdoc = part.GetXDocument();
                rewrite(xdoc);
                part.PutXDocument();
            }

            return ms.ToArray();
        }

        /// <summary>
        /// The id of the first footnote the document actually cites. Not the first definition in
        /// the part: the fixture's part opens with <c>separator</c>, <c>continuationSeparator</c>
        /// and <c>continuationNotice</c> entries, none of which any story references.
        /// </summary>
        private static string FirstCitedFootnoteId()
        {
            var bytes = File.ReadAllBytes(NvcaPath);
            using var ms = new MemoryStream(bytes);
            using var wordDoc = WordprocessingDocument.Open(ms, false);
            return (string)wordDoc.MainDocumentPart!.GetXDocument()
                .Descendants(W.footnoteReference).First().Attribute(W.id)!;
        }

        [Fact]
        public void NvcaFixtureCarriesAnUncitedEndnoteDefinition()
        {
            // Guards the premise of the tests below: if the fixture ever stops carrying an
            // orphaned endnote, they would pass without exercising anything.
            var bytes = File.ReadAllBytes(NvcaPath);
            using var ms = new MemoryStream(bytes);
            using var wordDoc = WordprocessingDocument.Open(ms, false);

            var definedIds = wordDoc.MainDocumentPart!.EndnotesPart!.GetXDocument().Root!
                .Elements(W.endnote)
                .Where(en => (string?)en.Attribute(W.type) is not ("separator" or "continuationSeparator"))
                .Select(en => (string)en.Attribute(W.id)!)
                .ToList();
            var citedIds = wordDoc.MainDocumentPart.GetXDocument().Descendants(W.endnoteReference)
                .Select(r => (string)r.Attribute(W.id)!)
                .ToList();

            Assert.NotEmpty(definedIds);
            Assert.Empty(citedIds);
        }

        [Fact]
        public void UncitedEndnoteDefinitionIsRenderedButCarriesNoReturnArrow()
        {
            var html = Convert(File.ReadAllBytes(NvcaPath));

            // The definition is still rendered — an orphan is retained for package fidelity.
            var endnoteItem = html.Descendants(Xh + "li")
                .SingleOrDefault(li => (string?)li.Attribute("id") == "en-1");
            Assert.NotNull(endnoteItem);

            // …but nothing links back to a reference that was never emitted.
            Assert.Empty(LinksTargeting(html, "#en-ref-"));
            Assert.DoesNotContain("en-ref-1", DanglingFragmentTargets(html));
        }

        [Fact]
        public void NvcaConversionHasNoDanglingFragmentTargets()
        {
            Assert.Empty(DanglingFragmentTargets(Convert(File.ReadAllBytes(NvcaPath))));
        }

        [Fact]
        public void CitedFootnotesKeepBothDirectionsOfNavigation()
        {
            var html = Convert(File.ReadAllBytes(NvcaPath));

            var markers = LinksTargeting(html, "#fn-");
            var returnArrows = LinksTargeting(html, "#fn-ref-");

            // The fixture cites 94 footnotes; both directions survive for every one of them.
            Assert.Equal(94, returnArrows.Count);
            Assert.Equal(94, markers.Count - returnArrows.Count);
            Assert.All(markers, a => Assert.NotNull(a.Attribute("href")));
            Assert.Empty(DanglingFragmentTargets(html));
        }

        [Fact]
        public void AnUncitedFootnoteAndAnUncitedEndnoteProduceNoDanglingLinks()
        {
            // The fixture's endnote is already uncited; add an uncited footnote alongside it by
            // defining a note id that no w:footnoteReference in any story mentions.
            var docx = WithRewrittenNotesPart("footnotes", xdoc =>
            {
                var maxId = xdoc.Root!.Elements(W.footnote)
                    .Select(fn => int.Parse((string)fn.Attribute(W.id)!))
                    .Max();
                xdoc.Root.Add(new XElement(
                    W.footnote,
                    new XAttribute(W.id, maxId + 1),
                    new XElement(W.p, new XElement(W.r, new XElement(W.t, "Orphaned drafting note.")))));
            });

            var html = Convert(docx);

            Assert.Empty(DanglingFragmentTargets(html));
            // The 94 cited footnotes keep their arrows; the orphan gets none.
            Assert.Equal(94, LinksTargeting(html, "#fn-ref-").Count);
        }

        [Fact]
        public void AnIncrementalBlockRenderKeepsAMarkerWhoseDefinitionIsOutsideTheFragment()
        {
            // The completeness sweep asks "is the target id in this output?", which is only a fair
            // question of a whole document. The editor's incremental render builds a shell holding
            // just the requested blocks, so a citing paragraph rendered on its own has no notes
            // section to point at — and stripping its href there silently removed the citation's
            // navigation from the live DOM on every keystroke commit.
            using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
            session.ReplaceText(session.ListBlocks().Body[0].Id, "AAA");
            Assert.True(session.InsertFootnote(session.ListBlocks().Body[0].Id, 3, "A note.").Success);

            var citing = session.ListBlocks().Body[0].Id;
            var html = HtmlConversionOps.RenderBlockHtml(
                session, citing, new HtmlConversionOptions { RenderFootnotesAndEndnotes = true });

            Assert.Contains("class=\"footnote-ref\"", html);
            Assert.Contains("href=\"#fn-", html);
        }

        [Fact]
        public void AReferenceWhoseDefinitionIsMissingKeepsItsMarkerButLosesItsHref()
        {
            // The mirror of the reported bug: a cited note whose definition is absent from the
            // part. The marker carries the note's visible number, which is content, so it must
            // survive — only its unresolvable href is dropped.
            var removedId = FirstCitedFootnoteId();
            var docx = WithRewrittenNotesPart("footnotes", xdoc =>
                xdoc.Root!.Elements(W.footnote)
                    .Single(fn => (string?)fn.Attribute(W.id) == removedId)
                    .Remove());

            var html = Convert(docx);

            var inertMarker = html.Descendants(Xh + "a")
                .SingleOrDefault(a => (string?)a.Attribute("id") == $"fn-ref-{removedId}");
            Assert.NotNull(inertMarker);
            Assert.Null(inertMarker!.Attribute("href"));
            Assert.NotEmpty(inertMarker.Value);
            Assert.Empty(DanglingFragmentTargets(html));
        }
    }
}
