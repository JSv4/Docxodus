using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;

namespace Docxodus;

public static partial class WmlToHtmlConverter
{
    private static readonly XName HiddenRevisionContentAttribute = XName.Get("hidden-revision-content", "urn:docxodus:rendering");
    private static readonly XName HiddenRevisionParagraphMarkAttribute = XName.Get("hidden-revision-paragraph-mark", "urn:docxodus:rendering");
    private static readonly XName JoinedParagraphSourceAttribute = XName.Get("joined-paragraph-source", "urn:docxodus:rendering");
    private static readonly XName JoinedParagraphSourcesAttribute = XName.Get("joined-paragraph-sources", "urn:docxodus:rendering");
    private static readonly XName NumberingRevisionAttribute = XName.Get("numbering-revision", "urn:docxodus:rendering");
    private static readonly XName EmptyRevisionDisplayAttribute = XName.Get("empty-revision-display", "urn:docxodus:rendering");
    private sealed class HiddenRevisionComments(IEnumerable<int> ids) : HashSet<int>(ids) { }
    private sealed class JoinedParagraphSource(XElement paragraph, string? anchorId)
    {
        public XElement Paragraph { get; } = paragraph;
        public string? AnchorId { get; } = anchorId;
        public bool HasFormattingAnchor { get; set; }
    }
    private sealed class JoinedParagraphSources : Dictionary<string, JoinedParagraphSource> { }

    private static bool IsFormattingSnapshot(XElement element) =>
        element.Name == W.rPrChange || element.Name == W.pPrChange || element.Name == W.sectPrChange ||
        element.Name == W.tblPrChange || element.Name == W.tblGridChange || element.Name == W.trPrChange ||
        element.Name == W.tblPrExChange || element.Name == W.tcPrChange;

    private static bool IsCommentBoundary(XElement element) =>
        element.Name == W.commentRangeStart || element.Name == W.commentRangeEnd;

    // Resolve display geometry on the converter's owned working tree, after source identities
    // and author colours are indexed. No package acceptance/rejection or source restamping.
    private static void PrepareRevisionVisibility(WordprocessingDocument document, WmlToHtmlConverterSettings settings)
    {
        document.MainDocumentPart!.RemoveAnnotations<HiddenRevisionComments>();
        if (!settings.RenderTrackedChanges) return;
        var roots = OwnedPartRelationships.StoryParts(document).Select(owner => owner.Part.GetXDocument().Root!).ToList();
        if (!settings.ShowFormattingChanges)
            foreach (var root in roots)
                root.Descendants().Where(e => IsFormattingSnapshot(e) || e.Name == W.numberingChange).Remove();
        if (settings.ShowInsertionsAndDeletions) return;
        // Numbering history normally renders as old/new list labels. Preserve its formatting
        // annotation separately when the text display uses only the current list label.
        foreach (var paragraph in roots.SelectMany(root => root.Descendants(W.p)))
            if (paragraph.Element(W.pPr)?.Element(W.numPr)?.Element(W.numberingChange) is { } change)
            {
                var marker = new XElement(change);
                if (settings.RevisionPresentation == RevisionPresentation.Word)
                {
                    marker.SetAttributeValue(WordFormatWording.BalloonData + "format-description", "Numbering changed");
                    string? author = (string?)marker.Attribute(W.author);
                    if (document.MainDocumentPart.Annotation<RevisionAuthorIndex>()?.IndexOf(author) is { } index)
                        marker.SetAttributeValue(WordFormatWording.BalloonData + "author-color", WordAuthorColor(settings, author, index));
                }
                paragraph.SetAttributeValue(NumberingRevisionAttribute, marker.ToString(SaveOptions.DisableFormatting));
            }
        var body = document.MainDocumentPart.GetXDocument().Root?.Element(W.body);
        bool hadBodyContent = body?.Descendants().Any(e => e.Name == W.p || e.Name == W.tbl) == true;
        var originalComments = roots.Where(root => root.Name != W.comments).SelectMany(root => CurrentContent(root))
            .Where(e => IsCommentBoundary(e) || e.Name == W.commentReference)
            .Select(e => (int?)e.Attribute(W.id)).OfType<int>().ToHashSet();

        foreach (var root in roots) Prepare(root);
        foreach (var root in roots)
            foreach (var group in RevisionParagraphGroups(root))
                foreach (var paragraph in group.Take(group.Count - 1))
                    foreach (var boundary in paragraph.ElementsAfterSelf().TakeWhile(e => IsBlockBoundaryMarker(e) || IsCommentBoundary(e)).ToArray())
                    {
                        boundary.Remove();
                        paragraph.Add(boundary);
                    }

        var visibleComments = new HashSet<int>();
        var commentStories = roots.Where(root => root.Name != W.comments)
            .SelectMany(root => root.Name == W.footnotes || root.Name == W.endnotes
                ? root.Elements().Where(e => e.Name == W.footnote || e.Name == W.endnote)
                : new[] { root });
        foreach (var root in commentStories)
        {
            var open = new HashSet<int>();
            foreach (var element in CurrentContent(root))
            {
                var id = (int?)element.Attribute(W.id);
                if (element.Name == W.commentRangeStart && id != null) open.Add(id.Value);
                else if (element.Name == W.commentRangeEnd && id != null) open.Remove(id.Value);
                else if (element.Name == W.commentReference && id != null) visibleComments.Add(id.Value);
                else if (element.Name == W.r && element.Elements().Any(e => e.Name != W.rPr)) visibleComments.UnionWith(open);
            }
        }
        var hiddenComments = new HiddenRevisionComments(originalComments.Except(visibleComments));
        document.MainDocumentPart.AddAnnotation(hiddenComments);
        foreach (var root in roots)
            root.Descendants().Where(e => (IsCommentBoundary(e) || e.Name == W.commentReference)
                && (int?)e.Attribute(W.id) is { } id && hiddenComments.Contains(id)).Remove();
        if (hadBodyContent && body != null && !CurrentContent(body).Any(e => e.Name == W.p || e.Name == W.tbl))
        {
            // A deliberately empty display still has one page. This generated line has no
            // source identity; its explicit marker permits only an empty PageMap inventory.
            var blank = new XElement(W.p, new XAttribute(EmptyRevisionDisplayAttribute, true));
            if (body.Element(W.sectPr) is { } section) section.AddBeforeSelf(blank); else body.Add(blank);
        }

        static IEnumerable<XElement> CurrentContent(XElement element)
        {
            if (IsFormattingSnapshot(element) || element.Attribute(HiddenRevisionContentAttribute) != null) yield break;
            yield return element;
            foreach (var child in element.Elements())
                foreach (var descendant in CurrentContent(child)) yield return descendant;
        }

        static void Hide(XElement element, bool remove)
        {
            // Range ends inside hidden content must still close a surviving range; starts must
            // still open it. Move only zero-width boundaries, preserving their document order.
            var boundaries = element.Descendants().Where(e => IsCommentBoundary(e) || IsBlockBoundaryMarker(e)).ToArray();
            foreach (var boundary in boundaries) { boundary.Remove(); element.AddBeforeSelf(boundary); }
            if (remove) element.Remove();
            else element.SetAttributeValue(HiddenRevisionContentAttribute, true);
        }

        static void Prepare(XElement element)
        {
            if (IsFormattingSnapshot(element)) return;
            if ((element.Name == W.del || element.Name == W.moveFrom)
                && element.Parent?.Name != W.rPr && element.Parent?.Name != W.trPr)
            {
                // Keep field delimiters available to the existing field grouping/visibility
                // logic, while flow and required-textbox inventories skip this hidden subtree.
                Hide(element, remove: false);
                return;
            }
            if (element.Name == W.tr)
            {
                if (element.Elements(W.trPr).Elements(W.del).Any()) { Hide(element, remove: true); return; }
                bool currentGrid = TableGridModel.RetainedCellsFillGrid(element);
                XElement? survivor = null;
                foreach (var cell in TableGridModel.RowGrid(element))
                {
                    if (!cell.Cell.Elements(W.tcPr).Elements(W.cellDel).Any()) { survivor = cell.Cell; continue; }
                    if (!currentGrid && survivor != null)
                    {
                        var properties = survivor.Element(W.tcPr);
                        if (properties == null) { properties = new XElement(W.tcPr); survivor.AddFirst(properties); }
                        var span = properties.Element(W.gridSpan);
                        if (span == null) { span = new XElement(W.gridSpan, new XAttribute(W.val, 1)); properties.Add(span); }
                        span.SetAttributeValue(W.val, ((int?)span.Attribute(W.val) ?? 1) + cell.Span);
                    }
                    Hide(cell.Cell, remove: true);
                }
            }
            if (element.Name == W.tc && element.Element(W.tcPr) is { } cellProperties
                && (string?)cellProperties.Element(W.cellMerge)?.Attribute(W.vMerge) is { } merge
                && (merge == "rest" || merge == "cont"))
            {
                cellProperties.Elements(W.vMerge).Remove();
                cellProperties.Add(new XElement(W.vMerge, new XAttribute(W.val, merge == "rest" ? "restart" : "continue")));
            }
            if (element.Name == W.p && element.Elements(W.pPr).Elements(W.rPr)
                .Elements().Any(e => e.Name == W.del || e.Name == W.moveFrom))
                element.SetAttributeValue(HiddenRevisionParagraphMarkAttribute, true);
            if (element.Name == W.p)
                element.Elements(W.pPr).Elements(W.rPr).Elements()
                    .Where(e => e.Name == W.ins || e.Name == W.del || e.Name == W.moveFrom || e.Name == W.moveTo).Remove();
            bool hadRows = element.Name == W.tbl && WordprocessingMLUtil.TableRows(element).Any();
            foreach (var child in element.Elements().ToArray()) Prepare(child);
            if (hadRows && !WordprocessingMLUtil.TableRows(element).Any()) Hide(element, remove: true);
        }
    }

    private static void PrepareRevisionParagraphGroups(WordprocessingDocument document, WmlToHtmlConverterSettings settings)
    {
        if (!settings.RenderTrackedChanges || settings.ShowInsertionsAndDeletions) return;
        // Wording is stamped from the original direct properties first. Current group
        // properties then determine inherited fonts and numbering during assembly.
        foreach (var owner in OwnedPartRelationships.StoryParts(document))
            foreach (var group in RevisionParagraphGroups(owner.Part.GetXDocument().Root!)) ResolveParagraphGroup(group);
    }

    // Discover groups after removed tables/rows have exposed their boundary markers.
    private static IEnumerable<List<XElement>> RevisionParagraphGroups(XElement root)
    {
        foreach (var parent in root.DescendantsAndSelf().Where(e => e.Elements(W.p).Any()))
        {
            var group = new List<XElement>();
            foreach (var child in parent.Elements())
            {
                if (IsBlockBoundaryMarker(child) || IsCommentBoundary(child)) continue;
                if (child.Name != W.p || group.Count > 0 && !HasHiddenParagraphMark(group[^1]))
                {
                    if (group.Count > 1 && group.Any(p => p.Attribute(HiddenRevisionParagraphMarkAttribute) != null)) yield return group;
                    group = new List<XElement>();
                }
                if (child.Name == W.p) group.Add(child);
            }
            if (group.Count > 1 && group.Any(p => p.Attribute(HiddenRevisionParagraphMarkAttribute) != null)) yield return group;
        }
    }

    private static void ResolveParagraphGroup(List<XElement> group)
    {
        var current = new XElement(group[^1].Element(W.pPr) ?? new XElement(W.pPr));
        static bool MarkRevision(XElement e) => e.Name == W.ins || e.Name == W.del ||
            e.Name == W.moveFrom || e.Name == W.moveTo || IsFormattingSnapshot(e);
        for (int index = 0; index < group.Count; index++)
        {
            var paragraph = group[index];
            var original = paragraph.Element(W.pPr);
            var properties = new XElement(current);
            properties.Elements().Where(e => IsFormattingSnapshot(e) || e.Name == W.numberingChange).Remove();
            properties.Elements(W.numPr).Elements(W.numberingChange).Remove();
            properties.Elements(W.rPr).Elements().Where(MarkRevision).Remove();
            if (index != group.Count - 1) properties.Elements(W.sectPr).Remove();
            if (original != null)
            {
                properties.Add(original.Elements().Where(e => IsFormattingSnapshot(e) || e.Name == W.numberingChange));
                var markRevisions = original.Elements(W.rPr).Elements().Where(MarkRevision).ToList();
                if (markRevisions.Count > 0)
                {
                    var mark = properties.Element(W.rPr);
                    if (mark == null) { mark = new XElement(W.rPr); properties.Add(mark); }
                    mark.Add(markRevisions);
                }
            }
            // Only the group head owns the list label; explicit numId=0 also overrides a
            // numbered paragraph style on its continuation members.
            if (index > 0)
            {
                properties.Elements(W.numPr).Remove();
                properties.Add(new XElement(W.numPr, new XElement(W.numId, new XAttribute(W.val, 0))));
            }
            if (original == null) paragraph.AddFirst(properties); else original.ReplaceWith(properties);
        }
    }

    private static void MergeRevisionParagraphGroups(WordprocessingDocument document, WmlToHtmlConverterSettings settings)
    {
        var main = document.MainDocumentPart!;
        main.RemoveAnnotations<JoinedParagraphSources>();
        if (!settings.RenderTrackedChanges || settings.ShowInsertionsAndDeletions) return;
        var sources = new JoinedParagraphSources();
        main.AddAnnotation(sources);
        foreach (var owner in OwnedPartRelationships.StoryParts(document))
        {
            var root = owner.Part.GetXDocument().Root!;
            root.RemoveAnnotations<OpenXmlPart>();
            root.AddAnnotation(owner.Part);
            foreach (var group in RevisionParagraphGroups(root).ToList())
            {
                var head = group[0];
                var keys = new List<string>();
                // One content stream lets fields and tabs cross the removed mark. Source IDs
                // travel with the original inline content, never with a fabricated paragraph.
                foreach (var paragraph in group.Skip(1))
                {
                    string key = sources.Count.ToString(CultureInfo.InvariantCulture);
                    keys.Add(key);
                    sources.Add(key, new JoinedParagraphSource(new XElement(W.p, paragraph.Attributes(), paragraph.Element(W.pPr)),
                        settings.SourceAnchorIdentityProvider?.Invoke(paragraph)));
                    foreach (var child in paragraph.Elements().Where(e => e.Name != W.pPr && e.Name != PtOpenXml.pPr).ToList())
                    {
                        child.SetAttributeValue(JoinedParagraphSourceAttribute, key);
                        child.Remove();
                        head.Add(child);
                    }
                    if (paragraph.Element(W.pPr)?.Element(W.sectPr) is { } section)
                    {
                        var currentSection = new XElement(section);
                        currentSection.Elements(W.sectPrChange).Remove();
                        head.Element(W.pPr)!.Add(currentSection);
                    }
                    paragraph.Remove();
                }
                head.Attribute(HiddenRevisionParagraphMarkAttribute)?.Remove();
                head.SetAttributeValue(JoinedParagraphSourcesAttribute, string.Join(",", keys));
            }
        }
    }

    private static object? WithJoinedParagraphSource(WordprocessingDocument document,
        WmlToHtmlConverterSettings settings, XElement source, object? content)
    {
        if (content == null || (string?)source.Attribute(JoinedParagraphSourceAttribute) is not { } key ||
            document.MainDocumentPart!.Annotation<JoinedParagraphSources>() is not { } sources ||
            !sources.TryGetValue(key, out var origin)) return content;
        var span = new XElement(Xhtml.span, content);
        // Instruction-only and hidden runs cannot stand in for a visible source fragment.
        if (span.Value.Length == 0 && !span.Descendants().Any(e => e.Name == Xhtml.img ||
            e.Name.LocalName == "svg" || e.Name == Xhtml.br || e.Attribute("data-docx-tab") != null ||
            e.Attribute("data-field") != null || e.Attribute("data-docx-textbox") != null)) return span.Nodes().ToList();
        if (origin.AnchorId != null) span.SetAttributeValue("data-source-anchor-id", origin.AnchorId);
        if (settings.StampAnchors && origin.Paragraph.Attribute(PtOpenXml.Unid) is { } unid)
            span.SetAttributeValue("data-anchor", unid.Value);
        span.AddAnnotation(origin);
        if (!origin.HasFormattingAnchor)
        {
            ApplyParagraphRevisionMarkup(settings, origin.Paragraph, span);
            origin.HasFormattingAnchor = true;
        }
        return span;
    }

    private static XElement PreserveJoinedFieldSources(XElement field)
    {
        // Page-number substitution replaces the marker's children. Keep each real joined
        // source around the atomic field so its identity and formatting survive substitution.
        var origins = field.Descendants().Where(e => e.Annotation<JoinedParagraphSource>() != null)
            .GroupBy(e => e.Annotation<JoinedParagraphSource>()!).ToList();
        foreach (var origin in origins)
        {
            var wrapper = new XElement(Xhtml.span, origin.First().Attributes());
            foreach (var span in origin) span.Attributes().Remove();
            wrapper.Add(field);
            field = wrapper;
        }
        return field;
    }

    private static void ApplyUnplacedJoinedFormatting(WordprocessingDocument document,
        WmlToHtmlConverterSettings settings, XElement source, XElement paragraph)
    {
        if ((string?)source.Attribute(JoinedParagraphSourcesAttribute) is not { } keys ||
            document.MainDocumentPart!.Annotation<JoinedParagraphSources>() is not { } sources) return;
        foreach (string key in keys.Split(','))
        {
            var origin = sources[key];
            if (origin.HasFormattingAnchor) continue;
            var span = new XElement(Xhtml.span);
            ApplyParagraphRevisionMarkup(settings, origin.Paragraph, span);
            if (!span.HasAttributes) continue;
            // An empty continuation has no source fragment of its own. Its formatting change
            // still describes the joined paragraph, so anchor that annotation to its content.
            foreach (var node in paragraph.Nodes().ToList()) { node.Remove(); span.Add(node); }
            paragraph.Add(span);
        }
    }
}
