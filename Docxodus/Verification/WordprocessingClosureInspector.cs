// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
using System.Xml.Linq;

namespace Docxodus.Verification;

/// <summary>
/// Cross-part checks that neither OPC relationship validation nor the Open XML SDK schema
/// validator performs. The scanner consumes only XML retained by the bounded manifest pass.
/// </summary>
internal static class WordprocessingClosureInspector
{
    private const string TransitionalWord =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StrictWord = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    private const string Word2010 = "http://schemas.microsoft.com/office/word/2010/wordml";
    internal static DeliverableCheckResult Inspect(
        PackageManifestInspection inspection,
        WordprocessingInspectionGraph graph,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings,
        DeliverableInspectionBudget budget)
    {
        int before = observations.Count;
        var sink = new FindingSink(observations, maximumFindings, budget);
        var storyParts = graph.StoryParts;
        try
        {
            // Count while traversing, before any detector can accidentally turn a large retained
            // XML tree into unbounded repeated work.
            if (!sink.Stopped)
                foreach (var part in storyParts.Concat(graph.NumberingParts))
                foreach (var _ in part.Xml!.Root!.DescendantsAndSelf())
                    if (!budget.Node() || !budget.Step()) break;

            if (!sink.Stopped) InspectSemanticMultiplicity(graph, sink);
            if (!sink.Stopped) InspectMediaClosure(graph.ReachableRelationships, sink);
            if (!sink.Stopped) InspectBookmarks(storyParts, sink);
            if (!sink.Stopped) InspectComments(storyParts, graph.CommentParts, sink);
            if (!sink.Stopped)
                InspectNotes(storyParts, graph.FootnoteParts, sink,
                    "footnote", "footnoteReference", "/word/footnotes.xml");
            if (!sink.Stopped)
                InspectNotes(storyParts, graph.EndnoteParts, sink,
                    "endnote", "endnoteReference", "/word/endnotes.xml");
            if (!sink.Stopped) InspectMoves(storyParts, sink);
            if (!sink.Stopped) InspectContentControls(graph, storyParts, sink);
            if (!sink.Stopped) InspectNumbering(storyParts, graph.NumberingParts, sink);
            if (!sink.Stopped) InspectFields(storyParts, sink);
            if (!sink.Stopped) InspectStaticRenderRisks(graph, storyParts, sink);
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            sink.Add("structure.closure_inspection_unavailable",
                DeliverableFindingCategory.Structure, VerificationFindingSeverity.Error,
                $"Bounded Wordprocessing closure inspection failed ({exception.GetType().Name}).", "/",
                "Repair the implicated Wordprocessing markup before delivery.",
                new ChangeLocation { PropertyPath = "wordprocessingClosure" },
                subject: exception.GetType().FullName);
            sink.ForceUnavailable(exception.GetType().Name);
        }

        return new DeliverableCheckResult
        {
            Check = "wordprocessing_closure",
            Status = sink.Stopped
                ? DeliverableCheckStatus.UnavailableEvidence
                : DeliverableCheckStatus.Completed,
            FindingCount = observations.Count - before,
            Diagnostic = sink.Diagnostic,
        };
    }

    private static void InspectSemanticMultiplicity(
        WordprocessingInspectionGraph graph,
        FindingSink sink)
    {
        ReportSingletonRole(graph, sink, WordprocessingInspectionGraph.SemanticRole.Comments,
            "comments", "structure.comments_part_ambiguous");
        ReportSingletonRole(graph, sink, WordprocessingInspectionGraph.SemanticRole.Settings,
            "settings", "structure.settings_part_ambiguous");
        ReportSingletonRole(graph, sink, WordprocessingInspectionGraph.SemanticRole.Endnotes,
            "endnotes", "structure.endnotes_part_ambiguous");
        ReportSingletonRole(graph, sink, WordprocessingInspectionGraph.SemanticRole.Footnotes,
            "footnotes", "structure.footnotes_part_ambiguous");
        ReportSingletonRole(graph, sink, WordprocessingInspectionGraph.SemanticRole.Numbering,
            "numbering", "structure.numbering_part_ambiguous");
        ReportSingletonRole(graph, sink, WordprocessingInspectionGraph.SemanticRole.Styles,
            "styles", "structure.styles_part_ambiguous");

        foreach (var group in graph.SemanticRelationshipEdges
                     .Where(edge => edge.OwnerRole == WordprocessingInspectionGraph.SemanticRole.CustomXml
                         && edge.TargetRole
                         == WordprocessingInspectionGraph.SemanticRole.CustomXmlProperties)
                     .GroupBy(edge => edge.Relationship.OwnerUri, StringComparer.OrdinalIgnoreCase)
                     .Where(group => group.Count() > 1))
        foreach (var edge in group)
            sink.Add("structure.custom_xml_properties_part_ambiguous",
                DeliverableFindingCategory.Structure, VerificationFindingSeverity.Error,
                $"Custom XML item '{group.Key}' owns more than one properties relationship.",
                group.Key,
                "Retain exactly one customXmlProps relationship for each custom XML item.",
                new ChangeLocation
                {
                    OwnerUri = group.Key,
                    RelationshipId = edge.Relationship.Id,
                    TargetUri = edge.Relationship.ResolvedTargetUri ?? edge.Relationship.Target,
                },
                subject: edge.Relationship.Id);
    }

    private static void ReportSingletonRole(
        WordprocessingInspectionGraph graph,
        FindingSink sink,
        WordprocessingInspectionGraph.SemanticRole role,
        string roleName,
        string code)
    {
        var edges = graph.SemanticRelationshipEdges.Where(edge =>
            edge.OwnerRole == WordprocessingInspectionGraph.SemanticRole.MainDocument
            && edge.TargetRole == role).ToArray();
        if (edges.Length <= 1) return;
        foreach (var edge in edges)
            sink.Add(code, DeliverableFindingCategory.Structure, VerificationFindingSeverity.Error,
                $"The main document owns more than one {roleName} relationship.",
                edge.Relationship.OwnerUri,
                $"Retain exactly one valid {roleName} relationship on the main document.",
                new ChangeLocation
                {
                    OwnerUri = edge.Relationship.OwnerUri,
                    RelationshipId = edge.Relationship.Id,
                    TargetUri = edge.Relationship.ResolvedTargetUri ?? edge.Relationship.Target,
                },
                subject: edge.Relationship.Id);
    }

    private static void InspectMediaClosure(
        IReadOnlyList<PackageRelationship> relationships,
        FindingSink sink)
    {
        foreach (var relationship in relationships.Where(relationship =>
                     relationship.TargetMode == "Internal"
                     && relationship.IsTargetPresent == false
                     && OpenXmlRelationshipVocabulary.IsOfficeType(relationship.Type, "image")))
        {
            sink.Add("relationship.media_target_missing", DeliverableFindingCategory.Relationship,
                VerificationFindingSeverity.Error,
                "An image relationship points to a media part that is not present in the package.",
                relationship.OwnerUri,
                "Restore the media part or remove and rewrite the referencing drawing.",
                new ChangeLocation
                {
                    OwnerUri = relationship.OwnerUri,
                    RelationshipId = relationship.Id,
                    TargetUri = relationship.ResolvedTargetUri ?? relationship.Target,
                },
                subject: relationship.Id);
        }
    }

    private static void InspectBookmarks(
        IReadOnlyList<PackageManifestInspectionEntry> parts,
        FindingSink sink)
    {
        var names = new Dictionary<string, List<(PackageManifestInspectionEntry Part, XElement Element)>>(
            StringComparer.Ordinal);
        foreach (var part in parts)
        {
            var startsById = new Dictionary<string, List<(XElement Element, int Position)>>(StringComparer.Ordinal);
            var endsById = new Dictionary<string, List<(XElement Element, int Position)>>(StringComparer.Ordinal);
            int position = 0;
            foreach (var marker in part.Xml!.Descendants().Where(element =>
                         IsElement(element, "bookmarkStart") || IsElement(element, "bookmarkEnd")))
            {
                position++;
                var id = WordAttribute(marker, "id");
                if (string.IsNullOrEmpty(id))
                {
                    AddElementFinding(sink, part, marker,
                        IsElement(marker, "bookmarkStart")
                            ? "structure.bookmark_id_missing" : "structure.bookmark_end_id_missing",
                        VerificationFindingSeverity.Error,
                        "A bookmark marker has no w:id.",
                        "Assign a story-part-unique numeric w:id and pair the bookmark markers.",
                        WordAttribute(marker, "name"));
                    continue;
                }
                var index = IsElement(marker, "bookmarkStart") ? startsById : endsById;
                if (!index.TryGetValue(id, out var occurrences)) index[id] = occurrences = new();
                occurrences.Add((marker, position));
            }

            foreach (var startOccurrence in startsById.Values.SelectMany(value => value))
            {
                var start = startOccurrence.Element;
                var id = WordAttribute(start, "id");
                var name = WordAttribute(start, "name");
                if (!string.IsNullOrEmpty(name))
                {
                    if (!names.TryGetValue(name, out var values))
                        names.Add(name, values = new());
                    values.Add((part, start));
                }
            }

            foreach (var id in startsById.Keys.Concat(endsById.Keys).Distinct(StringComparer.Ordinal))
            {
                var starts = startsById.GetValueOrDefault(id) ?? new();
                var ends = endsById.GetValueOrDefault(id) ?? new();
                var marker = starts.FirstOrDefault().Element ?? ends[0].Element;
                if (starts.Count != 1 || ends.Count != 1)
                    AddElementFinding(sink, part, marker, "structure.bookmark_pair_invalid",
                        VerificationFindingSeverity.Error,
                        $"Bookmark id '{id}' has {starts.Count} start marker(s) and {ends.Count} end marker(s) in this part.",
                        "Use exactly one bookmark start and one bookmark end for each id within its story part.", id);
                else if (ends[0].Position <= starts[0].Position)
                    AddElementFinding(sink, part, marker, "structure.bookmark_order_invalid",
                        VerificationFindingSeverity.Error,
                        $"Bookmark end id '{id}' occurs before its start.",
                        "Place each bookmark end after its matching start in the same story part.", id);
            }
            InspectRangeNesting(part, "bookmarkStart", "bookmarkEnd",
                "structure.bookmark_nesting_invalid", sink);
        }

        foreach (var duplicate in names.Where(pair => pair.Value.Count > 1))
        {
            foreach (var occurrence in duplicate.Value)
            {
                AddElementFinding(sink, occurrence.Part, occurrence.Element,
                    "structure.bookmark_name_duplicate", VerificationFindingSeverity.Error,
                    $"Bookmark name '{duplicate.Key}' is declared more than once.",
                    "Rename or remove duplicate bookmark declarations so cross-references resolve unambiguously.",
                    duplicate.Key);
            }
        }

        foreach (var part in parts)
        foreach (var hyperlink in Descendants(part, "hyperlink"))
        {
            var anchor = WordAttribute(hyperlink, "anchor");
            if (!string.IsNullOrEmpty(anchor) && !names.ContainsKey(anchor))
            {
                AddElementFinding(sink, part, hyperlink, "structure.hyperlink_bookmark_missing",
                    VerificationFindingSeverity.Error,
                    $"Internal hyperlink target '{anchor}' does not name a bookmark.",
                    "Restore the bookmark or retarget/remove the internal hyperlink.", anchor);
            }
        }
    }

    private static void InspectComments(
        IReadOnlyList<PackageManifestInspectionEntry> storyParts,
        IReadOnlyList<PackageManifestInspectionEntry> commentParts,
        FindingSink sink)
    {
        foreach (var part in commentParts)
        foreach (var definition in DirectChildren(part, "comment")
                     .Where(element => string.IsNullOrEmpty(WordAttribute(element, "id"))))
            AddElementFinding(sink, part, definition, "structure.comment_id_missing",
                VerificationFindingSeverity.Error, "A comment definition has no w:id.",
                "Assign the comment a unique numeric id and update its markers.", null);
        foreach (var part in storyParts)
        foreach (var marker in Descendants(part, "commentReference")
                     .Concat(Descendants(part, "commentRangeStart"))
                     .Concat(Descendants(part, "commentRangeEnd"))
                     .Where(element => string.IsNullOrEmpty(WordAttribute(element, "id"))))
            AddElementFinding(sink, part, marker, "structure.comment_marker_id_missing",
                VerificationFindingSeverity.Error, "A comment marker has no w:id.",
                "Assign the marker its comment definition id or remove it.", null);
        var definitions = commentParts.SelectMany(part => DirectChildren(part, "comment")
                .Select(element => (Part: part, Element: element, Id: WordAttribute(element, "id"))))
            .Where(item => item.Id is not null)
            .ToArray();
        var references = storyParts.SelectMany(part =>
                Descendants(part, "commentReference")
                    .Concat(Descendants(part, "commentRangeStart"))
                    .Concat(Descendants(part, "commentRangeEnd"))
                    .Select(element => (Part: part, Element: element, Id: WordAttribute(element, "id"))))
            .Where(item => item.Id is not null)
            .ToArray();

        foreach (var duplicate in definitions.GroupBy(item => item.Id!, StringComparer.Ordinal)
                     .Where(group => group.Count() > 1))
        foreach (var item in duplicate)
            AddElementFinding(sink, item.Part, item.Element, "structure.comment_id_duplicate",
                VerificationFindingSeverity.Error,
                $"Comment id '{duplicate.Key}' has multiple definitions.",
                "Give each comment definition a unique numeric id and update all markers.", duplicate.Key);

        var definitionIds = definitions.Select(item => item.Id!).ToHashSet(StringComparer.Ordinal);
        var referenceIds = references.Select(item => item.Id!).ToHashSet(StringComparer.Ordinal);
        foreach (var reference in references.Where(item => !definitionIds.Contains(item.Id!)))
            AddElementFinding(sink, reference.Part, reference.Element,
                "structure.comment_definition_missing", VerificationFindingSeverity.Error,
                $"Comment marker id '{reference.Id}' has no definition.",
                "Restore the comment definition or remove all markers for this id.", reference.Id);

        foreach (var definition in definitions.Where(item => !referenceIds.Contains(item.Id!)))
            AddElementFinding(sink, definition.Part, definition.Element,
                "structure.comment_unreachable", VerificationFindingSeverity.Warning,
                $"Comment definition id '{definition.Id}' is not referenced by document markup.",
                "Remove the orphan definition or restore its range/reference markers.", definition.Id);

        foreach (var group in references.GroupBy(
                     item => (item.Part.Uri, Id: item.Id!),
                     StringTupleComparer.Instance))
        {
            int starts = group.Count(item => IsElement(item.Element, "commentRangeStart"));
            int ends = group.Count(item => IsElement(item.Element, "commentRangeEnd"));
            int marks = group.Count(item => IsElement(item.Element, "commentReference"));
            if (starts == 1 && ends == 1)
            {
                var start = group.Single(item => IsElement(item.Element, "commentRangeStart"));
                var end = group.Single(item => IsElement(item.Element, "commentRangeEnd"));
                if (XNode.CompareDocumentOrder(start.Element, end.Element) >= 0)
                    AddElementFinding(sink, start.Part, start.Element,
                        "structure.comment_range_order_invalid", VerificationFindingSeverity.Error,
                        $"Comment range end id '{group.Key.Id}' does not follow its start.",
                        "Place the range end after its matching start within the story.", group.Key.Id);
            }
            if (starts == ends && starts <= 1 && marks > 0) continue;
            var item = group.First();
            AddElementFinding(sink, item.Part, item.Element, "structure.comment_markers_invalid",
                VerificationFindingSeverity.Error,
                $"Comment id '{group.Key.Id}' in '{group.Key.Uri}' has {starts} range start(s), {ends} range end(s), and {marks} reference mark(s).",
                "Pair comment range markers within each story part and retain at least one comment reference mark.",
                group.Key.Id);
        }
        foreach (var part in storyParts)
            InspectRangeNesting(part, "commentRangeStart", "commentRangeEnd",
                "structure.comment_range_nesting_invalid", sink);
    }

    private static void InspectNotes(
        IReadOnlyList<PackageManifestInspectionEntry> storyParts,
        IReadOnlyList<PackageManifestInspectionEntry> definitionParts,
        FindingSink sink,
        string definitionName,
        string referenceName,
        string conventionalPartUri)
    {
        foreach (var part in definitionParts)
        foreach (var definition in DirectChildren(part, definitionName).Where(element =>
                     WordAttribute(element, "type") is not ("separator" or "continuationSeparator")
                     && string.IsNullOrEmpty(WordAttribute(element, "id"))))
            AddElementFinding(sink, part, definition,
                $"structure.{definitionName}_id_missing", VerificationFindingSeverity.Error,
                $"A {definitionName} definition has no w:id.",
                $"Assign the {definitionName} a unique id and update its reference.", null);
        foreach (var part in storyParts)
        foreach (var reference in Descendants(part, referenceName)
                     .Where(element => string.IsNullOrEmpty(WordAttribute(element, "id"))))
            AddElementFinding(sink, part, reference,
                $"structure.{definitionName}_reference_id_missing", VerificationFindingSeverity.Error,
                $"A {definitionName} reference has no w:id.",
                $"Assign the reference its {definitionName} id or remove it.", null);
        var definitions = definitionParts.SelectMany(part => DirectChildren(part, definitionName)
                .Select(element => (Part: part, Element: element, Id: WordAttribute(element, "id"),
                    Type: WordAttribute(element, "type"))))
            .Where(item => item.Type is not ("separator" or "continuationSeparator"))
            .Where(item => item.Id is not null)
            .ToArray();
        var references = storyParts.SelectMany(part => Descendants(part, referenceName)
                .Select(element => (Part: part, Element: element, Id: WordAttribute(element, "id"))))
            .Where(item => item.Id is not null)
            .ToArray();

        foreach (var duplicate in definitions.GroupBy(item => item.Id!, StringComparer.Ordinal)
                     .Where(group => group.Count() > 1))
        foreach (var item in duplicate)
            AddElementFinding(sink, item.Part, item.Element,
                $"structure.{definitionName}_id_duplicate", VerificationFindingSeverity.Error,
                $"{Title(definitionName)} id '{duplicate.Key}' has multiple definitions.",
                $"Give each {definitionName} definition a unique id and update its references.", duplicate.Key);

        var definitionIds = definitions.Select(item => item.Id!).ToHashSet(StringComparer.Ordinal);
        var referenceIds = references.Select(item => item.Id!).ToHashSet(StringComparer.Ordinal);
        foreach (var reference in references.Where(item => !definitionIds.Contains(item.Id!)))
            AddElementFinding(sink, reference.Part, reference.Element,
                $"structure.{definitionName}_definition_missing", VerificationFindingSeverity.Error,
                $"{Title(definitionName)} reference id '{reference.Id}' has no definition in {conventionalPartUri}.",
                $"Restore the {definitionName} definition or remove the reference.", reference.Id);
        foreach (var definition in definitions.Where(item => !referenceIds.Contains(item.Id!)))
            AddElementFinding(sink, definition.Part, definition.Element,
                $"structure.{definitionName}_unreachable", VerificationFindingSeverity.Warning,
                $"{Title(definitionName)} definition id '{definition.Id}' is not referenced.",
                $"Remove the orphan {definitionName} or restore its reference.", definition.Id);
    }

    private static void InspectMoves(
        IReadOnlyList<PackageManifestInspectionEntry> parts,
        FindingSink sink)
    {
        foreach (var part in parts)
        {
            var wrapperFrom = MoveMarkers(part, "moveFrom");
            var wrapperTo = MoveMarkers(part, "moveTo");
            var rangeFromStarts = MoveMarkers(part, "moveFromRangeStart");
            var rangeFromEnds = MoveMarkers(part, "moveFromRangeEnd");
            var rangeToStarts = MoveMarkers(part, "moveToRangeStart");
            var rangeToEnds = MoveMarkers(part, "moveToRangeEnd");
            var all = wrapperFrom.Concat(wrapperTo)
                .Concat(rangeFromStarts).Concat(rangeFromEnds)
                .Concat(rangeToStarts).Concat(rangeToEnds)
                .ToArray();
            foreach (var item in all.Where(item => string.IsNullOrEmpty(item.Id)))
                AddElementFinding(sink, part, item.Element, "structure.move_id_missing",
                    VerificationFindingSeverity.Error,
                    "Tracked move markup has no w:id.",
                    "Assign the same non-empty move id to the paired move-from and move-to markup.", null);

            foreach (var id in all.Select(item => item.Id)
                         .Where(id => !string.IsNullOrEmpty(id)).Distinct(StringComparer.Ordinal))
            {
                int wrapperFromCount = wrapperFrom.Count(item => item.Id == id);
                int wrapperToCount = wrapperTo.Count(item => item.Id == id);
                var element = all.First(item => item.Id == id).Element;
                if (wrapperFromCount + wrapperToCount > 0
                    && (wrapperFromCount != 1 || wrapperToCount != 1))
                    AddElementFinding(sink, part, element, "structure.move_pair_invalid",
                        VerificationFindingSeverity.Error,
                        $"Tracked move id '{id}' has {wrapperFromCount} source wrapper(s) and {wrapperToCount} destination wrapper(s).",
                        "Repair or remove the incomplete tracked move wrapper pair.", id);

                int fromStartCount = rangeFromStarts.Count(item => item.Id == id);
                int fromEndCount = rangeFromEnds.Count(item => item.Id == id);
                int toStartCount = rangeToStarts.Count(item => item.Id == id);
                int toEndCount = rangeToEnds.Count(item => item.Id == id);
                bool hasRangeMarker = fromStartCount + fromEndCount + toStartCount + toEndCount > 0;
                if (hasRangeMarker && (fromStartCount != 1 || fromEndCount != 1
                                       || toStartCount != 1 || toEndCount != 1))
                    AddElementFinding(sink, part, element, "structure.move_range_pair_invalid",
                        VerificationFindingSeverity.Error,
                        $"Tracked move range id '{id}' has source start/end counts {fromStartCount}/{fromEndCount} and destination start/end counts {toStartCount}/{toEndCount}.",
                        "Pair every move range start/end and retain matching source and destination ranges.", id);
                else if (hasRangeMarker
                         && (XNode.CompareDocumentOrder(
                                 rangeFromStarts.Single(item => item.Id == id).Element,
                                 rangeFromEnds.Single(item => item.Id == id).Element) >= 0
                             || XNode.CompareDocumentOrder(
                                 rangeToStarts.Single(item => item.Id == id).Element,
                                 rangeToEnds.Single(item => item.Id == id).Element) >= 0))
                    AddElementFinding(sink, part, element, "structure.move_range_order_invalid",
                        VerificationFindingSeverity.Error,
                        $"Tracked move range id '{id}' has an end before its start.",
                        "Place each move range end after its matching start.", id);
            }
        }
    }

    private static (XElement Element, string? Id)[] MoveMarkers(
        PackageManifestInspectionEntry part,
        string localName) => Descendants(part, localName)
        .Select(element => (Element: element, Id: WordAttribute(element, "id")))
        .ToArray();

    private static void InspectContentControls(
        WordprocessingInspectionGraph graph,
        IReadOnlyList<PackageManifestInspectionEntry> parts,
        FindingSink sink)
    {
        var storeItems = graph.CustomXmlPropertyParts
            .Select(part => (Part: part,
                Id: CustomXmlAttribute(part.Xml!.Root!, "itemID")))
            .Where(item => !string.IsNullOrEmpty(item.Id))
            .GroupBy(item => NormalizeGuid(item.Id!), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToArray(),
                StringComparer.OrdinalIgnoreCase);
        foreach (var duplicate in storeItems.Where(pair => pair.Value.Length > 1))
        foreach (var item in duplicate.Value)
            AddElementFinding(sink, item.Part, item.Part.Xml!.Root!,
                "structure.custom_xml_store_item_id_duplicate",
                VerificationFindingSeverity.Error,
                $"Custom XML store item id '{item.Id}' is declared by multiple properties parts.",
                "Assign each custom XML item one unique datastore item id.", duplicate.Key);

        foreach (var part in parts)
        {
            var ids = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var sdt in Descendants(part, "sdt"))
            {
                var properties = sdt.Elements().FirstOrDefault(element => IsElement(element, "sdtPr"));
                var content = sdt.Elements().FirstOrDefault(element => IsElement(element, "sdtContent"));
                if (properties is null || content is null)
                    AddElementFinding(sink, part, sdt, "structure.content_control_incomplete",
                        VerificationFindingSeverity.Error,
                        "A content control is missing sdtPr or sdtContent.",
                        "Restore both the content-control properties and content containers.", null);

                if (properties is null) continue;
                var id = properties.Elements().FirstOrDefault(element => IsElement(element, "id")) is { } idElement
                    ? WordAttribute(idElement, "val")
                    : null;
                if (!string.IsNullOrEmpty(id)) ids[id] = ids.GetValueOrDefault(id) + 1;

                if (properties.Elements().Any(element => IsElement(element, "showingPlcHdr")))
                    AddElementFinding(sink, part, sdt, "workflow.content_control_placeholder",
                        VerificationFindingSeverity.Warning,
                        "A content control is still displaying placeholder content.",
                        "Populate the content control and clear w:showingPlcHdr before delivery.", id);

                foreach (var binding in properties.Elements().Where(element => IsElement(element, "dataBinding")))
                {
                    var storeItemId = WordAttribute(binding, "storeItemID");
                    if (string.IsNullOrEmpty(storeItemId)) continue;
                    var normalized = NormalizeGuid(storeItemId);
                    if (!storeItems.TryGetValue(normalized, out var matches))
                        AddElementFinding(sink, part, binding,
                            "structure.content_control_store_item_missing", VerificationFindingSeverity.Error,
                            $"Bound content control references unavailable custom XML store item '{storeItemId}'.",
                            "Restore the customXml item properties or update/remove the data binding.", storeItemId);
                    else if (matches.Length > 1)
                        AddElementFinding(sink, part, binding,
                            "structure.content_control_store_item_ambiguous",
                            VerificationFindingSeverity.Error,
                            $"Bound content control store item '{storeItemId}' has multiple definitions.",
                            "Retain one custom XML properties definition for this store item id.",
                            storeItemId);
                }
            }

            foreach (var duplicate in ids.Where(pair => pair.Value > 1))
            {
                var element = Descendants(part, "sdt").First(sdt =>
                    sdt.Descendants().Any(candidate => IsElement(candidate, "id")
                        && WordAttribute(candidate, "val") == duplicate.Key));
                AddElementFinding(sink, part, element, "structure.content_control_id_duplicate",
                    VerificationFindingSeverity.Error,
                    $"Content-control id '{duplicate.Key}' occurs {duplicate.Value} times in this part.",
                    "Assign each content control a unique native w:id.", duplicate.Key);
            }
        }
    }

    private static void InspectNumbering(
        IReadOnlyList<PackageManifestInspectionEntry> storyParts,
        IReadOnlyList<PackageManifestInspectionEntry> numberingParts,
        FindingSink sink)
    {
        var numberingPart = numberingParts.FirstOrDefault();
        var numberElements = numberingPart is null
            ? Array.Empty<XElement>()
            : DirectChildren(numberingPart, "num").ToArray();
        var abstractElements = numberingPart is null
            ? Array.Empty<XElement>()
            : DirectChildren(numberingPart, "abstractNum").ToArray();
        var nums = numberingPart is null
            ? new Dictionary<string, (string? AbstractId, HashSet<string> OverrideLevels)>(StringComparer.Ordinal)
            : numberElements
                .Where(element => WordAttribute(element, "numId") is not null)
                .GroupBy(element => WordAttribute(element, "numId")!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key,
                    group =>
                    {
                        var first = group.First();
                        return (
                            AbstractId: WordAttribute(first.Elements().FirstOrDefault(element =>
                                IsElement(element, "abstractNumId")), "val"),
                            OverrideLevels: first.Elements().Where(element => IsElement(element, "lvlOverride")
                                    && element.Elements().Any(child => IsElement(child, "lvl")))
                                .Select(element => WordAttribute(element, "ilvl"))
                                .Where(value => value is not null).Select(value => value!)
                                .ToHashSet(StringComparer.Ordinal));
                    },
                    StringComparer.Ordinal);
        var abstracts = numberingPart is null
            ? new Dictionary<string, HashSet<string>>(StringComparer.Ordinal)
            : abstractElements
                .Where(element => WordAttribute(element, "abstractNumId") is not null)
                .GroupBy(element => WordAttribute(element, "abstractNumId")!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key,
                    group => group.First().Elements().Where(element => IsElement(element, "lvl"))
                        .Select(element => WordAttribute(element, "ilvl"))
                        .Where(value => value is not null).Select(value => value!)
                        .ToHashSet(StringComparer.Ordinal), StringComparer.Ordinal);

        if (numberingPart is not null)
        {
            foreach (var element in numberElements
                         .Where(element => string.IsNullOrEmpty(WordAttribute(element, "numId"))))
                AddElementFinding(sink, numberingPart, element,
                    "structure.numbering_num_id_missing", VerificationFindingSeverity.Error,
                    "A numbering instance has no w:numId.",
                    "Assign one unique numbering instance id.", null);
            foreach (var element in abstractElements
                         .Where(element => string.IsNullOrEmpty(WordAttribute(element, "abstractNumId"))))
                AddElementFinding(sink, numberingPart, element,
                    "structure.numbering_abstract_id_missing", VerificationFindingSeverity.Error,
                    "An abstract numbering definition has no w:abstractNumId.",
                    "Assign one unique abstract numbering id.", null);
            foreach (var element in numberElements.Concat(abstractElements)
                         .SelectMany(element => element.DescendantsAndSelf())
                         .Where(element => IsElement(element, "lvl"))
                         .Where(element => string.IsNullOrEmpty(WordAttribute(element, "ilvl"))))
                AddElementFinding(sink, numberingPart, element,
                    "structure.numbering_level_id_missing", VerificationFindingSeverity.Error,
                    "A numbering level has no w:ilvl.",
                    "Assign the level an unambiguous level index.", null);
            ReportDuplicateIds(numberingPart, numberElements, "numId",
                "structure.numbering_num_id_duplicate", sink);
            ReportDuplicateIds(numberingPart, abstractElements, "abstractNumId",
                "structure.numbering_abstract_id_duplicate", sink);
            foreach (var numElement in numberElements)
            {
                var numId = WordAttribute(numElement, "numId");
                var overrides = numElement.Elements().Where(element => IsElement(element, "lvlOverride")).ToArray();
                foreach (var missing in overrides.Where(element =>
                             string.IsNullOrEmpty(WordAttribute(element, "ilvl"))))
                    AddElementFinding(sink, numberingPart, missing,
                        "structure.numbering_override_level_missing", VerificationFindingSeverity.Error,
                        $"Numbering instance '{numId ?? "(missing)"}' has a level override without w:ilvl.",
                        "Assign the override one unambiguous list level or remove it.", numId);
                foreach (var duplicate in overrides.Where(element =>
                             WordAttribute(element, "ilvl") is not null)
                             .GroupBy(element => WordAttribute(element, "ilvl")!, StringComparer.Ordinal)
                             .Where(group => group.Count() > 1))
                    foreach (var element in duplicate)
                        AddElementFinding(sink, numberingPart, element,
                            "structure.numbering_override_level_duplicate", VerificationFindingSeverity.Error,
                            $"Numbering instance '{numId}' has duplicate overrides for level '{duplicate.Key}'.",
                            "Retain at most one level override per numbering instance and level.",
                            numId + ":" + duplicate.Key);
            }
            foreach (var num in nums)
            {
                if (num.Value.AbstractId is not null && abstracts.ContainsKey(num.Value.AbstractId)) continue;
                var element = numberElements
                    .First(candidate => WordAttribute(candidate, "numId") == num.Key);
                AddElementFinding(sink, numberingPart, element,
                    "structure.numbering_abstract_missing", VerificationFindingSeverity.Error,
                    $"Numbering instance '{num.Key}' references missing abstract numbering '{num.Value.AbstractId ?? "(missing)"}'.",
                    "Restore the abstract numbering definition or retarget the numbering instance.", num.Key);
            }
        }

        foreach (var part in storyParts)
        foreach (var numPr in Descendants(part, "numPr"))
        {
            var numId = WordAttribute(numPr.Elements().FirstOrDefault(element => IsElement(element, "numId")), "val");
            var level = WordAttribute(numPr.Elements().FirstOrDefault(element => IsElement(element, "ilvl")), "val") ?? "0";
            if (string.IsNullOrEmpty(numId) || numId == "0") continue;
            if (!nums.TryGetValue(numId, out var num))
            {
                AddElementFinding(sink, part, numPr, "structure.numbering_instance_missing",
                    VerificationFindingSeverity.Error,
                    $"List item references missing numbering instance '{numId}'.",
                    "Restore the w:num definition or remove the paragraph numbering properties.", numId);
            }
            else if (!num.OverrideLevels.Contains(level)
                     && (num.AbstractId is null
                         || !abstracts.TryGetValue(num.AbstractId, out var levels)
                         || !levels.Contains(level)))
            {
                AddElementFinding(sink, part, numPr, "structure.numbering_level_missing",
                    VerificationFindingSeverity.Error,
                    $"List item references unavailable level '{level}' in numbering instance '{numId}'.",
                    "Define the requested list level or retarget the paragraph to an existing level.",
                    numId + ":" + level);
            }
        }
    }

    private static void InspectFields(
        IReadOnlyList<PackageManifestInspectionEntry> parts,
        FindingSink sink)
    {
        foreach (var part in parts)
        {
            var fieldStack = new Stack<bool>();
            foreach (var element in part.Xml!.Descendants())
            {
                if (IsElement(element, "fldSimple")
                    && string.IsNullOrWhiteSpace(WordAttribute(element, "instr")))
                    AddElementFinding(sink, part, element, "structure.field_instruction_missing",
                        VerificationFindingSeverity.Error,
                        "A simple field has no instruction.",
                        "Add a valid w:instr instruction or replace the field with ordinary content.", null);
                if (!IsElement(element, "fldChar")) continue;
                switch (WordAttribute(element, "fldCharType"))
                {
                    case "begin":
                        fieldStack.Push(false);
                        break;
                    case "separate" when fieldStack.Count == 0:
                    case "end" when fieldStack.Count == 0:
                        AddElementFinding(sink, part, element, "structure.field_sequence_invalid",
                            VerificationFindingSeverity.Error,
                            "A field separator/end appears without a matching field begin in this part.",
                            "Repair the complex-field begin/separate/end sequence.", null);
                        break;
                    case "separate" when fieldStack.Peek():
                        AddElementFinding(sink, part, element, "structure.field_sequence_invalid",
                            VerificationFindingSeverity.Error,
                            "A complex field contains more than one separator at the same nesting level.",
                            "Retain at most one separator between each field begin/end pair.", null);
                        break;
                    case "separate":
                        fieldStack.Pop();
                        fieldStack.Push(true);
                        break;
                    case "end":
                        fieldStack.Pop();
                        break;
                }
            }
            if (fieldStack.Count > 0)
                sink.Add("structure.field_end_missing", DeliverableFindingCategory.Structure,
                    VerificationFindingSeverity.Error,
                    $"This part has {fieldStack.Count} unterminated complex field(s).", part.Uri,
                    "Add the missing field end marker(s) or remove incomplete field markup.",
                    new ChangeLocation { EntryUri = part.Uri, PropertyPath = "fields" },
                    subject: fieldStack.Count.ToString(CultureInfo.InvariantCulture));
        }
    }

    private static void InspectStaticRenderRisks(
        WordprocessingInspectionGraph graph,
        IReadOnlyList<PackageManifestInspectionEntry> parts,
        FindingSink sink)
    {
        var riskyElements = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["altChunk"] = "Alternative-format chunks require renderer-specific conversion.",
            ["object"] = "Embedded OLE objects may not render consistently.",
            ["oMath"] = "Office Math rendering support varies by renderer.",
            ["oMathPara"] = "Office Math paragraph rendering support varies by renderer.",
            ["ruby"] = "Ruby annotations may not render consistently.",
            ["ffData"] = "Legacy form fields may not render or remain interactive.",
        };
        foreach (var part in parts)
        foreach (var pair in riskyElements)
        foreach (var element in part.Xml!.Descendants().Where(element => element.Name.LocalName == pair.Key))
            AddElementFinding(sink, part, element, "render.unsupported_content",
                VerificationFindingSeverity.Warning, pair.Value,
                "Inspect this content in the target renderer and attach its structured render diagnostics.",
                pair.Key, DeliverableFindingCategory.Render);

        foreach (var entry in graph.ReachableEntries.Select(item => item.ManifestEntry).Where(entry =>
                     entry.Uri.EndsWith(".wmf", StringComparison.OrdinalIgnoreCase)
                     || entry.Uri.EndsWith(".emf", StringComparison.OrdinalIgnoreCase)
                     || entry.Uri.EndsWith(".svg", StringComparison.OrdinalIgnoreCase)))
            sink.Add("render.vector_media", DeliverableFindingCategory.Render,
                VerificationFindingSeverity.Warning,
                $"Vector media part '{entry.Uri}' needs renderer-specific verification.", entry.Uri,
                "Render the document with the delivery renderer and attach its diagnostics.",
                new ChangeLocation { EntryUri = entry.Uri }, subject: entry.Uri);
    }

    private static void ReportDuplicateIds(
        PackageManifestInspectionEntry part,
        IEnumerable<XElement> elements,
        string attributeName,
        string code,
        FindingSink sink)
    {
        foreach (var duplicate in elements
                     .Where(element => WordAttribute(element, attributeName) is not null)
                     .GroupBy(element => WordAttribute(element, attributeName)!, StringComparer.Ordinal)
                     .Where(group => group.Count() > 1))
        foreach (var element in duplicate)
            AddElementFinding(sink, part, element, code, VerificationFindingSeverity.Error,
                $"{element.Name.LocalName} id '{duplicate.Key}' is duplicated.",
                "Assign unique numbering identifiers and update dependent references.", duplicate.Key);
    }

    private static void InspectRangeNesting(
        PackageManifestInspectionEntry part,
        string startName,
        string endName,
        string code,
        FindingSink sink)
    {
        var open = new Stack<string>();
        var openIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var marker in part.Xml!.Descendants().Where(element =>
                     IsElement(element, startName) || IsElement(element, endName)))
        {
            var id = WordAttribute(marker, "id");
            if (string.IsNullOrEmpty(id)) continue;
            if (IsElement(marker, startName))
            {
                if (openIds.Add(id)) open.Push(id);
                continue;
            }
            if (!openIds.Remove(id)) continue;
            if (open.Count > 0 && open.Peek() == id)
            {
                open.Pop();
                continue;
            }
            AddElementFinding(sink, part, marker, code, VerificationFindingSeverity.Error,
                $"Range id '{id}' crosses another range instead of closing in nested order.",
                "Repair the start/end ordering so ranges are properly nested.", id);
            while (open.Count > 0)
            {
                var popped = open.Pop();
                openIds.Remove(popped);
                if (popped == id) break;
            }
        }
    }

    private static void AddElementFinding(
        FindingSink sink,
        PackageManifestInspectionEntry part,
        XElement element,
        string code,
        VerificationFindingSeverity severity,
        string message,
        string remediation,
        string? subject,
        DeliverableFindingCategory category = DeliverableFindingCategory.Structure) =>
        sink.Add(code, category, severity, message, part.Uri, remediation,
            new ChangeLocation { EntryUri = part.Uri, PropertyPath = ElementPath(element) },
            xpath: ElementPath(element), subject: subject);

    private static IEnumerable<XElement> Descendants(
        PackageManifestInspectionEntry part,
        string localName) => part.Xml!.Descendants()
        .Where(element => IsWord(element) && element.Name.LocalName == localName);

    private static IEnumerable<XElement> DirectChildren(
        PackageManifestInspectionEntry part,
        string localName) => part.Xml!.Root!.Elements()
        .Where(element => IsWord(element) && element.Name.LocalName == localName);

    private static bool IsElement(XElement element, string localName) =>
        IsWord(element) && element.Name.LocalName == localName;

    private static bool IsWord(XElement element) =>
        element.Name.NamespaceName is TransitionalWord or StrictWord;

    private static string? WordAttribute(XElement? element, string localName) => element?
        .Attributes().FirstOrDefault(attribute =>
            attribute.Name.LocalName == localName
            && (attribute.Name.NamespaceName is TransitionalWord or StrictWord or Word2010
                || string.IsNullOrEmpty(attribute.Name.NamespaceName)))?.Value;

    private static string? CustomXmlAttribute(XElement element, string localName) =>
        element.Attribute(XName.Get(localName, element.Name.NamespaceName))?.Value;

    private static string NormalizeGuid(string value) => value.Trim().Trim('{', '}');

    private static string Title(string value) => char.ToUpperInvariant(value[0]) + value[1..];

    private static string ElementPath(XElement element)
    {
        var segments = element.AncestorsAndSelf().Reverse().Select(current =>
        {
            int position = current.Parent?.Elements(current.Name).TakeWhile(candidate => candidate != current).Count() + 1 ?? 1;
            return $"{current.Name.LocalName}[{position.ToString(CultureInfo.InvariantCulture)}]";
        });
        return "/" + string.Join("/", segments);
    }

    private sealed class FindingSink(
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings,
        DeliverableInspectionBudget budget)
    {
        private readonly int _maximum = Math.Max(0, maximumFindings);
        internal bool Truncated { get; private set; } =
            observations.Count >= Math.Max(0, maximumFindings);
        internal string? ForcedDiagnostic { get; private set; }
        private bool FindingLimitReached => Truncated || observations.Count >= _maximum;
        internal bool Stopped => FindingLimitReached || budget.Exhausted || ForcedDiagnostic is not null;
        internal string? Diagnostic => ForcedDiagnostic
            ?? (budget.Exhausted ? "resource budget exceeded: " + budget.ExhaustedResource
                : FindingLimitReached ? "finding limit reached" : null);

        internal void ForceUnavailable(string diagnostic) => ForcedDiagnostic = diagnostic;

        internal void Add(
            string code,
            DeliverableFindingCategory category,
            VerificationFindingSeverity severity,
            string message,
            string owningPart,
            string remediation,
            ChangeLocation? location = null,
            string? xpath = null,
            string? subject = null)
        {
            if (!budget.Step()) return;
            if (observations.Count >= _maximum)
            {
                Truncated = true;
                return;
            }
            observations.Add(DeliverableFindingObservation.Create(
                code, category, severity, message, owningPart, remediation,
                location, xpath: xpath, subjectKey: subject));
        }
    }

    private sealed class StringTupleComparer : IEqualityComparer<(string Uri, string Id)>
    {
        internal static readonly StringTupleComparer Instance = new();

        public bool Equals((string Uri, string Id) left, (string Uri, string Id) right) =>
            string.Equals(left.Uri, right.Uri, StringComparison.OrdinalIgnoreCase)
            && string.Equals(left.Id, right.Id, StringComparison.Ordinal);

        public int GetHashCode((string Uri, string Id) value) => HashCode.Combine(
            StringComparer.OrdinalIgnoreCase.GetHashCode(value.Uri),
            StringComparer.Ordinal.GetHashCode(value.Id));
    }
}
