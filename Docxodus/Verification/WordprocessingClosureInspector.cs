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
    private const string CustomXmlPropertiesContentType =
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";

    internal static DeliverableCheckResult Inspect(
        PackageManifestInspection inspection,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        int before = observations.Count;
        var sink = new FindingSink(observations, maximumFindings);
        var wordParts = inspection.Entries
            .Where(entry => entry.Xml?.Root is not null && IsWord(entry.Xml.Root))
            .ToArray();

        InspectMediaClosure(inspection.Manifest, sink);
        InspectBookmarks(wordParts, sink);
        InspectComments(wordParts, sink);
        InspectNotes(wordParts, sink, "footnote", "footnoteReference", "/word/footnotes.xml");
        InspectNotes(wordParts, sink, "endnote", "endnoteReference", "/word/endnotes.xml");
        InspectMoves(wordParts, sink);
        InspectContentControls(inspection, wordParts, sink);
        InspectNumbering(wordParts, sink);
        InspectFields(wordParts, sink);
        InspectStaticRenderRisks(inspection, wordParts, sink);

        return new DeliverableCheckResult
        {
            Check = "wordprocessing_closure",
            Status = sink.Truncated
                ? DeliverableCheckStatus.UnavailableEvidence
                : DeliverableCheckStatus.Completed,
            FindingCount = observations.Count - before,
            Diagnostic = sink.Truncated ? "finding limit reached" : null,
        };
    }

    private static void InspectMediaClosure(PackageManifest manifest, FindingSink sink)
    {
        foreach (var relationship in manifest.Relationships.Where(relationship =>
                     relationship.TargetMode == "Internal"
                     && relationship.IsTargetPresent == false
                     && (relationship.Type.EndsWith("/image", StringComparison.OrdinalIgnoreCase)
                         || (relationship.ResolvedTargetUri ?? relationship.Target)
                             .Contains("/media/", StringComparison.OrdinalIgnoreCase))))
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
            var starts = Descendants(part, "bookmarkStart").ToArray();
            var ends = Descendants(part, "bookmarkEnd").ToArray();
            foreach (var start in starts)
            {
                var id = WordAttribute(start, "id");
                var name = WordAttribute(start, "name");
                if (!string.IsNullOrEmpty(name))
                {
                    if (!names.TryGetValue(name, out var values))
                        names.Add(name, values = new());
                    values.Add((part, start));
                }
                if (string.IsNullOrEmpty(id))
                {
                    AddElementFinding(sink, part, start, "structure.bookmark_id_missing",
                        VerificationFindingSeverity.Error,
                        "A bookmark start has no w:id.",
                        "Assign a story-part-unique numeric w:id and add its matching bookmark end.",
                        name);
                    continue;
                }
                int startCount = starts.Count(candidate => WordAttribute(candidate, "id") == id);
                int endCount = ends.Count(candidate => WordAttribute(candidate, "id") == id);
                if (startCount != 1 || endCount != 1)
                {
                    AddElementFinding(sink, part, start, "structure.bookmark_pair_invalid",
                        VerificationFindingSeverity.Error,
                        $"Bookmark id '{id}' has {startCount} start marker(s) and {endCount} end marker(s) in this part.",
                        "Use exactly one bookmark start and one bookmark end for each id within its story part.",
                        id);
                }
            }
            foreach (var end in ends.Where(end =>
                         string.IsNullOrEmpty(WordAttribute(end, "id"))
                         || starts.All(start => WordAttribute(start, "id") != WordAttribute(end, "id"))))
            {
                var id = WordAttribute(end, "id");
                AddElementFinding(sink, part, end, "structure.bookmark_start_missing",
                    VerificationFindingSeverity.Error,
                    $"Bookmark end id '{id ?? "(missing)"}' has no start in this part.",
                    "Remove the orphan end or restore its matching bookmark start.", id);
            }
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
        IReadOnlyList<PackageManifestInspectionEntry> parts,
        FindingSink sink)
    {
        var commentParts = parts.Where(part => IsElement(part.Xml!.Root!, "comments")).ToArray();
        var definitions = commentParts.SelectMany(part => Descendants(part, "comment")
                .Select(element => (Part: part, Element: element, Id: WordAttribute(element, "id"))))
            .Where(item => item.Id is not null)
            .ToArray();
        var references = parts.SelectMany(part =>
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
            if (starts == ends && marks > 0) continue;
            var item = group.First();
            AddElementFinding(sink, item.Part, item.Element, "structure.comment_markers_invalid",
                VerificationFindingSeverity.Error,
                $"Comment id '{group.Key.Id}' in '{group.Key.Uri}' has {starts} range start(s), {ends} range end(s), and {marks} reference mark(s).",
                "Pair comment range markers within each story part and retain at least one comment reference mark.",
                group.Key.Id);
        }
    }

    private static void InspectNotes(
        IReadOnlyList<PackageManifestInspectionEntry> parts,
        FindingSink sink,
        string definitionName,
        string referenceName,
        string conventionalPartUri)
    {
        var definitions = parts.SelectMany(part => Descendants(part, definitionName)
                .Select(element => (Part: part, Element: element, Id: WordAttribute(element, "id"),
                    Type: WordAttribute(element, "type"))))
            .Where(item => item.Type is not ("separator" or "continuationSeparator"))
            .Where(item => item.Id is not null)
            .ToArray();
        var references = parts.SelectMany(part => Descendants(part, referenceName)
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
                if (wrapperFromCount != wrapperToCount)
                    AddElementFinding(sink, part, element, "structure.move_pair_invalid",
                        VerificationFindingSeverity.Error,
                        $"Tracked move id '{id}' has {wrapperFromCount} source wrapper(s) and {wrapperToCount} destination wrapper(s).",
                        "Repair or remove the incomplete tracked move wrapper pair.", id);

                int fromStartCount = rangeFromStarts.Count(item => item.Id == id);
                int fromEndCount = rangeFromEnds.Count(item => item.Id == id);
                int toStartCount = rangeToStarts.Count(item => item.Id == id);
                int toEndCount = rangeToEnds.Count(item => item.Id == id);
                bool hasRangeMarker = fromStartCount + fromEndCount + toStartCount + toEndCount > 0;
                if (hasRangeMarker && (fromStartCount != fromEndCount
                                       || toStartCount != toEndCount
                                       || fromStartCount != toStartCount))
                    AddElementFinding(sink, part, element, "structure.move_range_pair_invalid",
                        VerificationFindingSeverity.Error,
                        $"Tracked move range id '{id}' has source start/end counts {fromStartCount}/{fromEndCount} and destination start/end counts {toStartCount}/{toEndCount}.",
                        "Pair every move range start/end and retain matching source and destination ranges.", id);
            }
        }
    }

    private static (XElement Element, string? Id)[] MoveMarkers(
        PackageManifestInspectionEntry part,
        string localName) => Descendants(part, localName)
        .Select(element => (Element: element, Id: WordAttribute(element, "id")))
        .ToArray();

    private static void InspectContentControls(
        PackageManifestInspection inspection,
        IReadOnlyList<PackageManifestInspectionEntry> parts,
        FindingSink sink)
    {
        var storeItemIds = inspection.Entries
            .Where(entry => string.Equals(entry.ManifestEntry.ContentType,
                CustomXmlPropertiesContentType, StringComparison.OrdinalIgnoreCase))
            .SelectMany(entry => entry.Xml?.Root?.DescendantsAndSelf() ?? Enumerable.Empty<XElement>())
            .SelectMany(element => element.Attributes())
            .Where(attribute => attribute.Name.LocalName.Equals("itemID", StringComparison.OrdinalIgnoreCase))
            .Select(attribute => NormalizeGuid(attribute.Value))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

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
                    if (!string.IsNullOrEmpty(storeItemId)
                        && !storeItemIds.Contains(NormalizeGuid(storeItemId)))
                        AddElementFinding(sink, part, binding,
                            "structure.content_control_store_item_missing", VerificationFindingSeverity.Error,
                            $"Bound content control references unavailable custom XML store item '{storeItemId}'.",
                            "Restore the customXml item properties or update/remove the data binding.", storeItemId);
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
        IReadOnlyList<PackageManifestInspectionEntry> parts,
        FindingSink sink)
    {
        var numberingPart = parts.FirstOrDefault(part => IsElement(part.Xml!.Root!, "numbering"));
        var nums = numberingPart is null
            ? new Dictionary<string, string?>(StringComparer.Ordinal)
            : Descendants(numberingPart, "num")
                .Where(element => WordAttribute(element, "numId") is not null)
                .GroupBy(element => WordAttribute(element, "numId")!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key,
                    group => WordAttribute(group.First().Elements()
                        .FirstOrDefault(element => IsElement(element, "abstractNumId")), "val"),
                    StringComparer.Ordinal);
        var abstracts = numberingPart is null
            ? new Dictionary<string, HashSet<string>>(StringComparer.Ordinal)
            : Descendants(numberingPart, "abstractNum")
                .Where(element => WordAttribute(element, "abstractNumId") is not null)
                .GroupBy(element => WordAttribute(element, "abstractNumId")!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key,
                    group => group.First().Elements().Where(element => IsElement(element, "lvl"))
                        .Select(element => WordAttribute(element, "ilvl"))
                        .Where(value => value is not null).Select(value => value!)
                        .ToHashSet(StringComparer.Ordinal), StringComparer.Ordinal);

        if (numberingPart is not null)
        {
            ReportDuplicateIds(numberingPart, "num", "numId", "structure.numbering_num_id_duplicate", sink);
            ReportDuplicateIds(numberingPart, "abstractNum", "abstractNumId",
                "structure.numbering_abstract_id_duplicate", sink);
            foreach (var num in nums)
            {
                if (num.Value is not null && abstracts.ContainsKey(num.Value)) continue;
                var element = Descendants(numberingPart, "num")
                    .First(candidate => WordAttribute(candidate, "numId") == num.Key);
                AddElementFinding(sink, numberingPart, element,
                    "structure.numbering_abstract_missing", VerificationFindingSeverity.Error,
                    $"Numbering instance '{num.Key}' references missing abstract numbering '{num.Value ?? "(missing)"}'.",
                    "Restore the abstract numbering definition or retarget the numbering instance.", num.Key);
            }
        }

        foreach (var part in parts)
        foreach (var numPr in Descendants(part, "numPr"))
        {
            var numId = WordAttribute(numPr.Elements().FirstOrDefault(element => IsElement(element, "numId")), "val");
            var level = WordAttribute(numPr.Elements().FirstOrDefault(element => IsElement(element, "ilvl")), "val") ?? "0";
            if (string.IsNullOrEmpty(numId) || numId == "0") continue;
            if (!nums.TryGetValue(numId, out var abstractId))
            {
                AddElementFinding(sink, part, numPr, "structure.numbering_instance_missing",
                    VerificationFindingSeverity.Error,
                    $"List item references missing numbering instance '{numId}'.",
                    "Restore the w:num definition or remove the paragraph numbering properties.", numId);
            }
            else if (abstractId is null || !abstracts.TryGetValue(abstractId, out var levels)
                     || !levels.Contains(level))
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
        PackageManifestInspection inspection,
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

        foreach (var entry in inspection.Manifest.Entries.Where(entry =>
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
        string elementName,
        string attributeName,
        string code,
        FindingSink sink)
    {
        foreach (var duplicate in Descendants(part, elementName)
                     .Where(element => WordAttribute(element, attributeName) is not null)
                     .GroupBy(element => WordAttribute(element, attributeName)!, StringComparer.Ordinal)
                     .Where(group => group.Count() > 1))
        foreach (var element in duplicate)
            AddElementFinding(sink, part, element, code, VerificationFindingSeverity.Error,
                $"{elementName} id '{duplicate.Key}' is duplicated.",
                "Assign unique numbering identifiers and update dependent references.", duplicate.Key);
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

    private static bool IsElement(XElement element, string localName) =>
        IsWord(element) && element.Name.LocalName == localName;

    private static bool IsWord(XElement element) =>
        element.Name.NamespaceName is TransitionalWord or StrictWord;

    private static string? WordAttribute(XElement? element, string localName) => element?
        .Attributes().FirstOrDefault(attribute =>
            attribute.Name.LocalName == localName
            && (attribute.Name.NamespaceName is TransitionalWord or StrictWord or Word2010
                || string.IsNullOrEmpty(attribute.Name.NamespaceName)))?.Value;

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
        int maximumFindings)
    {
        private readonly int _maximum = Math.Max(0, maximumFindings);
        internal bool Truncated { get; private set; }

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
