// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Docxodus.Internal;

namespace Docxodus.Verification;

/// <summary>Bounded story-part workflow and revision inspection over manifest-owned XML clones.</summary>
internal static class DeliverableSessionInspector
{
    private const string TransitionalWord =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StrictWord = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    private const string InternalAnchorNamespace = "http://powertools.codeplex.com/2011";

    private static readonly (Regex Pattern, string Kind, string Code)[] HighConfidencePatterns =
    {
        (BoundedRegex(@"\{\{[^{}\r\n]{1,256}\}\}"), "double_brace", "workflow.placeholder_remaining"),
        (BoundedRegex(@"\$\{[^{}\r\n]{1,256}\}"), "dollar_brace", "workflow.placeholder_remaining"),
        (BoundedRegex(@"<<[^<>\r\n]{1,256}>>"), "angle_placeholder", "workflow.placeholder_remaining"),
        (BoundedRegex(@"\[(?:_{3,}|\.{3,}|(?:INSERT|ENTER|TYPE)\b[^\]\r\n]{0,192})\]"),
            "explicit_bracket_placeholder", "workflow.placeholder_remaining"),
        (BoundedRegex(@"_{3,}"), "underscore", "workflow.blank_run_remaining"),
    };

    private static readonly Regex AlternativeClause = BoundedRegex(@"\[[^\]\r\n]{1,256}\]");

    internal static DeliverableCheckResult Inspect(
        WordprocessingInspectionGraph graph,
        DeliverableVerificationOptions options,
        ICollection<DeliverableFindingObservation> observations,
        DeliverableInspectionBudget budget)
    {
        int before = observations.Count;
        try
        {
            foreach (var part in graph.StoryParts)
            {
                if (budget.Exhausted || observations.Count >= options.MaxFindings) break;
                InspectPart(part, options, observations, budget);
            }

            if (!budget.Exhausted && observations.Count < options.MaxFindings)
                InspectRevisionRegistry(graph, options, observations, budget);

            bool truncated = budget.Exhausted || observations.Count >= options.MaxFindings;
            return new DeliverableCheckResult
            {
                Check = "workflow_and_revision_registry",
                Status = truncated
                    ? DeliverableCheckStatus.UnavailableEvidence
                    : DeliverableCheckStatus.Completed,
                FindingCount = observations.Count - before,
                Diagnostic = budget.Exhausted
                    ? "resource budget exceeded: " + budget.ExhaustedResource
                    : truncated ? "finding limit reached" : null,
            };
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            Add(observations, options.MaxFindings, DeliverableFindingObservation.Create(
                "structure.workflow_inspection_unavailable",
                DeliverableFindingCategory.Structure,
                VerificationFindingSeverity.Error,
                $"Bounded workflow inspection could not be completed ({exception.GetType().Name}).",
                "/",
                "Repair the implicated story markup or reduce it to fit the detector budget.",
                new ChangeLocation { PropertyPath = "workflowInspection" },
                subjectKey: exception.GetType().FullName));
            return new DeliverableCheckResult
            {
                Check = "workflow_and_revision_registry",
                Status = DeliverableCheckStatus.UnavailableEvidence,
                FindingCount = observations.Count - before,
                Diagnostic = exception.GetType().Name,
            };
        }
    }

    private static void InspectPart(
        PackageManifestInspectionEntry part,
        DeliverableVerificationOptions options,
        ICollection<DeliverableFindingObservation> observations,
        DeliverableInspectionBudget budget)
    {
        var root = part.Xml!.Root!;
        var pending = new Stack<(XElement Element, string Path)>();
        pending.Push((root, "/" + root.Name.LocalName + "[1]"));
        while (pending.Count > 0 && !budget.Exhausted && observations.Count < options.MaxFindings)
        {
            var (element, path) = pending.Pop();
            if (!budget.Node() || !budget.Step()) break;
            if (IsWord(element) && element.Name.LocalName == "p")
                InspectParagraph(part.Uri, element, path, options, observations, budget);

            var children = element.Elements().ToArray();
            var positions = new int[children.Length];
            var counts = new Dictionary<XName, int>();
            for (int index = 0; index < children.Length; index++)
            {
                counts[children[index].Name] = counts.GetValueOrDefault(children[index].Name) + 1;
                positions[index] = counts[children[index].Name];
            }
            for (int index = children.Length - 1; index >= 0; index--)
                pending.Push((children[index], path + "/" + children[index].Name.LocalName + "["
                    + positions[index].ToString(CultureInfo.InvariantCulture) + "]"));
        }
    }

    private static void InspectParagraph(
        string partUri,
        XElement paragraph,
        string path,
        DeliverableVerificationOptions options,
        ICollection<DeliverableFindingObservation> observations,
        DeliverableInspectionBudget budget)
    {
        var builder = new StringBuilder();
        foreach (var element in paragraph.Descendants())
        {
            if (!budget.Node() || !budget.Step()) return;
            if (IsWord(element)
                && element.Name.LocalName is "t" or "delText" or "instrText")
                builder.Append(element.Value);
        }
        var text = builder.ToString();
        if (text.Length == 0 || !budget.Text(text.Length)) return;
        var anchor = paragraph.Attributes().FirstOrDefault(attribute =>
            attribute.Name.NamespaceName == InternalAnchorNamespace
            && attribute.Name.LocalName == "Unid")?.Value;
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var (pattern, kind, code) in HighConfidencePatterns)
            ScanRegex(pattern, kind, code, text, partUri, path, anchor,
                VerificationFindingSeverity.Warning, options, observations, budget, seen);

        foreach (var token in options.PlaceholderTokens.OrderBy(value => value, StringComparer.Ordinal))
        {
            int offset = 0;
            while (!budget.Exhausted && (offset = text.IndexOf(token, offset, StringComparison.Ordinal)) >= 0)
            {
                if (!budget.RegexMatch() || !budget.Step()) return;
                AddTextFinding(observations, options.MaxFindings, partUri, path, anchor,
                    offset, token.Length, "configured_token", "workflow.placeholder_remaining",
                    VerificationFindingSeverity.Warning, token, seen);
                offset += Math.Max(1, token.Length);
            }
        }

        foreach (var marker in options.EditorialMarkers.OrderBy(value => value, StringComparer.Ordinal))
        {
            int offset = 0;
            while (!budget.Exhausted && (offset = text.IndexOf(marker, offset, StringComparison.Ordinal)) >= 0)
            {
                if (!budget.RegexMatch() || !budget.Step()) return;
                if (IsBoundary(text, offset - 1) && IsBoundary(text, offset + marker.Length))
                    AddTextFinding(observations, options.MaxFindings, partUri, path, anchor,
                        offset, marker.Length, "configured_editorial_marker", "workflow.editorial_marker",
                        VerificationFindingSeverity.Warning, marker, seen);
                offset += Math.Max(1, marker.Length);
            }
        }

        if (options.DetectBracketedAlternativeClauses)
            ScanRegex(AlternativeClause, "alternative_clause", "workflow.alternative_clause",
                text, partUri, path, anchor, VerificationFindingSeverity.Warning,
                options, observations, budget, seen);
    }

    private static void ScanRegex(
        Regex regex,
        string kind,
        string code,
        string text,
        string partUri,
        string path,
        string? anchor,
        VerificationFindingSeverity severity,
        DeliverableVerificationOptions options,
        ICollection<DeliverableFindingObservation> observations,
        DeliverableInspectionBudget budget,
        ISet<string> seen)
    {
        for (var match = regex.Match(text); match.Success && !budget.Exhausted; match = match.NextMatch())
        {
            if (!budget.RegexMatch() || !budget.Step()) return;
            AddTextFinding(observations, options.MaxFindings, partUri, path, anchor,
                match.Index, match.Length, kind, code, severity, match.Value, seen);
        }
    }

    private static void AddTextFinding(
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings,
        string partUri,
        string paragraphPath,
        string? anchor,
        int start,
        int length,
        string kind,
        string code,
        VerificationFindingSeverity severity,
        string matchedText,
        ISet<string> seen)
    {
        var occurrence = string.Create(CultureInfo.InvariantCulture, $"{start}:{length}:{code}");
        if (!seen.Add(occurrence)) return;
        var propertyPath = paragraphPath + string.Create(CultureInfo.InvariantCulture,
            $"/text[{start}:{length}]");
        Add(observations, maximumFindings, DeliverableFindingObservation.Create(
            code,
            DeliverableFindingCategory.Workflow,
            severity,
            code == "workflow.blank_run_remaining"
                ? "An unresolved underscore blank remains."
                : code == "workflow.alternative_clause"
                    ? "A bracketed alternative clause remains for optional review."
                    : "High-confidence unresolved template state remains.",
            partUri,
            code == "workflow.alternative_clause"
                ? "Review the alternative clause if this detector was enabled for the workflow."
                : "Replace or intentionally remove the template token before delivery.",
            new ChangeLocation { EntryUri = partUri, PropertyPath = propertyPath },
            anchor,
            Scope(partUri),
            paragraphPath,
            subjectKey: string.Join("\u001f", kind, start.ToString(CultureInfo.InvariantCulture),
                length.ToString(CultureInfo.InvariantCulture),
                DeliverableVerificationIdentity.Token("docxodus.workflow-token.v1", matchedText))));
    }

    private static void InspectRevisionRegistry(
        WordprocessingInspectionGraph graph,
        DeliverableVerificationOptions options,
        ICollection<DeliverableFindingObservation> observations,
        DeliverableInspectionBudget budget)
    {
        var parts = new List<RevisionRegistry.Part>();
        foreach (var part in graph.StoryParts)
        {
            if (part.Xml?.Root is not { } root) continue;
            foreach (var _ in root.DescendantsAndSelf())
            {
                // Account for the registry's second traversal separately from the story scan.
                if (!budget.Node() || !budget.Step()) return;
            }
            parts.Add(new RevisionRegistry.Part(part.Uri, Scope(part.Uri), root));
        }

        var registry = RevisionRegistry.Build(parts);
        foreach (var group in registry.Entries
            .Where(entry => entry.ResolutionStatus != RevisionResolutionStatus.Supported)
            .OrderBy(entry => entry.PartUri, StringComparer.Ordinal)
            .ThenBy(entry => entry.Id, StringComparer.Ordinal))
        {
            if (observations.Count >= options.MaxFindings || !budget.Step()) return;
            var status = group.ResolutionStatus.ToString().ToLowerInvariant();
            var element = group.Units.FirstOrDefault()?.Element
                ?? group.RangeMarkers.FirstOrDefault();
            var path = element is null ? "/revisions/" + group.Id : ElementPath(element);
            var severity = group.ResolutionStatus == RevisionResolutionStatus.Unsupported
                ? VerificationFindingSeverity.Warning
                : VerificationFindingSeverity.Error;
            Add(observations, options.MaxFindings, DeliverableFindingObservation.Create(
                "structure.revision_" + status,
                DeliverableFindingCategory.Structure,
                severity,
                "A tracked revision group is " + status + ".",
                group.PartUri,
                "Repair or remove the implicated tracked revision before delivery.",
                new ChangeLocation { EntryUri = group.PartUri, PropertyPath = path },
                scope: group.Scope,
                xpath: path,
                subjectKey: string.Join("\u001f", group.Id, group.Family,
                    group.ResolutionStatus, group.Diagnostic?.Code ?? string.Empty)));
        }
    }

    private static string ElementPath(XElement element)
    {
        var segments = new Stack<string>();
        for (var current = element; current is not null; current = current.Parent)
        {
            int position = 1;
            for (var sibling = current.PreviousNode; sibling is not null; sibling = sibling.PreviousNode)
            {
                if (sibling is XElement siblingElement && siblingElement.Name == current.Name)
                    position++;
            }
            segments.Push(current.Name.LocalName + "["
                + position.ToString(CultureInfo.InvariantCulture) + "]");
        }
        return "/" + string.Join("/", segments);
    }

    private static bool IsBoundary(string text, int index) => index < 0 || index >= text.Length
        || !(char.IsLetterOrDigit(text[index]) || text[index] == '_');

    private static bool IsWord(XElement element) =>
        element.Name.NamespaceName is TransitionalWord or StrictWord;

    private static string Scope(string partUri) =>
        partUri.Contains("/header", StringComparison.OrdinalIgnoreCase) ? "header"
        : partUri.Contains("/footer", StringComparison.OrdinalIgnoreCase) ? "footer"
        : partUri.Contains("/footnotes", StringComparison.OrdinalIgnoreCase) ? "footnote"
        : partUri.Contains("/endnotes", StringComparison.OrdinalIgnoreCase) ? "endnote"
        : partUri.Contains("/comments", StringComparison.OrdinalIgnoreCase) ? "comment"
        : "body";

    private static Regex BoundedRegex(string pattern) => new(pattern,
        RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture,
        TimeSpan.FromMilliseconds(250));

    private static void Add(
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings,
        DeliverableFindingObservation observation)
    {
        if (observations.Count < maximumFindings) observations.Add(observation);
    }
}
