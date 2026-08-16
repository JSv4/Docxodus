// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;

namespace Docxodus.Verification;

/// <summary>Uses the existing part-aware projection and revision registry on a private clone.</summary>
internal static class DeliverableSessionInspector
{
    internal static DeliverableCheckResult Inspect(
        byte[] packageBytes,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        int before = observations.Count;
        try
        {
            using var session = new DocxSession(packageBytes.ToArray(), new DocxSessionSettings
            {
                CaptureInitialProjection = false,
                PersistAnchorIds = false,
                UndoDepth = 1,
                UndoMemoryBudgetBytes = 1,
                EmitMarkdownPatch = false,
            });

            var summary = session.GetEditSummary();
            foreach (var placeholder in summary.RemainingPlaceholders)
            {
                if (observations.Count >= maximumFindings) break;
                AddTextFinding(observations, placeholder.Match,
                    "workflow.placeholder_remaining", VerificationFindingSeverity.Warning,
                    $"An unresolved {placeholder.Kind} placeholder remains.",
                    "Replace or intentionally remove the placeholder before delivery.",
                    placeholder.Kind.ToString());
            }
            foreach (var match in summary.BareUnderscoreRuns)
            {
                if (observations.Count >= maximumFindings) break;
                AddTextFinding(observations, match,
                    "workflow.blank_run_remaining", VerificationFindingSeverity.Warning,
                    "An unresolved underscore blank remains.",
                    "Fill or intentionally remove the underscore blank before delivery.",
                    "underscore");
            }

            var workflowPatterns = new[]
            {
                (Pattern: @"\{\{[^{}\r\n]+\}\}", Kind: "double_brace"),
                (Pattern: @"\$\{[^{}\r\n]+\}", Kind: "dollar_brace"),
                (Pattern: @"<<[^<>\r\n]+>>", Kind: "angle_placeholder"),
                (Pattern: @"\b(?:TODO|TBD|FIXME)\b", Kind: "editorial_marker"),
            };
            foreach (var (pattern, kind) in workflowPatterns)
            foreach (var match in session.Grep(pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                         ProjectionScopes.All, contextChars: 0))
            {
                if (observations.Count >= maximumFindings) break;
                AddTextFinding(observations, match,
                    "workflow.placeholder_remaining", VerificationFindingSeverity.Warning,
                    "An unresolved workflow placeholder or editorial marker remains.",
                    "Replace or intentionally remove the marker before delivery.", kind);
            }

            foreach (var revision in session.ListRevisions().Where(revision =>
                         revision.ResolutionStatus != RevisionResolutionStatus.Supported))
            {
                if (observations.Count >= maximumFindings) break;
                var severity = revision.ResolutionStatus == RevisionResolutionStatus.Unsupported
                    ? VerificationFindingSeverity.Warning
                    : VerificationFindingSeverity.Error;
                observations.Add(DeliverableFindingObservation.Create(
                    "structure.revision_" + revision.ResolutionStatus.ToString().ToLowerInvariant(),
                    DeliverableFindingCategory.Structure,
                    severity,
                    revision.Diagnostic?.Message
                        ?? $"A tracked revision has {revision.ResolutionStatus.ToString().ToLowerInvariant()} markup.",
                    revision.PartUri,
                    "Resolve or repair the tracked revision with a Word-compatible editor before delivery.",
                    new ChangeLocation
                    {
                        EntryUri = revision.PartUri,
                        PropertyPath = "revisions/" + revision.Id,
                    },
                    revision.AnchorId,
                    revision.Scope,
                    subjectKey: string.Join("\u001f", revision.Id, revision.Family,
                        revision.ResolutionStatus, revision.Diagnostic?.Code)));
            }

            bool truncated = observations.Count >= maximumFindings;
            return new DeliverableCheckResult
            {
                Check = "workflow_and_revision_registry",
                Status = truncated
                    ? DeliverableCheckStatus.UnavailableEvidence
                    : DeliverableCheckStatus.Completed,
                FindingCount = observations.Count - before,
                Diagnostic = truncated ? "finding limit reached" : null,
            };
        }
        catch (Exception exception) when (exception is InvalidDataException or IOException
            or ArgumentException or FormatException or InvalidOperationException
            or PowerToolsDocumentException
            or XmlException)
        {
            if (observations.Count < maximumFindings)
            {
                observations.Add(DeliverableFindingObservation.Create(
                    "structure.session_inspection_unavailable",
                    DeliverableFindingCategory.Structure,
                    VerificationFindingSeverity.Error,
                    $"Part-aware workflow inspection could not be completed ({exception.GetType().Name}).",
                    "/",
                    "Repair the package so Docxodus can project every Word story part.",
                    new ChangeLocation { PropertyPath = "sessionInspection" },
                    subjectKey: exception.GetType().FullName));
            }
            return new DeliverableCheckResult
            {
                Check = "workflow_and_revision_registry",
                Status = DeliverableCheckStatus.UnavailableEvidence,
                FindingCount = observations.Count - before,
                Diagnostic = exception.GetType().Name,
            };
        }
    }

    private static void AddTextFinding(
        ICollection<DeliverableFindingObservation> observations,
        TextMatch match,
        string code,
        VerificationFindingSeverity severity,
        string message,
        string remediation,
        string kind)
    {
        var target = match.EnclosingAnchor;
        var propertyPath = string.Create(CultureInfo.InvariantCulture,
            $"text[{match.Span.Start}:{match.Span.Length}]");
        observations.Add(DeliverableFindingObservation.Create(
            code,
            DeliverableFindingCategory.Workflow,
            severity,
            message,
            target.PartUri,
            remediation,
            new ChangeLocation { EntryUri = target.PartUri, PropertyPath = propertyPath },
            target.Anchor.Id,
            target.Anchor.Scope,
            subjectKey: string.Join("\u001f", kind, match.Span.Start, match.Span.Length)));
    }
}
