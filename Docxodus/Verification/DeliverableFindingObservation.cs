// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.Verification;

/// <summary>Detector output before baseline disposition and policy are applied.</summary>
internal sealed record DeliverableFindingObservation
{
    required public string IdentityKey { get; init; }
    required public string OccurrenceKey { get; init; }
    required public string Code { get; init; }
    required public DeliverableFindingCategory Category { get; init; }
    required public VerificationFindingSeverity Severity { get; init; }
    required public string Message { get; init; }
    required public string OwningPartUri { get; init; }
    public ChangeLocation? Location { get; init; }
    public string? AnchorId { get; init; }
    public string? Scope { get; init; }
    public string? XPath { get; init; }
    required public string Remediation { get; init; }

    internal static DeliverableFindingObservation Create(
        string code,
        DeliverableFindingCategory category,
        VerificationFindingSeverity severity,
        string message,
        string owningPartUri,
        string remediation,
        ChangeLocation? location = null,
        string? anchorId = null,
        string? scope = null,
        string? xpath = null,
        string? subjectKey = null,
        string ruleVersion = "1")
    {
        var normalizedOwner = string.IsNullOrWhiteSpace(owningPartUri) ? "/" : owningPartUri;
        var nativeIdentity = string.Join("\u001e", new[]
        {
            "docxodus.deliverable-finding.v2",
            code,
            ruleVersion,
            normalizedOwner,
            DeliverableVerificationIdentity.LocationKey(location),
            scope ?? string.Empty,
            xpath ?? string.Empty,
            anchorId ?? string.Empty,
            subjectKey ?? string.Empty,
        });
        return new DeliverableFindingObservation
        {
            IdentityKey = nativeIdentity,
            OccurrenceKey = nativeIdentity,
            Code = code,
            Category = category,
            Severity = severity,
            Message = message,
            OwningPartUri = normalizedOwner,
            Location = location,
            AnchorId = anchorId,
            Scope = scope,
            XPath = xpath,
            Remediation = remediation,
        };
    }
}

internal sealed record DeliverableInspectionSnapshot
{
    required public PackageManifestInspection Inspection { get; init; }
    required public IReadOnlyList<DeliverableFindingObservation> Observations { get; init; }
    required public IReadOnlyList<DeliverableCheckResult> Checks { get; init; }
    required public bool AnalysisCompleted { get; init; }
}
