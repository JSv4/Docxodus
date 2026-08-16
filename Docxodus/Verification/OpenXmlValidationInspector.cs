// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Xml;

namespace Docxodus.Verification;

internal static class OpenXmlValidationInspector
{
    private const string InternalAnchorNamespace = "http://powertools.codeplex.com/2011";

    internal static DeliverableCheckResult Inspect(
        byte[] packageBytes,
        FileFormatVersions version,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        int before = observations.Count;
        int remaining = Math.Max(0, maximumFindings - observations.Count);
        var profile = "fileFormatVersion=" + version;
        try
        {
            using var stream = new MemoryStream(packageBytes, writable: false);
            using var document = WordprocessingDocument.Open(stream, isEditable: false);
            int retainedBoundary = remaining == int.MaxValue ? int.MaxValue : remaining + 1;
            var errors = new OpenXmlValidator(version).Validate(document)
                .Where(error => error.Node?.NamespaceUri != InternalAnchorNamespace)
                .Take(retainedBoundary)
                .ToArray();
            foreach (var error in errors.Take(remaining))
            {
                var partUri = error.Part?.Uri.ToString()
                    ?? error.RelatedPart?.Uri.ToString()
                    ?? "/";
                var xpath = error.Path?.XPath;
                var validatorId = string.IsNullOrWhiteSpace(error.Id)
                    ? error.ErrorType.ToString()
                    : error.Id;
                var code = "openxml." + DeliverableVerificationIdentity.SanitizeCode(validatorId);
                var nodeName = error.Node is null
                    ? string.Empty
                    : $"{{{error.Node.NamespaceUri}}}{error.Node.LocalName}";
                observations.Add(DeliverableFindingObservation.Create(
                    code,
                    DeliverableFindingCategory.OpenXml,
                    VerificationFindingSeverity.Error,
                    $"Open XML validation {error.ErrorType} '{validatorId}' at "
                    + $"'{xpath ?? "/"}' on '{(nodeName.Length == 0 ? "unknown node" : nodeName)}'.",
                    partUri,
                    "Repair the reported Open XML schema or semantic constraint without suppressing the validator error.",
                    new ChangeLocation { EntryUri = partUri, PropertyPath = xpath },
                    xpath: xpath,
                    subjectKey: string.Join("\u001f", error.ErrorType, validatorId, nodeName)));
            }

            if (errors.Length > remaining && observations.Count < maximumFindings)
            {
                observations.Add(DeliverableFindingObservation.Create(
                    "openxml.finding_limit_exceeded",
                    DeliverableFindingCategory.OpenXml,
                    VerificationFindingSeverity.Error,
                    $"Open XML validation exceeded the finding budget; only the first {remaining} errors were retained.",
                    "/",
                    "Repair validation errors in smaller batches or raise the bounded MaxFindings policy.",
                    new ChangeLocation { PropertyPath = "openXmlValidation" },
                    subjectKey: errors.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            }

            bool truncated = errors.Length > remaining;
            return new DeliverableCheckResult
            {
                Check = "open_xml",
                Status = truncated
                    ? DeliverableCheckStatus.UnavailableEvidence
                    : DeliverableCheckStatus.Completed,
                FindingCount = observations.Count - before,
                Diagnostic = truncated ? profile + "; finding limit reached" : profile,
            };
        }
        catch (Exception exception) when (exception is OpenXmlPackageException
            or InvalidDataException or IOException or ArgumentException or FormatException
            or NotSupportedException
            or XmlException)
        {
            if (observations.Count < maximumFindings)
                observations.Add(DeliverableFindingObservation.Create(
                    "openxml.validation_unavailable",
                    DeliverableFindingCategory.OpenXml,
                    VerificationFindingSeverity.Error,
                    $"Open XML validation could not be completed ({exception.GetType().Name}).",
                    "/",
                    "Repair the package so the Open XML SDK can open and validate it.",
                    new ChangeLocation { PropertyPath = "openXmlValidation" },
                    subjectKey: exception.GetType().FullName));
            return new DeliverableCheckResult
            {
                Check = "open_xml",
                Status = DeliverableCheckStatus.UnavailableEvidence,
                FindingCount = observations.Count - before,
                Diagnostic = profile + "; " + exception.GetType().Name,
            };
        }
    }
}
