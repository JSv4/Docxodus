// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using System.Text.Json;
using Docxodus.Verification;

namespace Docxodus.Delivery;

/// <summary>Independent verification of one exact render source/profile cohort.</summary>
public sealed record DeliveryRenderCohortValidation
{
    required public DeliveryReviewProfile ReviewProfile { get; init; }
    required public DeliveryCommentProfile CommentProfile { get; init; }
    required public DeliveryBundleDocumentIdentity SourceDocument { get; init; }
    required public DeliverableVerificationResult Verification { get; init; }
}

/// <summary>
/// Bundle-level validation covering the final DOCX and every exact render cohort. A single
/// deliverable report cannot validate markup/original companions because they are intentionally
/// bound to different package bytes, so this aggregate retains each independently bound run.
/// </summary>
public sealed record DeliveryBundleValidationReport
{
    public const string SchemaId =
        "https://docxodus.dev/schemas/delivery/delivery-bundle-validation/v1";

    public string Schema { get; init; } = SchemaId;
    public int SchemaVersion { get; init; } = 1;
    required public DeliverableVerificationDecision Decision { get; init; }
    required public DeliverableVerificationResult FinalDeliverable { get; init; }
    public IReadOnlyList<DeliveryRenderCohortValidation> RenderCohorts { get; init; } =
        Array.Empty<DeliveryRenderCohortValidation>();

    public byte[] ToCanonicalUtf8Bytes() =>
        JsonSerializer.SerializeToUtf8Bytes(this, DeliveryBundleCanonicalJson.Compact);

    public string ToCanonicalJson() => Encoding.UTF8.GetString(ToCanonicalUtf8Bytes());
}
