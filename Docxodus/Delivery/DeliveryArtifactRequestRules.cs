// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.Delivery;

/// <summary>Shared request-shape rules used by the API, CLI, and MCP adapters.</summary>
public static class DeliveryArtifactRequestRules
{
    public const int MaximumArtifactCount = 1_024;
    public const int MaximumStringLength = 4_096;
    public const long MaximumInputPackageBytes = 100L * 1024 * 1024;

    public static bool IsProfiledRenderKind(DeliveryArtifactKind kind) => kind is
        DeliveryArtifactKind.StandaloneHtml
        or DeliveryArtifactKind.FinalPdf
        or DeliveryArtifactKind.ReviewPdf
        or DeliveryArtifactKind.PageMap
        or DeliveryArtifactKind.RenderReport;

    /// <summary>Validate the profile fields whose meaning is shared by every transport.</summary>
    public static void ValidateProfileSelection(DeliveryArtifactRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        var isRender = IsProfiledRenderKind(request.Kind);
        if (isRender)
        {
            if (request.ReviewProfile is null || !Enum.IsDefined(request.ReviewProfile.Value)
                || request.CommentProfile is null || !Enum.IsDefined(request.CommentProfile.Value))
                throw new ArgumentException(
                    $"Render artifact '{request.ArtifactId}' requires explicit review and comment profiles.");
            if (request.Kind == DeliveryArtifactKind.FinalPdf
                && request.ReviewProfile != DeliveryReviewProfile.Final)
                throw new ArgumentException(
                    $"Final PDF artifact '{request.ArtifactId}' requires the final review profile.");
            if (request.Kind == DeliveryArtifactKind.ReviewPdf
                && request.ReviewProfile != DeliveryReviewProfile.Markup)
                throw new ArgumentException(
                    $"Review PDF artifact '{request.ArtifactId}' requires the markup review profile.");
        }
        else if (request.ReviewProfile is not null || request.CommentProfile is not null)
        {
            throw new ArgumentException(
                $"Non-render artifact '{request.ArtifactId}' cannot select render profiles.");
        }
    }
}
