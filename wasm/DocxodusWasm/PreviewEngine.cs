// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Linq;
using System.Runtime.Versioning;
using Docxodus;

namespace DocxodusWasm;

/// <summary>
/// The preview path's one-time warm-up, and the invariant <see cref="DocxSessionBridge.OpenPreviewSession"/>
/// holds before handing out the first shadow of a module instance. Same mechanism as
/// <see cref="ComparisonEngine.EnsureWarm"/> (issues #695, #696), on a different allocation-heavy
/// cold path: a preview clones the live package into a second session and then, on its first
/// transaction, serializes every part into a package snapshot. Run cold — interpreted, with the
/// live package, the clone and the caller's listings already on the heap — that snapshot can
/// stop making progress under Mono's conservative root scans, and <c>previewBatch</c> never
/// returns. The seed below walks the identical path on a one-paragraph document that allocates
/// almost nothing, so the caller's real preview runs warm whatever its size.
/// </summary>
[SupportedOSPlatform("browser")]
internal static class PreviewEngine
{
    private static bool warmed;

    /// <summary>Clone, begin, mutate and roll back a seed session once per module instance.</summary>
    /// <returns><c>"ok"</c> on success, or a JSON error object.</returns>
    /// <remarks>Latched even on failure, for the same reason as the comparison warm-up: a
    /// warm-up that throws has still forced the cold path to run, and a caller must not pay a
    /// failing warm-up on every preview.</remarks>
    internal static string EnsureWarm()
    {
        if (warmed)
            return "ok";
        warmed = true;

        try
        {
            using var seed = new DocxSession(ComparisonEngine.BuildSeedDocx("warmup preview"));
            using var shadow = seed.CreateShadowSession();
            using var transaction = shadow.BeginTransaction();
            var anchor = shadow.Project().AnchorIndex.Values
                .First(target => target.Anchor.Kind == "p").Anchor.Id;
            _ = shadow.InsertParagraph(anchor, Position.After, "warmup");
            transaction.Rollback();
            return "ok";
        }
        catch (Exception ex)
        {
            return DocumentConverter.SerializeError(ex.Message, ex.GetType().Name);
        }
    }
}
