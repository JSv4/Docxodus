// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.Verification;

namespace Docxodus.Internal;

/// <summary>
/// Shared wire-format facade for package verification. WASM, Python, and MCP route through this
/// owner so stateless bytes and live-session checkpoints expose the same canonical JSON schema.
/// </summary>
internal static class VerificationOps
{
    /// <summary>Generate canonical manifest JSON directly from supplied package bytes.</summary>
    public static string GeneratePackageManifest(byte[] packageBytes) =>
        PackageManifestGenerator.GenerateJson(packageBytes);

    /// <summary>
    /// Generate canonical manifest JSON while applying the caller's effective inspection limits.
    /// Export callers use this overload so a lowered ceiling constrains inspection itself rather
    /// than inspecting with the defaults and rejecting only after expansion has occurred.
    /// </summary>
    public static string GeneratePackageManifest(
        byte[] packageBytes,
        PackageManifestOptions options) =>
        PackageManifestGenerator.GenerateJson(packageBytes, options);

    /// <summary>Generate canonical manifest JSON for a live session's logical checkpoint.</summary>
    public static string GetPackageManifest(int handle) =>
        SessionRegistry.Get(handle).GetPackageManifest().ToJson();

    /// <summary>Run the default deliverable-verification policy on exact supplied bytes.</summary>
    public static string VerifyDeliverable(byte[] packageBytes) =>
        DeliverableVerifier.VerifyDeliverable(packageBytes).ToCanonicalJson();

    /// <summary>
    /// Run the default deliverable-verification policy and classify the exact delivered bytes
    /// relative to exact baseline bytes.
    /// </summary>
    public static string VerifyDeliverable(byte[] baselineBytes, byte[] packageBytes) =>
        DeliverableVerifier.VerifyDeliverable(baselineBytes, packageBytes).ToCanonicalJson();

    /// <summary>
    /// Prove that a redline's generated revisions accept to the intended final and reject to the
    /// baseline without consuming pre-existing review state.
    /// </summary>
    /// <remarks>
    /// Only the canonical proof JSON crosses this boundary. The two rebuilt packages stay inside
    /// the process: every transport routing through this facade is a JSON wire, and base64 of two
    /// further packages would multiply the payload for evidence the proof already carries as
    /// digests and structured divergences. A caller that needs those bytes uses
    /// <see cref="RedlineReversibilityVerifier.Prove"/> directly in-process.
    /// </remarks>
    public static string ProveRedlineReversibility(
        byte[] baselineBytes,
        byte[] intendedFinalBytes,
        byte[] redlineBytes) =>
        RedlineReversibilityVerifier
            .Prove(baselineBytes, intendedFinalBytes, redlineBytes)
            .Proof.ToCanonicalJson();

    /// <summary>
    /// Prove redline reversibility while applying the caller's effective proof limits, so a
    /// lowered ceiling constrains the proof itself rather than rejecting it after the work is done.
    /// </summary>
    public static string ProveRedlineReversibility(
        byte[] baselineBytes,
        byte[] intendedFinalBytes,
        byte[] redlineBytes,
        RedlineReversibilityProofOptions options) =>
        RedlineReversibilityVerifier
            .Prove(baselineBytes, intendedFinalBytes, redlineBytes, options)
            .Proof.ToCanonicalJson();
}
