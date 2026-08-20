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
}
