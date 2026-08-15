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

    /// <summary>Generate canonical manifest JSON for a live session's logical checkpoint.</summary>
    public static string GetPackageManifest(int handle) =>
        SessionRegistry.Get(handle).GetPackageManifest().ToJson();
}
