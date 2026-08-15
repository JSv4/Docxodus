// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Collections.Generic;
using System.Xml.Linq;

namespace Docxodus.Verification;

/// <summary>
/// Internal, handle-free view of data materialized while generating a package manifest. Parsed XML
/// trees are detached and exclusively owned by the inspection consumer, which may mutate them.
/// Consumers must reject <see cref="Manifest"/> when it is invalid before using partial entries.
/// </summary>
internal sealed record PackageManifestInspection(
    PackageManifest Manifest,
    IReadOnlyList<PackageManifestInspectionEntry> Entries);

/// <summary>One inspected entry from the same pass that produced its manifest record.</summary>
internal sealed record PackageManifestInspectionEntry(
    PackageManifestEntry ManifestEntry,
    XDocument? Xml)
{
    public string Uri => ManifestEntry.Uri;
    public int Occurrence => ManifestEntry.Occurrence;
    public bool PayloadWasRead => ManifestEntry.RawBytesDigest is not null;
}
