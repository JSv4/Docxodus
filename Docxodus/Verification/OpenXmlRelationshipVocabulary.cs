// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.Verification;

/// <summary>Exact transitional/strict Office relationship vocabulary.</summary>
internal static class OpenXmlRelationshipVocabulary
{
    internal const string TransitionalOfficeNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    internal const string StrictOfficeNamespace =
        "http://purl.oclc.org/ooxml/officeDocument/relationships";
    private const string TransitionalOfficeTypePrefix = TransitionalOfficeNamespace + "/";
    private const string StrictOfficeTypePrefix = StrictOfficeNamespace + "/";

    internal static bool IsOfficeNamespace(string value) =>
        value is TransitionalOfficeNamespace or StrictOfficeNamespace;

    internal static bool IsOfficeType(string value, string localType) =>
        value.Equals(TransitionalOfficeTypePrefix + localType, StringComparison.Ordinal)
        || value.Equals(StrictOfficeTypePrefix + localType, StringComparison.Ordinal);

    internal static bool IsStrictOfficeType(string value) =>
        value.StartsWith(StrictOfficeTypePrefix, StringComparison.Ordinal);
}
