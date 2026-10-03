// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// One validator error, materialized while its package is still open so no
/// <see cref="ValidationErrorInfo"/> (which holds live part and node references) outlives the package.
/// </summary>
internal sealed record PackageValidationError(
    string Id,
    string? PartUri,
    string? XPath,
    string Description,
    ValidationErrorType ErrorType,
    string? NodePrefix,
    string? NodeLocalName);

/// <summary>
/// The one place tests run <see cref="OpenXmlValidator"/> over a package. Every overload takes the
/// <see cref="FileFormatVersions"/> explicitly: the SDK's parameterless validator targets
/// <see cref="FileFormatVersions.Office2007"/>, and a site migrated from it passes that value rather than
/// inheriting a default that would silently re-target it.
/// </summary>
internal static class PackageValidation
{
    /// <summary>Every validator error for an already-open package (including unsaved in-memory edits).</summary>
    internal static IReadOnlyList<PackageValidationError> Errors(WordprocessingDocument document, FileFormatVersions version) =>
        Materialize(new OpenXmlValidator(version).Validate(document));

    /// <summary>Every validator error for one part of an open package (validates that part only).</summary>
    internal static IReadOnlyList<PackageValidationError> Errors(OpenXmlPart part, FileFormatVersions version) =>
        Materialize(new OpenXmlValidator(version).Validate(part));

    /// <summary>Every validator error for a serialized package.</summary>
    internal static IReadOnlyList<PackageValidationError> Errors(byte[] bytes, FileFormatVersions version)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return Errors(document, version);
    }

    /// <inheritdoc cref="Errors(byte[], FileFormatVersions)"/>
    internal static IReadOnlyList<PackageValidationError> Errors(WmlDocument document, FileFormatVersions version) =>
        Errors(document.DocumentByteArray, version);

    /// <summary>The set of <paramref name="key"/> strings over a package's errors.</summary>
    internal static HashSet<string> ErrorKeys(byte[] bytes, FileFormatVersions version, Func<PackageValidationError, string> key) =>
        Errors(bytes, version).Select(key).ToHashSet();

    /// <inheritdoc cref="ErrorKeys(byte[], FileFormatVersions, Func{PackageValidationError, string})"/>
    internal static HashSet<string> ErrorKeys(WmlDocument document, FileFormatVersions version, Func<PackageValidationError, string> key) =>
        ErrorKeys(document.DocumentByteArray, version, key);

    /// <summary>
    /// Asserts the open package has no validator error that <paramref name="counts"/> accepts (every error
    /// counts when it is null).
    /// </summary>
    internal static void AssertValid(WordprocessingDocument document, FileFormatVersions version,
        Func<PackageValidationError, bool>? counts = null) =>
        AssertNone(Errors(document, version), version, counts);

    /// <summary>Asserts one part of an open package validates with no error <paramref name="counts"/> accepts.</summary>
    internal static void AssertValid(OpenXmlPart part, FileFormatVersions version,
        Func<PackageValidationError, bool>? counts = null) =>
        AssertNone(Errors(part, version), version, counts);

    /// <summary>Asserts the package has no validator error that <paramref name="counts"/> accepts.</summary>
    internal static void AssertValid(byte[] bytes, FileFormatVersions version,
        Func<PackageValidationError, bool>? counts = null) =>
        AssertNone(Errors(bytes, version), version, counts);

    /// <inheritdoc cref="AssertValid(byte[], FileFormatVersions, Func{PackageValidationError, bool}?)"/>
    internal static void AssertValid(WmlDocument document, FileFormatVersions version,
        Func<PackageValidationError, bool>? counts = null) =>
        AssertValid(document.DocumentByteArray, version, counts);

    /// <summary>
    /// Asserts <paramref name="after"/> has no error whose <paramref name="key"/> is absent from
    /// <paramref name="before"/>'s keys (set semantics: a key already present before may repeat freely).
    /// When <paramref name="counts"/> is given, only the errors it accepts take part, on both sides.
    /// </summary>
    internal static void AssertNoNewErrors(byte[] before, byte[] after, FileFormatVersions version,
        Func<PackageValidationError, string> key, Func<PackageValidationError, bool>? counts = null)
    {
        counts ??= static _ => true;
        var baseline = Errors(before, version).Where(counts).Select(key).ToHashSet();
        var introduced = Errors(after, version).Where(counts).Where(error => !baseline.Contains(key(error))).ToList();
        Assert.True(introduced.Count == 0, $"New OOXML validation errors ({version}):\n" + Describe(introduced));
    }

    /// <inheritdoc cref="AssertNoNewErrors(byte[], byte[], FileFormatVersions, Func{PackageValidationError, string}, Func{PackageValidationError, bool}?)"/>
    internal static void AssertNoNewErrors(WmlDocument before, WmlDocument after, FileFormatVersions version,
        Func<PackageValidationError, string> key, Func<PackageValidationError, bool>? counts = null) =>
        AssertNoNewErrors(before.DocumentByteArray, after.DocumentByteArray, version, key, counts);

    /// <summary>Keeps only <see cref="ValidationErrorType.Schema"/> errors (drops semantic and package findings).</summary>
    internal static bool IsSchemaError(PackageValidationError error) => error.ErrorType == ValidationErrorType.Schema;

    /// <summary>One line per error: id, part, XPath and description.</summary>
    internal static string Describe(IEnumerable<PackageValidationError> errors) =>
        string.Join("\n", errors.Select(error =>
            $"{error.Id} [{error.PartUri}] {error.XPath}: {error.Description}"));

    private static void AssertNone(IReadOnlyList<PackageValidationError> errors, FileFormatVersions version,
        Func<PackageValidationError, bool>? counts)
    {
        var counted = counts is null ? errors : errors.Where(counts).ToList();
        Assert.True(counted.Count == 0, $"OOXML validation errors ({version}):\n" + Describe(counted));
    }

    private static IReadOnlyList<PackageValidationError> Materialize(IEnumerable<ValidationErrorInfo> errors) =>
        errors.Select(error => new PackageValidationError(
            error.Id,
            error.Part?.Uri?.ToString(),
            error.Path?.XPath,
            error.Description ?? string.Empty,
            error.ErrorType,
            error.Node?.Prefix,
            error.Node?.LocalName)).ToList();

    /// <summary>
    /// Identity keys that error sets are compared by. Each reproduces, character for character, the string a
    /// test file built before this helper existed; which one a site uses is part of what it checks.
    /// </summary>
    internal static class Keys
    {
        /// <summary><c>Id|PartUri|XPath</c>.</summary>
        internal static string IdPartXPath(PackageValidationError e) => e.Id + "|" + e.PartUri + "|" + e.XPath;

        /// <summary><c>Id@XPath: Description</c>.</summary>
        internal static string IdAtXPathDescription(PackageValidationError e) => $"{e.Id}@{e.XPath}: {e.Description}";

        /// <summary><c>Id|Description|XPath</c>.</summary>
        internal static string IdDescriptionXPath(PackageValidationError e) => $"{e.Id}|{e.Description}|{e.XPath}";

        /// <summary>
        /// <c>Id@PartUri: Description</c> with every quoted number in the description replaced by <c>'#'</c>,
        /// so a legitimate id renumber does not read as a new error.
        /// </summary>
        internal static string IdAtPartNumberNormalizedDescription(PackageValidationError e) =>
            $"{e.Id}@{e.PartUri}: {Regex.Replace(e.Description, "'[0-9]+'", "'#'")}";

        /// <summary><c>PartUri: Description</c>.</summary>
        internal static string PartDescription(PackageValidationError e) => $"{e.PartUri}: {e.Description}";

        /// <summary><c>Id@NodeLocalName</c>.</summary>
        internal static string IdAtNode(PackageValidationError e) => $"{e.Id}@{e.NodeLocalName}";

        /// <summary><c>Description</c> alone.</summary>
        internal static string Description(PackageValidationError e) => e.Description;
    }
}
