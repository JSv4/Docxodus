// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO.Compression;
using System.Xml;

namespace LegalEval;

/// <summary>
/// Evaluation-boundary limits for caller-supplied packages. This intentionally remains a small
/// adapter seam: issue #456's package validator will replace the interim ZIP walk once its public
/// contract lands, rather than leaving a second permanent package policy in the eval runner.
/// </summary>
public sealed record EvaluationPackageLimits(
    long MaximumPackageBytes = 64L * 1024 * 1024,
    int MaximumEntryCount = 4096,
    long MaximumExpandedBytes = 256L * 1024 * 1024,
    long MaximumXmlPartBytes = 32L * 1024 * 1024,
    double MaximumCompressionRatio = 200,
    int MaximumEntryNameLength = 1024);

public interface IEvaluationPackageValidator
{
    void Validate(byte[] bytes, string label);
}

public sealed class InterimEvaluationPackageValidator : IEvaluationPackageValidator
{
    private readonly EvaluationPackageLimits _limits;

    public InterimEvaluationPackageValidator(EvaluationPackageLimits? limits = null) =>
        _limits = limits ?? new EvaluationPackageLimits();

    public void Validate(byte[] bytes, string label)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        if (bytes.LongLength > _limits.MaximumPackageBytes)
            throw new ScenarioValidationException(
                $"{label} exceeds the {_limits.MaximumPackageBytes}-byte package limit");

        try
        {
            using var archive = new ZipArchive(
                new MemoryStream(bytes, writable: false), ZipArchiveMode.Read, leaveOpen: false);
            if (archive.Entries.Count > _limits.MaximumEntryCount)
                throw new ScenarioValidationException(
                    $"{label} contains {archive.Entries.Count} entries; maximum is {_limits.MaximumEntryCount}");

            var names = new HashSet<string>(StringComparer.Ordinal);
            long expandedBytes = 0;
            foreach (var entry in archive.Entries)
            {
                ValidateName(entry.FullName, label);
                if (!names.Add(entry.FullName))
                    throw new ScenarioValidationException(
                        $"{label} contains duplicate ZIP entry '{entry.FullName}'");

                expandedBytes = checked(expandedBytes + entry.Length);
                if (expandedBytes > _limits.MaximumExpandedBytes)
                    throw new ScenarioValidationException(
                        $"{label} exceeds the {_limits.MaximumExpandedBytes}-byte expanded package limit");
                if (entry.CompressedLength > 0
                    && entry.Length / (double)entry.CompressedLength > _limits.MaximumCompressionRatio)
                    throw new ScenarioValidationException(
                        $"{label} entry '{entry.FullName}' exceeds the compression-ratio limit");

                if (!IsXml(entry.FullName)) continue;
                if (entry.Length > _limits.MaximumXmlPartBytes)
                    throw new ScenarioValidationException(
                        $"{label} XML entry '{entry.FullName}' exceeds the XML-part limit");
                using var stream = entry.Open();
                using var reader = XmlReader.Create(stream, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    MaxCharactersInDocument = _limits.MaximumXmlPartBytes,
                    MaxCharactersFromEntities = 0,
                });
                while (reader.Read()) { }
            }
        }
        catch (ScenarioValidationException) { throw; }
        catch (Exception exception) when (exception is InvalidDataException
            or IOException or XmlException or OverflowException)
        {
            throw new ScenarioValidationException(
                $"{label} is not a safe, readable OPC package: {exception.Message}");
        }
    }

    private void ValidateName(string name, string label)
    {
        if (string.IsNullOrEmpty(name) || name.Length > _limits.MaximumEntryNameLength)
            throw new ScenarioValidationException($"{label} contains an invalid ZIP entry name");
        if (name.StartsWith("/", StringComparison.Ordinal)
            || name.StartsWith('\\')
            || Path.IsPathRooted(name)
            || name.Split('/', '\\').Any(segment => segment is ".." or "."))
            throw new ScenarioValidationException(
                $"{label} contains unsafe ZIP entry path '{name}'");
    }

    private static bool IsXml(string name) =>
        name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
        || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase);
}
