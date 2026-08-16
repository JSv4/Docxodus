// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.Verification;

namespace LegalEval;

/// <summary>
/// Evaluation-boundary limits for caller-supplied packages. The adapter maps these workflow-sized
/// budgets onto #456's single bounded package-manifest inspector.
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
    PackageManifest Inspect(byte[] bytes, string label);
    PackageManifestOptions ManifestOptions { get; }
    long MaximumPackageBytes { get; }
}

public sealed class EvaluationPackageValidator : IEvaluationPackageValidator
{
    private readonly EvaluationPackageLimits _limits;

    public EvaluationPackageValidator(EvaluationPackageLimits? limits = null)
    {
        _limits = limits ?? new EvaluationPackageLimits();
        ManifestOptions = new PackageManifestOptions
        {
            MaxEntryCount = _limits.MaximumEntryCount,
            MaxEntryUncompressedBytes = _limits.MaximumExpandedBytes,
            MaxTotalUncompressedBytes = _limits.MaximumExpandedBytes,
            MaxXmlPartBytes = _limits.MaximumXmlPartBytes,
            MaxCompressionRatio = _limits.MaximumCompressionRatio,
            MaxUriLength = _limits.MaximumEntryNameLength,
        };
    }

    public PackageManifestOptions ManifestOptions { get; }
    public long MaximumPackageBytes => _limits.MaximumPackageBytes;

    public PackageManifest Inspect(byte[] bytes, string label)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        if (bytes.LongLength > _limits.MaximumPackageBytes)
            throw new ScenarioValidationException(
                $"{label} exceeds the {_limits.MaximumPackageBytes}-byte package limit");
        var manifest = PackageManifestGenerator.Generate(bytes, ManifestOptions);
        if (manifest.IsValid) return manifest;

        var errors = manifest.Findings
            .Where(finding => finding.Severity == VerificationFindingSeverity.Error)
            .Select(finding => finding.Code)
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal)
            .ToArray();
        throw new ScenarioValidationException(
            $"{label} failed bounded package-manifest validation: {string.Join(", ", errors)}");
    }
}
