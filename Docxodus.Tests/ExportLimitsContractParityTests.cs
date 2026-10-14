// <copyright file="ExportLimitsContractParityTests.cs" company="Docxodus">
// Licensed under the MIT License.
// </copyright>

#nullable enable

namespace Docxodus.Tests;

using System;
using System.IO;
using System.Text.Json;
using Docxodus.Verification;
using Xunit;

/// <summary>
/// The npm export-limits contract pins the OPC inspection ceilings as schema-v1 compatibility
/// boundaries, restating values that <see cref="PackageManifestOptions"/> also declares. The two
/// are allowed to be separate declarations, but they are not allowed to disagree: a silent drift
/// would make a browser export inspect a package under different ceilings than a .NET caller.
/// </summary>
public class ExportLimitsContractParityTests
{
    private static JsonElement Defaults()
    {
        var path = Path.Combine("../../../../npm/src/export-resource-limits-v1.json");
        Assert.True(File.Exists(path), $"export-limits contract not found at {path}");
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        return document.RootElement.GetProperty("defaults").Clone();
    }

    [Fact]
    public void EL001_ContractOpcCeilingsMatchTheManifestOptionDefaults()
    {
        var defaults = Defaults();
        var options = new PackageManifestOptions();

        Assert.Equal(options.MaxEntryCount, defaults.GetProperty("opcEntries").GetInt32());
        Assert.Equal(
            options.MaxTotalUncompressedBytes,
            defaults.GetProperty("expandedOpcBytes").GetInt64());
        Assert.Equal(options.MaxXmlPartBytes, defaults.GetProperty("xmlPartBytes").GetInt64());
        Assert.Equal(options.MaxUriLength, defaults.GetProperty("opcUriCharacters").GetInt32());
        Assert.Equal(
            options.MaxCompressionRatio,
            defaults.GetProperty("opcCompressionRatio").GetDouble());
    }

    [Fact]
    public void EL002_ContractCompressedInputMatchesTheWasmSafetyBoundary()
    {
        // The record fixes this at the existing WASM boundary; both must move together.
        Assert.Equal(104_857_600, Defaults().GetProperty("compressedDocxBytes").GetInt64());
    }
}
