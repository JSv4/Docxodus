// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Text.Json;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

public class SemanticChangeSetContractTests
{
    [Fact]
    public void Integer_values_use_the_cross_runtime_safe_range()
    {
        Assert.Equal(SemanticValue.MinSafeInteger,
            SemanticValue.Integer(SemanticValue.MinSafeInteger).IntegerValue);
        Assert.Equal(SemanticValue.MaxSafeInteger,
            SemanticValue.Integer(SemanticValue.MaxSafeInteger).IntegerValue);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            SemanticValue.Integer(SemanticValue.MinSafeInteger - 1));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            SemanticValue.Integer(SemanticValue.MaxSafeInteger + 1));

        var schemaPath = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "../../../../docs/schemas/semantic-changes-v1.schema.json"));
        using var schema = JsonDocument.Parse(File.ReadAllBytes(schemaPath));
        var integer = schema.RootElement.GetProperty("$defs")
            .GetProperty("integerValue")
            .GetProperty("properties")
            .GetProperty("value");
        Assert.Equal(SemanticValue.MinSafeInteger, integer.GetProperty("minimum").GetInt64());
        Assert.Equal(SemanticValue.MaxSafeInteger, integer.GetProperty("maximum").GetInt64());
    }
}
