// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Globalization;
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

    [Fact]
    public void Document_sourced_integers_degrade_to_strings_instead_of_throwing()
    {
        Assert.Equal(SemanticValueKind.Absent, SemanticValue.IntegerFromDocument(null).Kind);

        var inRange = SemanticValue.IntegerFromDocument(SemanticValue.MaxSafeInteger);
        Assert.Equal(SemanticValueKind.Integer, inRange.Kind);
        Assert.Equal(SemanticValue.MaxSafeInteger, inRange.IntegerValue);

        // A crafted OOXML attribute parses as an unbounded long. Projecting it must not throw, and
        // two distinct out-of-range values must not collapse into one indistinguishable record.
        var above = SemanticValue.IntegerFromDocument(SemanticValue.MaxSafeInteger + 1);
        var farAbove = SemanticValue.IntegerFromDocument(SemanticValue.MaxSafeInteger + 2);
        var below = SemanticValue.IntegerFromDocument(SemanticValue.MinSafeInteger - 1);
        Assert.Equal(SemanticValueKind.String, above.Kind);
        Assert.Equal("9007199254740992", above.StringValue);
        Assert.Equal("9007199254740993", farAbove.StringValue);
        Assert.Equal("-9007199254740992", below.StringValue);
        Assert.Equal(SemanticValueKind.String, below.Kind);

        Assert.Equal(long.MaxValue.ToString(CultureInfo.InvariantCulture),
            SemanticValue.IntegerFromDocument(long.MaxValue).StringValue);
        Assert.Equal(long.MinValue.ToString(CultureInfo.InvariantCulture),
            SemanticValue.IntegerFromDocument(long.MinValue).StringValue);
    }
}
