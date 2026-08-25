// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json;
using Docxodus.Delivery;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public sealed class DeliveryPackageDeltaReportTests
{
    [Fact]
    public void DB001_Report_ProjectsSharedPackageDeltaDeterministically()
    {
        var baselineBytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(baselineBytes);
        var anchor = session.Project().AnchorIndex.Values
            .First(value => value.Anchor.Kind == "p" && value.Anchor.Scope == "body").Anchor.Id;
        Assert.True(session.ReplaceText(anchor, "Delivery package delta.").Success);
        var finalBytes = session.Save(persistAnchorIds: false);
        var baseline = PackageManifestGenerator.Generate(baselineBytes);
        var final = PackageManifestGenerator.Generate(finalBytes);

        var first = DeliveryPackageDeltaReport.Create(baseline, final);
        var second = DeliveryPackageDeltaReport.Create(baseline, final);

        Assert.Equal(first.ToCanonicalJson(), second.ToCanonicalJson());
        Assert.Equal(first.ChangeCount, first.Changes.Count);
        Assert.NotEmpty(first.Changes);
        Assert.Contains(first.Changes, change =>
            change.Kind == DeliveryPackageDeltaChangeKind.EntryModified
            && change.Location.EntryUri == "/word/document.xml");
        Assert.Equal(first.Changes.Count,
            first.Changes.Select(change => change.ChangeId)
                .Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(baseline.RawPackageBytesDigest,
            first.BaselineDocument.RawPackageBytesDigest);
        Assert.Equal(final.RawPackageBytesDigest,
            first.FinalDocument.RawPackageBytesDigest);

        using var json = JsonDocument.Parse(first.ToCanonicalUtf8Bytes());
        Assert.Equal(DeliveryPackageDeltaReport.SchemaId,
            json.RootElement.GetProperty("schema").GetString());
        Assert.Equal(first.ChangeCount,
            json.RootElement.GetProperty("changeCount").GetInt32());
    }

    [Fact]
    public void DB002_Report_RejectsInvalidManifestAndDoesNotMutateValidInputs()
    {
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var before = bytes.ToArray();
        var valid = PackageManifestGenerator.Generate(bytes);
        var invalid = PackageManifestGenerator.Generate(new byte[] { 1, 2, 3 });

        var report = DeliveryPackageDeltaReport.Create(valid, valid);

        Assert.Empty(report.Changes);
        Assert.Equal(before, bytes);
        Assert.Throws<ArgumentException>(() =>
            DeliveryPackageDeltaReport.Create(invalid, valid));
        Assert.Throws<ArgumentException>(() =>
            DeliveryPackageDeltaReport.Create(valid, invalid));
    }
}
