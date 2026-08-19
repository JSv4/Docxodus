// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

/// <summary>
/// A session that made no edits must report no semantic changes, whatever the opening package
/// looked like: the baseline flows through the same checkpoint serialization as the current
/// side, so SDK clone normalization (dropped orphan parts, stray entries) is never reported as
/// a document change and never fails only one side's preflight.
/// </summary>
public class DocxSessionSemanticBaselineTests
{
    [Fact]
    public void Unedited_session_with_an_orphan_part_reports_no_semantic_changes()
    {
        var bytes = WithOrphanPart(IrTestDocuments.Create("Alpha").DocumentByteArray);

        using var session = new DocxSession(bytes);

        Assert.Empty(session.GetSemanticChanges().Changes);
    }

    [Fact]
    public void Unedited_session_with_a_stray_content_typeless_entry_reports_no_semantic_changes()
    {
        var bytes = WithStrayEntry(IrTestDocuments.Create("Alpha").DocumentByteArray);

        using var session = new DocxSession(bytes);

        Assert.Empty(session.GetSemanticChanges().Changes);
    }

    [Fact]
    public void Edits_on_an_orphan_part_package_still_report_normally()
    {
        var bytes = WithOrphanPart(
            IrTestDocuments.Create("Alpha", "Delete me").DocumentByteArray);
        using var session = new DocxSession(bytes);
        var target = session.Project().AnchorIndex.Values
            .First(anchor => anchor.TextPreview.Contains("Delete", StringComparison.Ordinal));
        Assert.True(session.DeleteBlock(target.Anchor.Id).Success);

        var result = session.GetSemanticChanges();

        Assert.Contains(result.Changes, change =>
            change.Operation == SemanticChangeOperation.Delete);
        Assert.DoesNotContain(result.Changes, change =>
            change.PartUri.Contains("orphan", StringComparison.Ordinal));
    }

    /// <summary>A valid declared part (content-type override) no relationship reaches.</summary>
    private static byte[] WithOrphanPart(byte[] package)
    {
        using var stream = new MemoryStream();
        stream.Write(package);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.CreateEntry("customXml/orphan.xml");
            using (var target = entry.Open())
            {
                var xml = Encoding.UTF8.GetBytes("<data xmlns=\"urn:example:orphan\"/>");
                target.Write(xml);
            }

            var typesEntry = archive.GetEntry("[Content_Types].xml")!;
            XDocument types;
            using (var input = typesEntry.Open())
                types = XDocument.Load(input);
            var ns = types.Root!.Name.Namespace;
            types.Root.Add(new XElement(ns + "Override",
                new XAttribute("PartName", "/customXml/orphan.xml"),
                new XAttribute("ContentType", "application/xml")));
            using var output = typesEntry.Open();
            output.SetLength(0);
            var bytes = Encoding.UTF8.GetBytes(types.ToString(SaveOptions.DisableFormatting));
            output.Write(bytes);
        }
        return stream.ToArray();
    }

    /// <summary>A ZIP entry with no content-type mapping at all — tolerated by the SDK open,
    /// rejected by manifest preflight.</summary>
    private static byte[] WithStrayEntry(byte[] package)
    {
        using var stream = new MemoryStream();
        stream.Write(package);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.CreateEntry("word/stray.bin");
            using var target = entry.Open();
            var payload = Encoding.UTF8.GetBytes("opaque payload");
            target.Write(payload);
        }
        return stream.ToArray();
    }
}
