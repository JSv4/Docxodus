// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Packages Word opens without complaint carry zero-length ZIP directory entries and relationship parts
/// for source parts that are not in the package (issue #853). Neither changes what any consumer reads:
/// the package manifest reports them as warnings, so the paginated export's preflight, which requires a
/// valid manifest, no longer refuses the document.
/// </summary>
public class PackageToleranceTests
{
    private const string OrphanRels = "word/_rels/comments.xml.rels";

    [Fact]
    public void Manifest_DirectoryEntriesAndOrphanRelationshipPart_IsValid()
    {
        var manifest = PackageManifestGenerator.Generate(Repack(withDirectories: true, withOrphan: true));

        Assert.True(manifest.IsValid, string.Join("\n", manifest.Findings
            .Where(f => f.Severity == VerificationFindingSeverity.Error).Select(f => f.Code + " " + f.Location?.EntryUri)));
        Assert.Equal("/word/document.xml", manifest.Facts.MainDocumentUri);
    }

    [Fact]
    public void Manifest_OrphanRelationshipPart_IsReportedAsAWarningNamingTheAbsentOwner()
    {
        var manifest = PackageManifestGenerator.Generate(Repack(withDirectories: false, withOrphan: true));

        var finding = Assert.Single(manifest.Findings, f => f.Code == "missing_relationship_owner");
        Assert.Equal(VerificationFindingSeverity.Warning, finding.Severity);
        Assert.Equal("/word/comments.xml", finding.Location?.OwnerUri);
    }

    [Fact]
    public void Manifest_OrphanRelationshipPart_ContributesNoRelationshipsOrTargetFindings()
    {
        var manifest = PackageManifestGenerator.Generate(Repack(withDirectories: false, withOrphan: true));

        Assert.DoesNotContain(manifest.Relationships, r => r.OwnerUri == "/word/comments.xml");
        Assert.DoesNotContain(manifest.Findings, f => f.Code == "missing_target");
    }

    [Fact]
    public void Manifest_OrphanRelationshipPart_StillCountsTowardContentIdentity()
    {
        var with = PackageManifestGenerator.Generate(Repack(withDirectories: false, withOrphan: true));
        var without = PackageManifestGenerator.Generate(Repack(withDirectories: false, withOrphan: false));

        Assert.Contains(with.Entries, e => e.Uri == "/" + OrphanRels);
        Assert.NotEqual(without.OrderedOpcContentDigest, with.OrderedOpcContentDigest);
    }

    [Fact]
    public void Manifest_RelationshipPartWithPresentOwnerAndAbsentTarget_IsStillAnError()
    {
        // The tolerance covers only a .rels part whose owner is absent; a live part's broken
        // relationship remains a package error.
        var bytes = Repack(withDirectories: false, withOrphan: false, extra: (
            "word/_rels/styles.xml.rels",
            "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
            "<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' " +
            "Target='media/absent.png'/></Relationships>"));

        var manifest = PackageManifestGenerator.Generate(bytes);

        Assert.False(manifest.IsValid);
        Assert.Contains(manifest.Findings, f => f.Code == "missing_target"
            && f.Severity == VerificationFindingSeverity.Error);
    }

    [Fact]
    public void DocumentEntryPoints_OpenTheTolerantPackage()
    {
        var bytes = Repack(withDirectories: true, withOrphan: true);

        using var session = new DocxSession(bytes);
        Assert.NotEmpty(session.Save());
        Assert.NotNull(WmlToHtmlConverter.ConvertToHtml(new WmlDocument("in.docx", bytes),
            new WmlToHtmlConverterSettings()));
    }

    /// <summary><c>Blank-wml.docx</c> rewritten with directory entries ahead of its parts and/or a
    /// relationship part for an absent <c>word/comments.xml</c> whose target is absent too.</summary>
    private static byte[] Repack(bool withDirectories, bool withOrphan, (string Name, string Xml)? extra = null)
    {
        var source = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles", "Blank-wml.docx"));
        using var input = new ZipArchive(new MemoryStream(source), ZipArchiveMode.Read);
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            if (withDirectories)
                foreach (var directory in new[] { "word/", "_rels/", "docProps/", "word/_rels/" })
                    archive.CreateEntry(directory);
            foreach (var entry in input.Entries)
            {
                using var from = entry.Open();
                using var to = archive.CreateEntry(entry.FullName).Open();
                from.CopyTo(to);
            }
            if (withOrphan)
                Write(archive, OrphanRels,
                    "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
                    "<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' " +
                    "Target='media/image1.png'/></Relationships>");
            if (extra is { } part)
                Write(archive, part.Name, part.Xml);
        }
        return output.ToArray();
    }

    private static void Write(ZipArchive archive, string name, string xml)
    {
        using var stream = archive.CreateEntry(name).Open();
        stream.Write(Encoding.UTF8.GetBytes(xml));
    }
}
