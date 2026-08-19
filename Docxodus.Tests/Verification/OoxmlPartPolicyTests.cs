// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

/// <summary>
/// One content-type-driven part policy for the manifest generator and the semantic detector:
/// Word-owned metadata parts (commentsExtended, people, glossary, VML, customXml properties)
/// are recognized by the spellings Word actually writes, get whitespace-insensitive normalized
/// identities on BOTH surfaces, and contribute their documented facts. Guards issues #512/#513.
/// </summary>
public class OoxmlPartPolicyTests
{
    private const string W = IrTestDocuments.W;
    private const string W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

    [Fact]
    public void Reindenting_word_owned_metadata_parts_is_serialization_only()
    {
        var left = WithCommentsExtended(IrTestDocuments.Create("Alpha"),
            "<w15:commentsEx xmlns:w15=\"" + W15 + "\">" +
            "<w15:commentEx w15:paraId=\"11111111\" w15:done=\"1\"/></w15:commentsEx>");
        var right = RewriteEntry(left, "word/commentsExtended.xml",
            "<w15:commentsEx xmlns:w15=\"" + W15 + "\">\n" +
            "  <w15:commentEx w15:paraId=\"11111111\" w15:done=\"1\"/>\n" +
            "</w15:commentsEx>\n");

        var leftManifest = PackageManifestGenerator.Generate(left.DocumentByteArray);
        var rightManifest = PackageManifestGenerator.Generate(right.DocumentByteArray);
        Assert.Equal(
            leftManifest.NormalizedSemanticDigest.Value,
            rightManifest.NormalizedSemanticDigest.Value);

        var result = SemanticDiff.Compare(left, right);
        Assert.Empty(result.Changes);
    }

    [Fact]
    public void Reindenting_a_vml_part_is_serialization_only()
    {
        const string compact =
            "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\">" +
            "<v:shape id=\"s1\" style=\"width:10pt\"/></xml>";
        const string indented =
            "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\">\n" +
            "  <v:shape id=\"s1\" style=\"width:10pt\"/>\n" +
            "</xml>\n";
        var left = WithHeaderVml(compact);
        var right = RewriteEntry(left, "word/vmlDrawing1.vml", indented);

        var leftManifest = PackageManifestGenerator.Generate(left.DocumentByteArray);
        var rightManifest = PackageManifestGenerator.Generate(right.DocumentByteArray);
        Assert.Equal(
            leftManifest.NormalizedSemanticDigest.Value,
            rightManifest.NormalizedSemanticDigest.Value);

        var result = SemanticDiff.Compare(left, right);
        Assert.Empty(result.Changes);
    }

    [Fact]
    public void A_real_vml_change_still_surfaces()
    {
        const string compact =
            "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\">" +
            "<v:shape id=\"s1\" style=\"width:10pt\"/></xml>";
        var left = WithHeaderVml(compact);
        var right = RewriteEntry(left, "word/vmlDrawing1.vml",
            compact.Replace("width:10pt", "width:99pt"));

        var result = SemanticDiff.Compare(left, right);

        Assert.Contains(result.Changes, change =>
            change.PartUri == "/word/vmlDrawing1.vml");
    }

    [Fact]
    public void Comment_metadata_facts_count_real_word_parts()
    {
        var document = WithPeople(WithCommentsExtended(IrTestDocuments.Create("Alpha"),
            "<w15:commentsEx xmlns:w15=\"" + W15 + "\">" +
            "<w15:commentEx w15:paraId=\"11111111\" w15:done=\"1\"/>" +
            "<w15:commentEx w15:paraId=\"22222222\" w15:paraIdParent=\"11111111\"/>" +
            "</w15:commentsEx>"));

        var manifest = PackageManifestGenerator.Generate(document.DocumentByteArray);

        Assert.Equal(2, manifest.Facts.Annotations.ThreadedCommentMetadata);
        Assert.Equal(1, manifest.Facts.Annotations.CommentReplies);
        Assert.Equal(1, manifest.Facts.Annotations.ResolvedComments);
        Assert.Equal(1, manifest.Facts.Annotations.People);
    }

    [Fact]
    public void Glossary_paragraphs_count_as_story_content()
    {
        var withGlossary = WithGlossary(IrTestDocuments.Create("Body one"),
            $"<w:glossaryDocument xmlns:w=\"{W}\"><w:docParts><w:docPart><w:docPartBody>" +
            "<w:p><w:r><w:t>Stored text</w:t></w:r></w:p>" +
            "</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>");

        var manifest = PackageManifestGenerator.Generate(withGlossary.DocumentByteArray);

        Assert.Equal(2, manifest.Facts.ParagraphCount);
    }

    private static WmlDocument WithCommentsExtended(WmlDocument source, string xml)
        => WithPart<WordprocessingCommentsExPart>(source, xml);

    private static WmlDocument WithPeople(WmlDocument source)
        => WithPart<WordprocessingPeoplePart>(source,
            "<w15:people xmlns:w15=\"" + W15 + "\"><w15:person w15:author=\"Ann\">" +
            "<w15:presenceInfo w15:providerId=\"None\" w15:userId=\"Ann\"/>" +
            "</w15:person></w15:people>");

    private static WmlDocument WithGlossary(WmlDocument source, string xml)
        => WithPart<GlossaryDocumentPart>(source, xml);

    private static WmlDocument WithPart<TPart>(WmlDocument source, string xml)
        where TPart : OpenXmlPart, IFixedContentTypePart
    {
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!.AddNewPart<TPart>();
            WritePartXml(part, xml);
        }
        return new WmlDocument(source.FileName, stream.ToArray());
    }

    private static WmlDocument WithHeaderVml(string vmlXml)
    {
        var basis = IrTestDocuments.WithHeaderAndFooter("Header text", "Footer text");
        using var stream = new MemoryStream();
        stream.Write(basis.DocumentByteArray);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            WriteEntry(archive, "word/vmlDrawing1.vml", vmlXml);
            WriteEntry(archive, "word/_rels/header1.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rIdVml1\" " +
                "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" " +
                "Target=\"vmlDrawing1.vml\"/></Relationships>");

            var typesEntry = archive.GetEntry("[Content_Types].xml")!;
            System.Xml.Linq.XDocument types;
            using (var input = typesEntry.Open())
                types = System.Xml.Linq.XDocument.Load(input);
            var ns = types.Root!.Name.Namespace;
            types.Root.Add(new System.Xml.Linq.XElement(ns + "Default",
                new System.Xml.Linq.XAttribute("Extension", "vml"),
                new System.Xml.Linq.XAttribute("ContentType",
                    "application/vnd.openxmlformats-officedocument.vmlDrawing")));
            using var output = typesEntry.Open();
            output.SetLength(0);
            var bytes = Encoding.UTF8.GetBytes(
                types.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
            output.Write(bytes);
        }
        return new WmlDocument("vml.docx", stream.ToArray());
    }

    private static void WriteEntry(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name);
        using var target = entry.Open();
        var bytes = Encoding.UTF8.GetBytes(content);
        target.Write(bytes);
    }

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var partStream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(partStream, Encoding.UTF8, 1024, leaveOpen: false);
        writer.Write(xml);
    }

    private static WmlDocument RewriteEntry(WmlDocument source, string entryName, string xml)
    {
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry(entryName)
                ?? throw new InvalidOperationException(
                    $"missing entry {entryName}; present: " + string.Join(", ",
                        archive.Entries.Select(item => item.FullName)));
            using var target = entry.Open();
            target.SetLength(0);
            var bytes = Encoding.UTF8.GetBytes(xml);
            target.Write(bytes);
        }
        return new WmlDocument(source.FileName, stream.ToArray());
    }
}
