// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Xml.Linq;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>Image-specific operations layered on the generic owning-part relationship seam.</summary>
internal static partial class OwnedPartRelationships
{
    internal const string ImageRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    internal readonly record struct OwnedImageRelationship(
        OpenXmlPart Owner, string RelationshipId, ImagePart Target);

    internal static IEnumerable<OwnedImageRelationship> ImageRelationships(OpenXmlPart owner) =>
        owner.Parts.Where(pair => pair.OpenXmlPart is ImagePart)
            .Select(pair => new OwnedImageRelationship(owner, pair.RelationshipId,
                (ImagePart)pair.OpenXmlPart));

    internal static IEnumerable<ExternalRelationship> ExternalImageRelationships(OpenXmlPart owner) =>
        owner.ExternalRelationships.Where(relationship =>
            relationship.RelationshipType == ImageRelationshipType);

    internal static ImagePart? ResolveImagePart(OpenXmlPart owner, string? relationshipId)
    {
        if (string.IsNullOrEmpty(relationshipId)) return null;
        return owner.Parts.FirstOrDefault(pair => pair.RelationshipId == relationshipId)
            .OpenXmlPart as ImagePart;
    }

    internal static byte[] ReadPartBytes(OpenXmlPart part)
    {
        using var input = part.GetStream(FileMode.Open, FileAccess.Read);
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    internal static string ImageContentHash(string contentType, byte[] bytes)
    {
        var contentTypeBytes = System.Text.Encoding.UTF8.GetBytes(contentType);
        var payload = new byte[contentTypeBytes.Length + 1 + bytes.Length];
        Buffer.BlockCopy(contentTypeBytes, 0, payload, 0, contentTypeBytes.Length);
        Buffer.BlockCopy(bytes, 0, payload, contentTypeBytes.Length + 1, bytes.Length);
        return Convert.ToHexString(SHA256.HashData(payload));
    }

    /// <summary>Find an identical image part anywhere in the editable stories, attaching it to
    /// <paramref name="owner"/> when necessary; otherwise create and feed a new owner-valid part.
    /// The returned relationship id is always owned by <paramref name="owner"/>.</summary>
    internal static (ImagePart Part, string RelationshipId, bool Reused) FindOrAddImagePart(
        WordprocessingDocument document, OpenXmlPart owner, byte[] bytes,
        string contentType, ImageBinaryFormat format)
    {
        var wantedHash = ImageContentHash(contentType, bytes);
        foreach (var relationship in ImageRelationships(owner))
        {
            if (relationship.Target.ContentType == contentType
                && ImageContentHash(contentType, ReadPartBytes(relationship.Target)) == wantedHash)
                return (relationship.Target, relationship.RelationshipId, true);
        }

        ImagePart? packageMatch = null;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var candidateOwner in StoryParts(document))
        {
            foreach (var relationship in ImageRelationships(candidateOwner.Part))
            {
                if (!seen.Add(relationship.Target.Uri.ToString())
                    || relationship.Target.ContentType != contentType) continue;
                if (ImageContentHash(contentType, ReadPartBytes(relationship.Target)) == wantedHash)
                {
                    packageMatch = relationship.Target;
                    break;
                }
            }
            if (packageMatch is not null) break;
        }

        if (packageMatch is not null)
        {
            var attached = owner.AddPart(packageMatch);
            return (attached, owner.GetIdOfPart(attached), true);
        }

        var partType = format switch
        {
            ImageBinaryFormat.Png => ImagePartType.Png,
            ImageBinaryFormat.Jpeg => ImagePartType.Jpeg,
            ImageBinaryFormat.Gif => ImagePartType.Gif,
            ImageBinaryFormat.Bmp => ImagePartType.Bmp,
            ImageBinaryFormat.Tiff => ImagePartType.Tiff,
            _ => throw new NotSupportedException($"unsupported image format: {format}"),
        };
        ImagePart created = owner switch
        {
            MainDocumentPart part => part.AddImagePart(partType),
            HeaderPart part => part.AddImagePart(partType),
            FooterPart part => part.AddImagePart(partType),
            FootnotesPart part => part.AddImagePart(partType),
            EndnotesPart part => part.AddImagePart(partType),
            WordprocessingCommentsPart part => part.AddImagePart(partType),
            _ => throw new NotSupportedException($"part cannot own a Word image: {owner.Uri}"),
        };
        using (var input = new MemoryStream(bytes, writable: false)) created.FeedData(input);
        return (created, owner.GetIdOfPart(created), false);
    }

    /// <summary>
    /// The subset of <paramref name="candidates"/> that ANY attribute anywhere in the owner's XML
    /// carries as its value.
    /// </summary>
    /// <remarks>
    /// Deliberately name-blind. Deletion is irreversible, so a whitelist of known reference
    /// attributes ("is it <c>r:embed</c> or <c>r:link</c>?") is the wrong shape: it silently
    /// destroys media named by anything outside the list (VML/OLE variants such as
    /// <c>o:relid</c> or <c>r:href</c>, and any attribute a future Word version invents). The
    /// question a sweep must answer is "is this provably unreferenced?", not "is it referenced in
    /// one of the two ways I know about".
    /// <para>
    /// Value equality without a name test is safe because relationship ids are unique within a
    /// part: an attribute whose value equals the id is either a real reference or a coincidence,
    /// and a coincidence errs toward KEEPING media, which is the recoverable direction.
    /// </para>
    /// <para>
    /// Answering for the whole candidate set in ONE walk — rather than re-walking per
    /// relationship — is what lets the sweep sit on the mutation path: the cost is a single
    /// attribute pass per owner regardless of how many images the document holds, and the
    /// returned set is bounded by the relationship count rather than by document size.
    /// </para>
    /// </remarks>
    internal static HashSet<string> ReferencedRelationshipIds(
        OpenXmlPart owner, IReadOnlyCollection<string> candidates)
    {
        var referenced = new HashSet<string>(StringComparer.Ordinal);
        if (candidates.Count == 0) return referenced;
        var wanted = candidates as HashSet<string> ?? new HashSet<string>(candidates, StringComparer.Ordinal);
        var root = owner.GetXDocument().Root;
        if (root is null) return referenced;
        foreach (var element in root.DescendantsAndSelf())
        {
            foreach (var attribute in element.Attributes())
            {
                if (attribute.IsNamespaceDeclaration) continue;
                if (wanted.Contains(attribute.Value)) referenced.Add(attribute.Value);
            }
            if (referenced.Count == wanted.Count) break;
        }
        return referenced;
    }

    /// <summary>Remove unreferenced embedded-image part relationships and external linked-image
    /// relationships from one owner. Shared package parts remain alive while any other owner still
    /// relates to them; the SDK removes the media part only after its last relationship is gone.</summary>
    /// <remarks>
    /// Returns before reading any XML when the owner holds no image relationship at all. That
    /// short-circuit is load-bearing on the mutation path: it keeps the sweep free for the
    /// overwhelming majority of stories (and of documents), and it guarantees the sweep is never
    /// the thing that materializes a part's <c>XDocument</c>.
    /// </remarks>
    internal static int SweepOrphanedImages(OpenXmlPart owner)
    {
        var embedded = ImageRelationships(owner).ToList();
        var external = owner.ExternalRelationships
            .Where(r => r.RelationshipType == ImageRelationshipType).ToList();
        if (embedded.Count == 0 && external.Count == 0) return 0;

        var candidates = new HashSet<string>(StringComparer.Ordinal);
        foreach (var relationship in embedded) candidates.Add(relationship.RelationshipId);
        foreach (var relationship in external) candidates.Add(relationship.Id);
        var referenced = ReferencedRelationshipIds(owner, candidates);

        int removed = 0;
        foreach (var relationship in embedded)
        {
            if (referenced.Contains(relationship.RelationshipId)) continue;
            owner.DeletePart(relationship.RelationshipId);
            removed++;
        }
        foreach (var relationship in external)
        {
            if (referenced.Contains(relationship.Id)) continue;
            owner.DeleteExternalRelationship(relationship.Id);
            removed++;
        }
        return removed;
    }

    /// <summary>Rebuild the snapshot's image layer at exact OPC part URIs. The high-level SDK
    /// controls relationship ids but allocates a fresh media filename (image2, image3, ...), so
    /// undo/redo topology restoration must use the package abstraction for this one operation.
    /// The caller reopens the SDK graph immediately afterward.</summary>
    internal static void RestoreExactImageTopology(
        WordprocessingDocument document,
        IReadOnlyDictionary<string, OpenXmlPart> owners,
        IReadOnlyList<(string PartUri, string ContentType, byte[] Bytes)> imageParts,
        IReadOnlyList<(string OwnerPartUri, string RelId, string TargetPartUri)> imageRelationships,
        IReadOnlyList<(string OwnerPartUri, string RelId, string TargetUri)> linkedRelationships)
    {
        foreach (var owner in owners.Values)
        {
            foreach (var relationship in ImageRelationships(owner).ToList())
                owner.DeletePart(relationship.RelationshipId);
            foreach (var relationship in ExternalImageRelationships(owner).ToList())
                owner.DeleteExternalRelationship(relationship.Id);
        }

        var package = document.GetPackage();
        foreach (var snapshot in imageParts)
        {
            var uri = new Uri(snapshot.PartUri, UriKind.RelativeOrAbsolute);
            if (package.PartExists(uri) && package.GetPart(uri).ContentType != snapshot.ContentType)
                package.DeletePart(uri);
            var part = package.PartExists(uri)
                ? package.GetPart(uri)
                : package.CreatePart(uri, snapshot.ContentType, CompressionOption.Normal);
            using var output = part.GetStream(FileMode.Create, FileAccess.Write);
            output.Write(snapshot.Bytes, 0, snapshot.Bytes.Length);
        }

        foreach (var relationship in imageRelationships)
        {
            if (!owners.TryGetValue(relationship.OwnerPartUri, out var owner)) continue;
            var ownerPackagePart = package.GetPart(owner.Uri);
            var targetUri = new Uri(relationship.TargetPartUri, UriKind.RelativeOrAbsolute);
            var relativeTarget = PackUriHelper.GetRelativeUri(owner.Uri, targetUri);
            ownerPackagePart.CreateRelationship(relativeTarget, TargetMode.Internal,
                ImageRelationshipType, relationship.RelId);
        }
        foreach (var relationship in linkedRelationships)
        {
            if (!owners.TryGetValue(relationship.OwnerPartUri, out var owner)) continue;
            var ownerPackagePart = package.GetPart(owner.Uri);
            ownerPackagePart.CreateRelationship(
                new Uri(relationship.TargetUri, UriKind.RelativeOrAbsolute), TargetMode.External,
                ImageRelationshipType, relationship.RelId);
        }
        package.Flush();
    }
}
