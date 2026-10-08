// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;

namespace Docxodus;

public sealed partial class DocxSession
{
    private static void SweepOrphanedStoryRelationships(OpenXmlPart owner)
    {
        OwnedPartRelationships.SweepOrphanedHyperlinks(owner, R.id);
        OwnedPartRelationships.SweepOrphanedImages(owner);
    }

    /// <summary>
    /// Remove only relationships that were referenced inside markup removed by
    /// the current revision operation. A whole-part sweep would also delete unrelated orphaned
    /// relationships that pre-dated the edit and violate package-preservation guarantees.
    /// </summary>
    private static void SweepOrphanedStoryRelationships(
        OpenXmlPart owner,
        IReadOnlyCollection<string> candidateIds)
    {
        if (candidateIds.Count == 0) return;
        var candidates = candidateIds as HashSet<string>
            ?? new HashSet<string>(candidateIds, StringComparer.Ordinal);
        var referenced = OwnedPartRelationships.ReferencedRelationshipIds(owner, candidates);

        foreach (var relationship in owner.Parts.ToList())
            if (candidates.Contains(relationship.RelationshipId)
                && !referenced.Contains(relationship.RelationshipId))
                owner.DeletePart(relationship.RelationshipId);
        foreach (var relationship in owner.HyperlinkRelationships.Cast<ReferenceRelationship>()
                     .Concat(owner.DataPartReferenceRelationships).ToList())
            if (candidates.Contains(relationship.Id) && !referenced.Contains(relationship.Id))
                owner.DeleteReferenceRelationship(relationship.Id);
        foreach (var relationship in owner.ExternalRelationships.ToList())
            if (candidates.Contains(relationship.Id) && !referenced.Contains(relationship.Id))
                owner.DeleteExternalRelationship(relationship.Id);
    }

    /// <summary>Drop every image relationship that no attribute in its owning story part names.
    /// Runs on the mutation boundary (<see cref="InvalidateProjectionCache"/>), not on save: only
    /// a mutation can orphan a relationship, and serialization must stay read-only with respect
    /// to package topology.</summary>
    private void SweepOrphanedStoryImageRelationships()
    {
        if (_disposed || _doc is null) return;
        foreach (var owner in OwnedPartRelationships.StoryParts(_doc))
            OwnedPartRelationships.SweepOrphanedImages(owner.Part);
    }

    /// <summary>
    /// The bytes of each live image part, read once and then shared by every snapshot taken while
    /// the part exists (issue #965). Without it, every mutation copied every image's full bytes
    /// into its undo snapshot, so a text edit on a document with large pictures cost the pictures'
    /// size in time and undo memory.
    ///
    /// <para>Keyed by the SDK part object, which is safe because the session never rewrites an
    /// existing image part's stream: inserting or replacing an image creates a new part (or reuses
    /// an identical one), and restoring image topology writes at the package level and then reopens
    /// the document, so every part object, and with it every entry here, is new. The arrays are
    /// shared, so nothing may write into them.</para>
    /// </summary>
    private readonly System.Runtime.CompilerServices.ConditionalWeakTable<OpenXmlPart, byte[]> _imageBytes = new();

    private byte[] SnapshotImageBytes(ImagePart part) =>
        _imageBytes.GetValue(part, static p => OwnedPartRelationships.ReadPartBytes(p));

    /// <summary>Restore image media and owner-local relationship topology after the owning XML
    /// stories have been restored, including the exact OPC target URI. Reopen the SDK graph once
    /// the low-level repair is complete so every subsequent typed read sees the restored parts.</summary>
    private void RestoreImageRelationships(DocumentSnapshot snapshot)
    {
        if (ImageTopologyMatches(snapshot)) return;

        var owners = OwnedPartRelationships.StoryParts(_doc!)
            .ToDictionary(owner => owner.PartUri, owner => owner.Part, StringComparer.Ordinal);
        // Most restored XML lives in the SDK XDocument cache until Save. Flush it before the
        // controlled package reopen or those just-restored trees would be lost.
        foreach (var part in EnumerateProjectedPartsForSnapshot())
            part.PutXDocument(new XDocument(part.GetXDocument()));
        OwnedPartRelationships.RestoreExactImageTopology(_doc!, owners, snapshot.ImageParts,
            snapshot.ImageRelationships, snapshot.LinkedImageRelationships);
        DisposeRenderShell();
        _doc!.Dispose();
        _stream!.Position = 0;
        _doc = WordprocessingDocument.Open(_stream, isEditable: true);
    }

    /// <summary>A text/format/layout-only undo already has the snapshot's binary topology. Avoid
    /// deleting/recreating media and reopening the SDK graph in that overwhelmingly common case.</summary>
    private bool ImageTopologyMatches(DocumentSnapshot snapshot)
    {
        var liveRelationships = new HashSet<(string OwnerPartUri, string RelId, string TargetPartUri)>();
        var liveLinked = new HashSet<(string OwnerPartUri, string RelId, string TargetUri)>();
        var liveParts = new Dictionary<string, ImagePart>(StringComparer.Ordinal);
        foreach (var owner in OwnedPartRelationships.StoryParts(_doc!))
        {
            foreach (var relationship in OwnedPartRelationships.ImageRelationships(owner.Part))
            {
                var targetUri = relationship.Target.Uri.ToString();
                liveRelationships.Add((owner.PartUri, relationship.RelationshipId, targetUri));
                liveParts[targetUri] = relationship.Target;
            }
            foreach (var relationship in OwnedPartRelationships.ExternalImageRelationships(owner.Part))
                liveLinked.Add((owner.PartUri, relationship.Id, relationship.Uri.ToString()));
        }

        if (!liveRelationships.SetEquals(snapshot.ImageRelationships)
            || !liveLinked.SetEquals(snapshot.LinkedImageRelationships)
            || liveParts.Count != snapshot.ImageParts.Count)
            return false;

        foreach (var expected in snapshot.ImageParts)
        {
            if (!liveParts.TryGetValue(expected.PartUri, out var live)
                || !string.Equals(live.ContentType, expected.ContentType, StringComparison.Ordinal))
                return false;
            // A snapshot taken while this part was live shares its bytes array: no need to reread it.
            if (_imageBytes.TryGetValue(live, out var shared) && ReferenceEquals(shared, expected.Bytes))
                continue;
            if (!PartBytesEqual(live, expected.Bytes))
                return false;
        }
        return true;
    }

    private static bool PartBytesEqual(OpenXmlPart part, byte[] expected)
    {
        using var input = part.GetStream(FileMode.Open, FileAccess.Read);
        if (input.CanSeek && input.Length != expected.Length) return false;
        var buffer = new byte[Math.Min(81920, Math.Max(1, expected.Length))];
        int offset = 0;
        while (offset < expected.Length)
        {
            int read = input.Read(buffer, 0, Math.Min(buffer.Length, expected.Length - offset));
            if (read == 0) return false;
            for (int i = 0; i < read; i++)
                if (buffer[i] != expected[offset + i]) return false;
            offset += read;
        }
        return input.ReadByte() == -1;
    }
}
