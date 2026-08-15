// Copyright (c) Microsoft. All rights reserved.
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
        OwnedPartRelationships.SweepOrphanedImages(owner, R.embed, R.link);
    }

    private void SweepOrphanedStoryImageRelationships()
    {
        foreach (var owner in OwnedPartRelationships.StoryParts(_doc!))
            OwnedPartRelationships.SweepOrphanedImages(owner.Part, R.embed, R.link);
    }

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
                || !string.Equals(live.ContentType, expected.ContentType, StringComparison.Ordinal)
                || !PartBytesEqual(live, expected.Bytes))
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
