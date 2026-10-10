using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;

namespace Docxodus;

public sealed partial class DocxSession
{
    internal sealed record SnapshotPartRelationship(
        string RelId, string PartUri, string ContentType, string RelationshipType);

    private IReadOnlyList<SnapshotPartRelationship> CapturePartRelationships(IReadOnlyList<PartSnapshot> parts)
    {
        var uris = new HashSet<string>(parts.Select(p => p.PartUri), StringComparer.Ordinal);
        return _doc!.MainDocumentPart!.Parts
            .Where(p => uris.Contains(p.OpenXmlPart.Uri.ToString()))
            .Select(p => new SnapshotPartRelationship(p.RelationshipId, p.OpenXmlPart.Uri.ToString(),
                p.OpenXmlPart.ContentType, p.OpenXmlPart.RelationshipType))
            .ToArray();
    }

    /// <summary>
    /// Restore optional XML parts at their exact OPC names. AddNewPart preserves a relationship
    /// id but can allocate footnotes2.xml after deleting footnotes.xml, which breaks later XML,
    /// hyperlink and image restores keyed by the original URI. One topology record covers every
    /// snapshot-scoped part; unrelated custom XML parts and their children remain untouched.
    /// </summary>
    private void ReconcileSnapshotParts(DocumentSnapshot snapshot, IReadOnlyDictionary<string, PartSnapshot> byUri)
    {
        var main = _doc!.MainDocumentPart!;
        var liveUris = new HashSet<string>(EnumerateProjectedPartsForSnapshot().Select(p => p.Uri.ToString()),
            StringComparer.Ordinal);
        var live = main.Parts.Where(p => liveUris.Contains(p.OpenXmlPart.Uri.ToString()))
            .Select(p => new SnapshotPartRelationship(p.RelationshipId, p.OpenXmlPart.Uri.ToString(),
                p.OpenXmlPart.ContentType, p.OpenXmlPart.RelationshipType)).ToArray();
        // Preserve relationship order too: it determines header/footer scope numbering.
        if (live.SequenceEqual(snapshot.PartRelationships)) return;

        var expected = snapshot.PartRelationships.ToHashSet();
        foreach (var relationship in live)
            if (!expected.Contains(relationship)) main.DeletePart(relationship.RelId);

        // The surviving trees have already been restored. Flush them before reopening the SDK
        // graph, including numbering/threading parts whose streams are not flushed by Save.
        foreach (var part in EnumerateProjectedPartsForSnapshot()) part.PutXDocument();

        var package = _doc.GetPackage();
        var mainPart = package.GetPart(main.Uri);
        foreach (var relationship in live)
            if (mainPart.RelationshipExists(relationship.RelId)) mainPart.DeleteRelationship(relationship.RelId);
        foreach (var relationship in snapshot.PartRelationships)
        {
            var uri = new Uri(relationship.PartUri, UriKind.Relative);
            if (package.PartExists(uri) && package.GetPart(uri).ContentType != relationship.ContentType)
                throw new InvalidOperationException($"Cannot restore part {uri}: its content type changed.");
            if (!package.PartExists(uri))
            {
                var part = package.CreatePart(uri, relationship.ContentType, CompressionOption.Normal);
                using var output = part.GetStream(FileMode.Create, FileAccess.Write);
                byUri[relationship.PartUri].Materialize().Save(output, SaveOptions.DisableFormatting);
            }
            mainPart.CreateRelationship(PackUriHelper.GetRelativeUri(main.Uri, uri), TargetMode.Internal,
                relationship.RelationshipType, relationship.RelId);
        }
        package.Flush();
        DisposeRenderShell();
        _doc.Dispose();
        _stream!.Position = 0;
        _doc = WordprocessingDocument.Open(_stream, isEditable: true);
    }
}
