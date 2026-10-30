using System.Xml.Linq;
using Docxodus.Verification;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Holds a session's live part tree to the depth door's limit for its part as edits change it
/// (issue #1389). The door (<see cref="PartDepthGuard"/>) measures what a session opens; this measures
/// what its edits add, from the tree's own LINQ to XML change events, so no op can add depth unseen.
/// </summary>
/// <remarks>
/// <para>An added element is measured as the levels above it plus its own subtree, so the cost is the
/// content the edit created or moved, never the part. Past the limit the handler throws
/// <see cref="XmlDepthLimitException"/> out of the mutating call; every op runs its mutation inside the
/// shared failure path, which restores the pre-op snapshot, so the edit is refused atomically.</para>
/// <para>One watch per tree, attached when the session snapshots the part before an op. A restored or
/// cloned tree carries no annotation and gets its own watch at the next snapshot. It subscribes after
/// <see cref="PartChangeTracker"/>, so the tracker has recorded a change before a refusal unwinds it.</para>
/// </remarks>
internal sealed class PartDepthWatch
{
    private readonly int _limit;
    private readonly string _partUri;

    private PartDepthWatch(XDocument document, int limit, string partUri)
    {
        _limit = limit;
        _partUri = partUri;
        document.Changed += OnChanged;
    }

    /// <summary>Watch <paramref name="part"/>'s cached tree, once.</summary>
    internal static void Attach(OpenXmlPart part)
    {
        var document = part.GetXDocument();
        if (document.Annotation<PartDepthWatch>() is not null) return;
        document.AddAnnotation(new PartDepthWatch(document, PartDepthGuard.LimitFor(part), part.Uri.OriginalString));
    }

    private void OnChanged(object? sender, XObjectChangeEventArgs e)
    {
        if (e.ObjectChange != XObjectChange.Add || sender is not XElement added) return;
        int above = 0;
        for (var parent = added.Parent; parent is not null; parent = parent.Parent) above++;
        if (above < _limit && !added.HasElements) return;
        try
        {
            XmlSemanticNormalizer.EnsureBoundedDepth(added, above, _limit);
        }
        catch (XmlDepthLimitException)
        {
            throw new XmlDepthLimitException(
                $"xml_depth_limit_exceeded: the edit would nest {_partUri} deeper than the {_limit} level safety limit.")
            { PartUri = _partUri };
        }
    }
}
