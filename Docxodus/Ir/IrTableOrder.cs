using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Ir;

/// <summary>Source order inside a table or row with interleaved children. Entries are in preorder.
/// Parent indexes identify individual content controls, even when two controls have identical properties.
/// Member indexes count only the rows or cells delivered by the container.</summary>
internal sealed record IrTableOrder(IrNodeList<IrTableChild> Children, IrNodeList<int> Members, IrContainerControls Summary)
{
    internal IEnumerable<IrTableChild> EnclosingControls(int member) => EnclosingControls(Children[Members[member]]);

    internal IEnumerable<IrTableChild> EnclosingControls(IrTableChild child)
    {
        for (int i = child.Parent; i >= 0; i = Children[i].Parent)
            yield return Children[i];
    }

    internal IEnumerable<IrTableChild> EmptyControls =>
        Children.Where(child => child.Name == W.sdt && child.MemberCount == 0 &&
            (child.Parent < 0 || Children[child.Parent].MemberCount != 0));
}

/// <summary>One child of a table, row or control's content. A control's digest covers its shell when it
/// delivers members and its complete subtree when it delivers none. Source follows the reader's retention
/// setting. End is the exclusive preorder index after this child's descendants.</summary>
internal readonly record struct IrTableChild(
    XName Name, int Parent, int Sibling, int FirstMember, int MemberCount, int End, IrHash Digest,
    IrProvenance Source);
