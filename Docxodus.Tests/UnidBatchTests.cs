#nullable enable
using System;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace Docxodus.Tests;

public class UnidBatchTests
{
    [Fact]
    public void BatchedAssignmentPreservesExistingIdsAndCreatesIndependentV4IdsAcrossBatches()
    {
        var root = new XElement(W.p, new XAttribute(PtOpenXml.Unid, "existing-root"),
            Enumerable.Range(0, 2500).Select(i => new XElement(W.r,
                i % 11 == 0 ? new XAttribute(PtOpenXml.Unid, "existing-" + i) : null,
                new XElement(W.t, "Text " + i))));
        UnidHelper.AssignToSelfAndDescendants(root);
        var ids = root.DescendantsAndSelf().Select(e => (string)e.Attribute(PtOpenXml.Unid)!).ToArray();
        Assert.Equal(5001, ids.Length);
        Assert.Equal(ids.Length, ids.Distinct().Count());
        Assert.Equal("existing-root", ids[0]);
        Assert.Equal(229, ids.Count(s => s.StartsWith("existing-", StringComparison.Ordinal)));
        foreach (var id in ids.Where(s => !s.StartsWith("existing-", StringComparison.Ordinal)))
        {
            Assert.True(Guid.TryParseExact(id, "N", out _));
            Assert.Equal('4', id[12]);
            Assert.Contains(id[16], "89ab");
        }
        var before = root.ToString();
        UnidHelper.AssignToSelfAndDescendants(root);
        Assert.Equal(before, root.ToString());
    }
}
