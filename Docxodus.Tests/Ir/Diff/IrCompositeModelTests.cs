#nullable enable
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

public class IrCompositeModelTests
{
    [Fact]
    public void Composite_op_records_are_value_equal()
    {
        var t = new IrTokenOp(IrTokenOpKind.Insert, 0, 0, 0, 2);
        var a = new IrAuthoredTokenOp(t, "Bob", 0);
        var b = new IrAuthoredTokenOp(t, "Bob", 0);
        Assert.Equal(a, b);

        var op = new IrEditOp(IrEditOpKind.InsertBlock, null, "p:body:x", null, null, null);
        var c1 = new IrCompositeOp(op, "Bob", 0);
        var c2 = new IrCompositeOp(op, "Bob", 0);
        Assert.Equal(c1, c2);
        Assert.Null(c1.AuthoredTokens);
        Assert.Null(c1.ConflictId);
    }

    [Fact]
    public void Composite_script_holds_operations_and_conflicts()
    {
        var op = new IrEditOp(IrEditOpKind.EqualBlock, "p:body:a", "p:body:a", null, null, null);
        var script = new IrCompositeScript(
            IrNodeList.From(new[] { new IrCompositeOp(op, "Bob", 0) }),
            IrNodeList.Empty<IrConflict>());
        Assert.Single(script.Operations);
        Assert.Empty(script.Conflicts);
    }

    [Fact]
    public void Conflict_records_are_value_equal_and_competitors_participate()
    {
        var c1 = new IrConflict(1, "p:body:a", 2, 3, Docxodus.ConflictResolution.BaseWins,
            IrNodeList.From(new[] { new IrConflictCompetitor("Bob", "X"), new IrConflictCompetitor("Fred", "Y") }));
        var c2 = new IrConflict(1, "p:body:a", 2, 3, Docxodus.ConflictResolution.BaseWins,
            IrNodeList.From(new[] { new IrConflictCompetitor("Bob", "X"), new IrConflictCompetitor("Fred", "Y") }));
        Assert.Equal(c1, c2);

        // a different competitor list makes the conflicts unequal (Competitors participates in equality)
        var c3 = c1 with { Competitors = IrNodeList.From(new[] { new IrConflictCompetitor("Bob", "X") }) };
        Assert.NotEqual(c1, c3);

        // a different applied policy makes them unequal
        var c4 = c1 with { AppliedPolicy = Docxodus.ConflictResolution.StackAll };
        Assert.NotEqual(c1, c4);

        // IrCompositeScript NoteOps defaults null
        var script = new IrCompositeScript(IrNodeList.Empty<IrCompositeOp>(), IrNodeList.From(new[] { c1 }));
        Assert.Null(script.NoteOps);
        Assert.Single(script.Conflicts);
    }
}
