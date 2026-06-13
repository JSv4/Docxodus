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
}
