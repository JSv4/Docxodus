#nullable enable

using System;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

public class IrAlignmentAllocationTests
{
    [Fact]
    public void Large_replacement_alignment_stays_within_allocation_budget()
    {
        var leftText = Enumerable.Range(0, 1000).Select(i =>
            $"Inventory entry {i}: stored sensor equipment requires scheduled inspection every month.").ToArray();
        var rightText = Enumerable.Range(0, 1000).Select(i =>
            $"Updated observatory entry {i}: telescope calibration uses reference signal {i + 7}.").ToArray();
        var options = new IrReaderOptions { RetainSources = false };
        var left = IrReader.Read(IrTestDocuments.Create(leftText), options);
        var right = IrReader.Read(IrTestDocuments.Create(rightText), options);
        var settings = new IrDiffSettings();
        for (int i = 0; i < 2; i++)
            _ = IrBlockAligner.Align(left, right, settings);

        long before = GC.GetAllocatedBytesForCurrentThread();
        var alignment = IrBlockAligner.Align(left, right, settings);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        IrAlignmentAsserts.AssertInvariants(left, right, alignment);
        Assert.True(allocated < 44_000_000, $"Alignment allocated {allocated:N0} bytes; budget is 44,000,000.");
    }
}
