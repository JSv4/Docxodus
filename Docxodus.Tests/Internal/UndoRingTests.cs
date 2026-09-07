#nullable enable

// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests.Internal;

public class UndoRingTests
{
    [Fact]
    public void DS001a_RecordAndUndo()
    {
        var r = new UndoRing<string>(3);
        r.RecordPreOp("v0");
        var (snap, ok) = r.PopForUndo();
        Assert.True(ok);
        Assert.Equal("v0", snap);
    }

    [Fact]
    public void DS001b_CapacityEvictsOldest()
    {
        var r = new UndoRing<string>(2);
        r.RecordPreOp("v0");
        r.RecordPreOp("v1");
        r.RecordPreOp("v2");
        Assert.Equal(2, r.UndoCount);
        var (snap, _) = r.PopForUndo();
        Assert.Equal("v2", snap);
        (snap, _) = r.PopForUndo();
        Assert.Equal("v1", snap);
    }

    [Fact]
    public void DS001c_RecordClearsRedo()
    {
        var r = new UndoRing<string>(5);
        r.RecordPreOp("v0");
        _ = r.PopForUndo();
        r.RecordForRedo("post0");
        Assert.Equal(1, r.RedoCount);

        r.RecordPreOp("v1");
        Assert.Equal(0, r.RedoCount);
    }

    [Fact]
    public void DS001d_PopFromEmpty()
    {
        var r = new UndoRing<string>(3);
        var (_, ok) = r.PopForUndo();
        Assert.False(ok);
        var (_, ok2) = r.PopForRedo();
        Assert.False(ok2);
    }
}
