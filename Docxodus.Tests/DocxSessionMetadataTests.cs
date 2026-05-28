#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for the block-metadata read surface on <see cref="DocxSession"/>
/// (<c>GetBlockMetadata</c>, <c>GetBlockMetadatas</c>, <c>GetListMembership</c>,
/// <c>GetSectionInfo</c>). Test IDs follow the <c>BM###</c> prefix convention.
/// </summary>
public class DocxSessionMetadataTests
{
    [Fact]
    public void BM001_GetBlockMetadata_PlainParagraph_ReturnsKindAndScope()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");

        var meta = session.GetBlockMetadata(anchor.Anchor.Id);

        Assert.NotNull(meta);
        Assert.Equal("p", meta!.Kind);
        Assert.Equal("body", meta.Scope);
        Assert.Null(meta.StyleId);
        Assert.Null(meta.StyleName);
        Assert.Null(meta.OutlineLevel);
        Assert.Null(meta.List);
        Assert.False(meta.HasInlineFormatting);
    }

    [Fact]
    public void BM002_GetListMembership_InlineNumPr_BulletList_ReturnsListFacts()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS002_BulletedList());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "li");

        var list = session.GetListMembership(anchor.Anchor.Id);

        Assert.NotNull(list);
        Assert.Equal(1, list!.NumId);
        Assert.Equal(0, list.AbstractNumId);
        Assert.Equal(0, list.Level);
        Assert.Equal(NumberFormat.Bullet, list.Format);
        Assert.True(list.IsAutoNumbered);
        Assert.False(list.FromStyle);
        Assert.Null(list.StartOverride);
    }

    [Fact]
    public void BM003_GetListMembership_NotAList_ReturnsNull()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");

        Assert.Null(session.GetListMembership(anchor.Anchor.Id));
    }

    [Fact]
    public void BM004_GetListMembership_StyleInheritedNumPr_SetsFromStyleTrue()
    {
        using var session = new DocxSession(DocxSessionTests.BuildBM_StyleInheritedList());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "li");

        var list = session.GetListMembership(anchor.Anchor.Id);

        Assert.NotNull(list);
        Assert.True(list!.FromStyle);
        Assert.Equal(1, list.NumId);
        Assert.Equal(0, list.Level);
        Assert.Equal(NumberFormat.Bullet, list.Format);
    }
}
