#nullable enable

// Regression tests for ListItemRetriever, found while removing #nullable disable
// from the file (issue #650).

using System.Xml.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

public class ListItemRetrieverTests
{
    /// <summary>
    /// Minimal numbering.xml with one abstractNum defining NO levels at all, and one num
    /// (numId=1) referencing it. ListItemSourceSet/ListItemSource's own internal per-level
    /// fallback (Main, then NumStyleLink) therefore exhausts without finding anything at
    /// any level, which is exactly the precondition for ListItemInfo.Lvl's own outer
    /// fallback loop to run.
    /// </summary>
    private static XDocument BuildNumberingWithNoLevels()
    {
        return new XDocument(
            new XElement(W.numbering,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XElement(W.abstractNum,
                    new XAttribute(W.abstractNumId, 1)),
                new XElement(W.num,
                    new XAttribute(W.numId, 1),
                    new XElement(W.abstractNumId, new XAttribute(W.val, 1)))));
    }

    [Fact]
    public void ListItemInfo_Lvl_StyleOnlySource_FallbackLoopDoesNotThrow()
    {
        // Regression test: ListItemInfo.Lvl's paragraph-less fallback loop used to walk
        // FromParagraph.Lvl(i) even when FromParagraph was null (a copy-paste bug from the
        // paragraph-sourced branch above it). ListItemSource.Lvl already exhausts every
        // level 0..ilvl internally (via Main then NumStyleLink) before returning null, so
        // this outer loop only runs once that inner search has already failed everywhere —
        // exactly the case here, with an abstractNum that defines no levels at all. Before
        // the fix, entering the loop threw a NullReferenceException on the first iteration
        // instead of returning null like every other "not found" path in this file.
        var numXDoc = BuildNumberingWithNoLevels();
        var styleSource = new ListItemRetriever.ListItemSource(numXDoc, numXDoc, numId: 1);

        var listItemInfo = new ListItemRetriever.ListItemInfo
        {
            FromStyle = styleSource,
            // FromParagraph left null: this is a style-only list item.
        };

        var lvl = listItemInfo.Lvl(1);

        Assert.Null(lvl);
    }
}
