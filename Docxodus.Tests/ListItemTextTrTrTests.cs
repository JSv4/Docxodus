// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

// Turkish (tr-TR) letter numbering must wrap like the default formatter does, so a large
// w:start cannot make one list marker millions of characters long.

using System;
using System.Linq;
using Xunit;

namespace Docxodus.Tests;

public class ListItemTextTrTrTests
{
    private const string Upper = "ABCÇDEFGĞHIİJKLMNOÖPRSŞTUÜVYZ";
    private const string Lower = "abcçdefgğhıijklmnoöprsştuüvyz";

    [Theory]
    [InlineData("upperLetter", Upper)]
    [InlineData("lowerLetter", Lower)]
    public void Letters_WithinOneCycle_RepeatTheLetterPerPass(string numFmt, string alphabet)
    {
        Assert.Equal(alphabet[0].ToString(), ListItemTextGetter_tr_TR.GetListItemText("tr-TR", 1, numFmt));
        Assert.Equal(alphabet[28].ToString(), ListItemTextGetter_tr_TR.GetListItemText("tr-TR", 29, numFmt));
        Assert.Equal(new string(alphabet[0], 2), ListItemTextGetter_tr_TR.GetListItemText("tr-TR", 30, numFmt));
        Assert.Equal(new string(alphabet[28], 30), ListItemTextGetter_tr_TR.GetListItemText("tr-TR", 870, numFmt));
    }

    [Theory]
    [InlineData("upperLetter")]
    [InlineData("lowerLetter")]
    public void Letters_AfterThirtyRepeats_WrapToTheStart(string numFmt)
    {
        for (var levelNumber = 1; levelNumber <= 870; levelNumber++)
        {
            Assert.Equal(
                ListItemTextGetter_tr_TR.GetListItemText("tr-TR", levelNumber, numFmt),
                ListItemTextGetter_tr_TR.GetListItemText("tr-TR", levelNumber + 870, numFmt));
        }
    }

    [Theory]
    [InlineData("upperLetter")]
    [InlineData("lowerLetter")]
    public void Letters_AtInt32MaxValue_StayBounded(string numFmt)
    {
        var text = ListItemTextGetter_tr_TR.GetListItemText("tr-TR", int.MaxValue, numFmt);

        Assert.InRange(text.Length, 1, 30);
        Assert.Single(text.Distinct());
    }
}
