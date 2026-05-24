// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.IO;
using Docxodus;
using Xunit;

namespace OxPt
{
    public class WmlComparerBodyLevelBookmarkTests
    {
        [Fact]
        public void Compare_WithBodyLevelBookmarks_DoesNotThrowNullReference()
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var original = new WmlDocument(Path.Combine(sourceDir.FullName, "WC/WC-BodyBookmarks-Before.docx"));
            var modified = new WmlDocument(Path.Combine(sourceDir.FullName, "WC/WC-BodyBookmarks-After.docx"));

            var settings = new WmlComparerSettings();

            // The modified document has bookmarkStart/bookmarkEnd as direct children
            // of w:body (siblings of w:p). This caused a NullReferenceException in
            // FindIndexOfNextParaMark because it assumed all ComparisonUnits were
            // ComparisonUnitWord, but body-level bookmarks produce other types.
            //
            // This document pair also triggers a separate "Internal error" in
            // ProcessFootnoteEndnote (the modified document converted endnotes to
            // footnotes). That is a different bug — this test verifies only that the
            // NullReferenceException in FindIndexOfNextParaMark is fixed.
            var ex = Record.Exception(() => WmlComparer.Compare(original, modified, settings));

            Assert.True(
                ex is not NullReferenceException,
                $"Expected no NullReferenceException but got: {ex}");
        }
    }
}
