// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;


namespace Docxodus.Tests
{
    public class RaTests
    {
        [Theory]
        [InlineData("RA001-Tracked-Revisions-01.docx")]
        [InlineData("RA001-Tracked-Revisions-02.docx")]

        public void RA001(string name)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

            WmlDocument notAccepted = new WmlDocument(sourceDocx.FullName);
            Assert.True(RevisionProcessor.HasTrackedRevisions(notAccepted), "fixture must carry tracked revisions");

            WmlDocument afterAccepting = RevisionAccepter.AcceptRevisions(notAccepted);
            Assert.False(RevisionProcessor.HasTrackedRevisions(afterAccepting));

            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-processed-by-RevisionAccepter.docx")));
            afterAccepting.SaveAs(processedDestDocx.FullName);
            Assert.False(RevisionProcessor.HasTrackedRevisions(new WmlDocument(processedDestDocx.FullName)));
        }

    }
}

