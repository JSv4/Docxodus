// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Docxodus;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class MgTests
    {
        [Theory]
        [InlineData("DA001-TemplateDocument.docx")]
        [InlineData("DA002-TemplateDocument.docx")]
        [InlineData("DA003-Select-XPathFindsNoData.docx")]
        [InlineData("DA004-Select-XPathFindsNoDataOptional.docx")]
        [InlineData("DA005-SelectRowData-NoData.docx")]
        [InlineData("DA006-SelectTestValue-NoData.docx")]
        public void MG001(string name)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo fi = new FileInfo(Path.Combine(sourceDir.FullName, name));

            MetricsGetterSettings settings = new MetricsGetterSettings()
            {
                IncludeTextInContentControls = false,
                RetrieveNamespaceList = true,
                RetrieveContentTypeList = true,
            };

            var extension = fi.Extension.ToLower();
            XElement metrics = null;
            if (Util.IsWordprocessingML(extension))
            {
                WmlDocument wmlDocument = new WmlDocument(fi.FullName);
                metrics = MetricsGetter.GetDocxMetrics(wmlDocument, settings);
            }

            Assert.NotNull(metrics);
        }
    }
}

#endif
