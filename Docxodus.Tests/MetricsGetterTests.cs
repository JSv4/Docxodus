// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;


namespace Docxodus.Tests
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

        [Fact]
        public void MG002_GetDocxMetrics_ByteLoadedDocumentWithNoFileName()
        {
            // Regression test: WmlDocument(fileName: null, bytes) is a normal, supported
            // construction (e.g. a document received over the wire rather than loaded from
            // disk). GetDocxMetrics used to build an XAttribute directly from the null
            // FileName, which throws ArgumentNullException at the point of construction.
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo fi = new FileInfo(Path.Combine(sourceDir.FullName, "DA001-TemplateDocument.docx"));
            byte[] bytes = File.ReadAllBytes(fi.FullName);
            WmlDocument wmlDocument = new WmlDocument(null, bytes);

            MetricsGetterSettings settings = new MetricsGetterSettings()
            {
                IncludeTextInContentControls = false,
            };

            XElement metrics = MetricsGetter.GetDocxMetrics(wmlDocument, settings);

            Assert.NotNull(metrics);
            Assert.Equal("", (string?)metrics.Attribute(H.FileName));
        }

        [Fact]
        public void MG003_ReferenceToNullImage_DanglingBlipRelationship_IsCounted()
        {
            // Regression test: ValidateImageExists located the referenced part with
            // Parts.FirstOrDefault(...), whose return type (IdPartPair) is a struct, then
            // compared the result to null. FirstOrDefault() on a struct sequence returns
            // default(IdPartPair) on a miss, never null, so the comparison was always false
            // and ReferenceToNullImage could never be incremented, even for a document with a
            // genuinely dangling image relationship.
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("hello")))));
                doc.Save();

                var xDoc = mainPart.GetXDocument();
                xDoc.Root!.Element(W.body)!.Add(
                    new XElement(a + "blip", new XAttribute(r + "embed", "rIdDoesNotExist")));
                mainPart.PutXDocument();
            }

            WmlDocument wmlDocument = new WmlDocument("dangling-image.docx", stream.ToArray());
            XElement metrics = MetricsGetter.GetDocxMetrics(wmlDocument, new MetricsGetterSettings());

            XElement? referenceToNullImage = metrics.Descendants(H.ReferenceToNullImage).FirstOrDefault();
            Assert.NotNull(referenceToNullImage);
            Assert.Equal(1, (int)referenceToNullImage!.Attribute(H.Val)!);
        }

        [Fact]
        public void MG004_TextWidth_IsInPixelsOfTheRunsPointSize()
        {
            // Every caller converts GetTextWidth from 96-DPI pixels. Measuring at the point size made
            // each installed font three quarters as wide as Word sets it (issue #1659): Word 2024 sets
            // this text in 10 pt Times New Roman 233.5 pt wide, which a host with the font measured as
            // 175 pt. Where the font is missing, Skia substitutes the same fallback on both sides.
            const string text = "the parties agree that the supplier shall deliver the services";
            using var typeface = SkiaSharp.SKTypeface.FromFamilyName("Times New Roman");
            using var tenPoints = new SkiaSharp.SKFont(typeface, 10f);
            var points = tenPoints.MeasureText(text);
            Assert.True(points > 0, "no font to measure with");

            var pixels = MetricsGetter.GetTextWidth("Times New Roman", false, false, 20m, text);
            Assert.InRange(pixels, (int)(points * 96 / 72) - 1, (int)(points * 96 / 72) + 1);
        }
    }
}

