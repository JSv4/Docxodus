// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
#if !WASM_BUILD
using SkiaSharp;
#endif
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Globalization;

namespace Docxodus
{
    public class MetricsGetterSettings
    {
        public bool IncludeTextInContentControls;
        public bool RetrieveNamespaceList;
        public bool RetrieveContentTypeList;
    }

    public class MetricsGetter
    {
        public static XElement? GetMetrics(string fileName, MetricsGetterSettings settings)
        {
            FileInfo fi = new FileInfo(fileName);
            if (!fi.Exists)
                throw new FileNotFoundException("{0} does not exist.", fi.FullName);
            if (Util.IsWordprocessingML(fi.Extension))
            {
                WmlDocument wmlDoc = new WmlDocument(fi.FullName, true);
                return GetDocxMetrics(wmlDoc, settings);
            }
            return null;
        }

        public static XElement GetDocxMetrics(WmlDocument wmlDoc, MetricsGetterSettings settings)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    ms.Write(wmlDoc.DocumentByteArray, 0, wmlDoc.DocumentByteArray.Length);
                    using (WordprocessingDocument document = WordprocessingDocument.Open(ms, true))
                    {
                        bool hasTrackedRevisions = RevisionAccepter.HasTrackedRevisions(document);
                        if (hasTrackedRevisions)
                            RevisionAccepter.AcceptRevisions(document);
                        XElement metrics1 = GetWmlMetrics(wmlDoc.FileName, false, document, settings);
                        if (hasTrackedRevisions)
                            metrics1.Add(new XElement(H.RevisionTracking, new XAttribute(H.Val, true)));
                        return metrics1;
                    }
                }
            }
            catch (DocxodusException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink"))
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        ms.Write(wmlDoc.DocumentByteArray, 0, wmlDoc.DocumentByteArray.Length);
#if !NET35
                        UriFixer.FixInvalidUri(ms, brokenUri => FixUri(brokenUri));
#endif
                        wmlDoc = new WmlDocument("dummy.docx", ms.ToArray());
                    }
                    using (MemoryStream ms = new MemoryStream())
                    {
                        ms.Write(wmlDoc.DocumentByteArray, 0, wmlDoc.DocumentByteArray.Length);
                        using (WordprocessingDocument document = WordprocessingDocument.Open(ms, true))
                        {
                            bool hasTrackedRevisions = RevisionAccepter.HasTrackedRevisions(document);
                            if (hasTrackedRevisions)
                                RevisionAccepter.AcceptRevisions(document);
                            XElement metrics2 = GetWmlMetrics(wmlDoc.FileName, true, document, settings);
                            if (hasTrackedRevisions)
                                metrics2.Add(new XElement(H.RevisionTracking, new XAttribute(H.Val, true)));
                            return metrics2;
                        }
                    }
                }
            }
            var metrics = new XElement(H.Metrics,
                new XAttribute(H.FileName, wmlDoc.FileName ?? ""),
                new XAttribute(H.FileType, "WordprocessingML"),
                new XAttribute(H.Error, "Unknown error, metrics not determined"));
            return metrics;
        }

        /// <summary>
        /// Character-class-aware text-width estimate for LIST-MARKER runs whose font cannot be
        /// measured with real metrics (WASM builds, and hosts where the font is not installed).
        /// Returns width in font-size units — (summed per-character em fractions) × (sz / 2)
        /// points — the same unit convention the measurement path returns.
        ///
        /// The general fallback (a flat 0.6 em per character) roughly doubles the width of
        /// narrow-glyph markers such as "(a)"/"(iii)", pushing the computed pen position of a
        /// numbered paragraph past its text-indent tab stop so the marker's suffix tab resolved
        /// to the next default tab stop — every clause body ~0.25" too far right (issue #415).
        /// The per-class widths below track Times New Roman / Calibri averages closely enough
        /// that "does the number end before the text indent" decisions match real renderers.
        ///
        /// This is deliberately NOT the general text fallback. For ordinary tab runs the HTML
        /// layout reserves (tab stop − estimated following-text width) and the browser paints
        /// the real glyphs after that reservation, so the flat 0.6 em estimate must stay at or
        /// above browser fallback-font advances (monospace fixtures and DejaVu-class fallbacks
        /// sit at ~0.6 em) or lines near the column edge overflow and wrap. A list marker is
        /// safe: its wrapper is (chosen stop − marker start) with a flex tab absorbing any
        /// glyph-width error, so only the stop choice needs realistic widths.
        /// </summary>
        public static int EstimateTextWidth(decimal sz, string text)
        {
            if (string.IsNullOrEmpty(text))
                return 0;
            float em = 0f;
            foreach (var c in text)
                em += EstimateCharWidthEm(c);
            return (int)(em * (float)sz / 2f);
        }

        private static float EstimateCharWidthEm(char c)
        {
            switch (c)
            {
                case 'i': case 'j': case 'l': case '.': case ',': case '\'': case '|': case '`':
                    return 0.25f;
                case 't': case 'f': case 'r': case 'I': case '!': case ';': case ':':
                case '(': case ')': case '[': case ']': case '{': case '}':
                case '"': case '/': case '\\': case ' ': case '-':
                    return 0.33f;
                case 'm': case 'w': case 'M': case 'W': case '@': case '%':
                    return 0.85f;
                case '\u00AD': case '\u200B': case '\u200C': case '\u200D':
                    return 0f; // soft hyphen / zero-width characters
                default:
                    // CJK ideographs, Hangul syllables, and fullwidth forms occupy a full em.
                    if (c >= 0x2E80 && (c <= 0xD7A3 || (c >= 0xF900 && c <= 0xFAFF) || (c >= 0xFF01 && c <= 0xFF60)))
                        return 1.0f;
                    return char.IsUpper(c) ? 0.7f : 0.5f;
            }
        }

        private static int _getTextWidth(string fontName, bool bold, bool italic, decimal sz, string text)
        {
#if WASM_BUILD
            // In WASM builds, return an estimate based on character count and font size
            // This is a rough approximation since we don't have SkiaSharp for text measurement
            if (string.IsNullOrEmpty(text))
                return 0;
            // Approximate character width: ~0.6 of the font size for most fonts
            float charWidth = (float)sz * 0.6f / 2f;
            return (int)(text.Length * charWidth);
#else
            try
            {
                var weight = bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal;
                var slant = italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright;
                using var typeface = SKTypeface.FromFamilyName(fontName, weight, SKFontStyleWidth.Normal, slant);
                using var font = new SKFont(typeface, (float)sz / 2f);

                var result = (int)font.MeasureText(text);

                // If measurement returns 0 (font unavailable), use estimation
                if (result == 0 && !string.IsNullOrEmpty(text))
                {
                    float charWidth = (float)sz * 0.6f / 2f;
                    return (int)(text.Length * charWidth);
                }

                return result;
            }
            catch
            {
                // Fallback to estimation on any error
                if (string.IsNullOrEmpty(text))
                    return 0;
                float charWidth = (float)sz * 0.6f / 2f;
                return (int)(text.Length * charWidth);
            }
#endif
        }

        public static int GetTextWidth(string fontName, bool bold, bool italic, decimal sz, string text)
        {
            try
            {
                return _getTextWidth(fontName, bold, italic, sz, text);
            }
            catch (ArgumentException)
            {
                try
                {
                    // Try with regular style
                    return _getTextWidth(fontName, false, false, sz, text);
                }
                catch (ArgumentException)
                {
                    try
                    {
                        // Try with bold
                        return _getTextWidth(fontName, true, false, sz, text);
                    }
                    catch (ArgumentException)
                    {
                        // if both regular and bold fail, then get metrics for Times New Roman
                        return _getTextWidth("Times New Roman", bold, italic, sz, text);
                    }
                }
            }
            catch (OverflowException)
            {
                // This happened on Azure but interestingly enough not while testing locally.
                // Use estimation instead of returning 0
                if (string.IsNullOrEmpty(text))
                    return 0;
                float charWidth = (float)sz * 0.6f / 2f;
                return (int)(text.Length * charWidth);
            }
        }

        private static Uri FixUri(string brokenUri)
        {
            return new Uri("http://broken-link/");
        }

        private static XElement GetWmlMetrics(string? fileName, bool invalidHyperlink, WordprocessingDocument wDoc, MetricsGetterSettings settings)
        {
            XElement? parts = new XElement(H.Parts,
                wDoc.GetAllParts().Select(part =>
                {
                    return GetMetricsForWmlPart(part, settings);
                }));
            if (!parts.HasElements)
                parts = null;
            var metrics = new XElement(H.Metrics,
                new XAttribute(H.FileName, fileName ?? ""),
                new XAttribute(H.FileType, "WordprocessingML"),
                GetStyleHierarchy(wDoc),
                GetMiscWmlMetrics(wDoc, invalidHyperlink),
                parts,
                settings.RetrieveNamespaceList ? RetrieveNamespaceList(wDoc) : null,
                settings.RetrieveContentTypeList ? RetrieveContentTypeList(wDoc) : null
                );
            return metrics;
        }

        private static XElement RetrieveContentTypeList(OpenXmlPackage oxPkg)
        {
            var allParts = oxPkg.GetAllParts();
            var contentTypes = allParts
                .Select(p => p.ContentType)
                .Where(ct => ct != "application/vnd.openxmlformats-package.relationships+xml")
                .OrderBy(t => t)
                .Distinct();
            var xe = new XElement(H.ContentTypes,
                contentTypes.Select(ct => new XElement(H.ContentType, new XAttribute(H.Val, ct))));
            return xe;
        }

        private static XElement RetrieveNamespaceList(OpenXmlPackage oxPkg)
        {
            var allParts = oxPkg.GetAllParts();
            var xmlParts = allParts
                .Where(p => p.ContentType != "application/vnd.openxmlformats-package.relationships+xml")
                .Where(p => p.ContentType.ToLower().EndsWith("xml"));

            var uniqueNamespaces = new HashSet<string>();
            foreach (var xp in xmlParts)
            {
                using (Stream st = xp.GetStream())
                {
                    try
                    {
                        XDocument xdoc = XDocument.Load(st);
                        var namespaces = xdoc
                            .Descendants()
                            .Attributes()
                            .Where(a => a.IsNamespaceDeclaration)
                            .Select(a => string.Format("{0}|{1}", a.Name.LocalName, a.Value))
                            .OrderBy(t => t)
                            .Distinct()
                            .ToList();
                        foreach (var item in namespaces)
		                    uniqueNamespaces.Add(item);
                    }
                    // if catch exception, forget about it.  Just trying to get a most complete survey possible of all namespaces in all documents.
                    // if caught exception, chances are the document is bad anyway.
                    catch (Exception)
                    {
                        continue;
                    }
                }
            }
            var xe = new XElement(H.Namespaces,
                uniqueNamespaces.OrderBy(t => t).Select(n =>
                {
                    var spl = n.Split('|');
                    return new XElement(H.Namespace,
                        new XAttribute(H.NamespacePrefix, spl[0]),
                        new XAttribute(H.NamespaceName, spl[1]));
                }));
            return xe;
        }

        private static List<XElement> GetMiscWmlMetrics(WordprocessingDocument document, bool invalidHyperlink)
        {
            List<XElement> metrics = new List<XElement>();
            List<string> notes = new List<string>();
            Dictionary<XName, int> elementCountDictionary = new Dictionary<XName, int>();

            if (invalidHyperlink)
                metrics.Add(new XElement(H.InvalidHyperlink, new XAttribute(H.Val, invalidHyperlink)));

            bool valid = ValidateWordprocessingDocument(document, metrics, notes, elementCountDictionary);
            if (invalidHyperlink)
                valid = false;

            return metrics;
        }

        private static bool ValidateWordprocessingDocument(WordprocessingDocument wDoc, List<XElement> metrics, List<string> notes, Dictionary<XName, int> metricCountDictionary)
        {
            bool valid = ValidateAgainstSpecificVersion(wDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2007, H.SdkValidationError2007);
            valid |= ValidateAgainstSpecificVersion(wDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2010, H.SdkValidationError2010);
#if !NET35
            valid |= ValidateAgainstSpecificVersion(wDoc, metrics, DocumentFormat.OpenXml.FileFormatVersions.Office2013, H.SdkValidationError2013);
#endif

            int elementCount = 0;
            int paragraphCount = 0;
            int textCount = 0;
            foreach (var part in wDoc.ContentParts())
            {
                XDocument xDoc = part.GetXDocument();
                foreach (var e in xDoc.Descendants())
                {
                    if (e.Name == W.txbxContent)
                        IncrementMetric(metricCountDictionary, H.TextBox);
                    else if (e.Name == W.sdt)
                        IncrementMetric(metricCountDictionary, H.ContentControl);
                    else if (e.Name == W.customXml)
                        IncrementMetric(metricCountDictionary, H.CustomXmlMarkup);
                    else if (e.Name == W.fldChar)
                        IncrementMetric(metricCountDictionary, H.ComplexField);
                    else if (e.Name == W.fldSimple)
                        IncrementMetric(metricCountDictionary, H.SimpleField);
                    else if (e.Name == W.altChunk)
                        IncrementMetric(metricCountDictionary, H.AltChunk);
                    else if (e.Name == W.tbl)
                        IncrementMetric(metricCountDictionary, H.Table);
                    else if (e.Name == W.hyperlink)
                        IncrementMetric(metricCountDictionary, H.Hyperlink);
                    else if (e.Name == W.framePr)
                        IncrementMetric(metricCountDictionary, H.LegacyFrame);
                    else if (e.Name == W.control)
                        IncrementMetric(metricCountDictionary, H.ActiveX);
                    else if (e.Name == W.subDoc)
                        IncrementMetric(metricCountDictionary, H.SubDocument);
                    else if (e.Name == VML.imagedata || e.Name == VML.fill || e.Name == VML.stroke || e.Name == A.blip)
                    {
                        var relId = (string?)e.Attribute(R.embed);
                        if (relId != null)
                            ValidateImageExists(part, relId, metricCountDictionary);
                        relId = (string?)e.Attribute(R.pict);
                        if (relId != null)
                            ValidateImageExists(part, relId, metricCountDictionary);
                        relId = (string?)e.Attribute(R.id);
                        if (relId != null)
                            ValidateImageExists(part, relId, metricCountDictionary);
                    }

                    if (part.Uri == wDoc.MainDocumentPart!.Uri)
                    {
                        elementCount++;
                        if (e.Name == W.p)
                            paragraphCount++;
                        if (e.Name == W.t)
                            textCount += ((string)e).Length;
                    }
                }
            }

            foreach (var item in metricCountDictionary)
            {
                metrics.Add(
                    new XElement(item.Key, new XAttribute(H.Val, item.Value)));
            }

            metrics.Add(new XElement(H.ElementCount, new XAttribute(H.Val, elementCount)));
            metrics.Add(new XElement(H.AverageParagraphLength, new XAttribute(H.Val, (int)((double)textCount / (double)paragraphCount))));

            if (wDoc.GetAllParts().Any(part => part.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                metrics.Add(new XElement(H.EmbeddedXlsx, new XAttribute(H.Val, true)));

            NumberingFormatListAssembly(wDoc, metrics);

            XDocument wxDoc = wDoc.MainDocumentPart.GetXDocument();

            foreach (var d in wxDoc.Descendants())
            {
                if (d.Name == W.saveThroughXslt)
                {
                    string? rid = (string?)d.Attribute(R.id);
                    var tempExternalRelationship = wDoc
                        .MainDocumentPart!
                        .DocumentSettingsPart!
                        .ExternalRelationships
                        .FirstOrDefault(h => h.Id == rid);
                    if (tempExternalRelationship == null)
                        metrics.Add(new XElement(H.InvalidSaveThroughXslt, new XAttribute(H.Val, true)));
                    valid = false;
                }
                else if (d.Name == W.trackRevisions)
                    metrics.Add(new XElement(H.TrackRevisionsEnabled, new XAttribute(H.Val, true)));
                else if (d.Name == W.documentProtection)
                    metrics.Add(new XElement(H.DocumentProtection, new XAttribute(H.Val, true)));
            }

            FontAndCharSetAnalysis(wDoc, metrics, notes);

            return valid;
        }

        private static bool ValidateAgainstSpecificVersion(WordprocessingDocument wDoc, List<XElement> metrics, DocumentFormat.OpenXml.FileFormatVersions versionToValidateAgainst, XName versionSpecificMetricName)
        {
            OpenXmlValidator validator = new OpenXmlValidator(versionToValidateAgainst);
            var errors = validator.Validate(wDoc);
            bool valid = errors.Count() == 0;
            if (!valid)
            {
                if (!metrics.Any(e => e.Name == H.SdkValidationError))
                    metrics.Add(new XElement(H.SdkValidationError, new XAttribute(H.Val, true)));
                metrics.Add(new XElement(versionSpecificMetricName, new XAttribute(H.Val, true),
                    errors.Take(3).Select(err =>
                    {
                        StringBuilder sb = new StringBuilder();
                        if (err.Description.Length > 300)
                            sb.Append(PtUtils.MakeValidXml(err.Description.Substring(0, 300) + " ... elided ...") + Environment.NewLine);
                        else
                            sb.Append(PtUtils.MakeValidXml(err.Description) + Environment.NewLine);
                        sb.Append("  in part " + PtUtils.MakeValidXml(err.Part!.Uri.ToString()) + Environment.NewLine);
                        sb.Append("  at " + PtUtils.MakeValidXml(err.Path!.XPath) + Environment.NewLine);
                        return sb.ToString();
                    })));
            }
            return valid;
        }

        private static void IncrementMetric(Dictionary<XName, int> metricCountDictionary, XName xName)
        {
            if (metricCountDictionary.ContainsKey(xName))
                metricCountDictionary[xName] = metricCountDictionary[xName] + 1;
            else
                metricCountDictionary.Add(xName, 1);
        }

        private static void ValidateImageExists(OpenXmlPart part, string relId, Dictionary<XName, int> metrics)
        {
            var imagePart = part.Parts.FirstOrDefault(ipp => ipp.RelationshipId == relId);
            if (imagePart == null)
                IncrementMetric(metrics, H.ReferenceToNullImage);
        }


        private static void NumberingFormatListAssembly(WordprocessingDocument wDoc, List<XElement> metrics)
        {
            List<string> numFmtList = new List<string>();
            foreach (var part in wDoc.ContentParts())
            {
                var xDoc = part.GetXDocument();
                numFmtList = numFmtList.Concat(xDoc
                    .Descendants(W.p)
                    .Select(p =>
                    {
                        ListItemRetriever.RetrieveListItem(wDoc, p, null);
                        ListItemRetriever.ListItemInfo? lif = p.Annotation<ListItemRetriever.ListItemInfo>();
                        if (lif != null && lif.IsListItem && lif.Lvl(ListItemRetriever.GetParagraphLevel(p)) != null)
                        {
                            string? numFmtForLevel = (string?)lif.Lvl(ListItemRetriever.GetParagraphLevel(p))!.Elements(W.numFmt).Attributes(W.val).FirstOrDefault();
                            if (numFmtForLevel == null)
                            {
                                var numFmtElement = lif.Lvl(ListItemRetriever.GetParagraphLevel(p))!.Elements(MC.AlternateContent).Elements(MC.Choice).Elements(W.numFmt).FirstOrDefault();
                                if (numFmtElement != null && (string?)numFmtElement.Attribute(W.val) == "custom")
                                    numFmtForLevel = (string?)numFmtElement.Attribute(W.format);
                            }
                            return numFmtForLevel;
                        }
                        return null;
                    })
                    .Where(s => s != null)
                    .Select(s => s!)
                    .Distinct())
                    .ToList();
            }
            if (numFmtList.Any())
            {
                var nfls = numFmtList.StringConcatenate(s => s + ",").TrimEnd(',');
                metrics.Add(new XElement(H.NumberingFormatList, new XAttribute(H.Val, PtUtils.MakeValidXml(nfls))));
            }
        }

        class FormattingMetrics
        {
            public int RunCount;
            public int RunWithoutRprCount;
            public int ZeroLengthText;
            public int MultiFontRun;

            public int AsciiCharCount;
            public int CSCharCount;
            public int EastAsiaCharCount;
            public int HAnsiCharCount;

            public int AsciiRunCount;
            public int CSRunCount;
            public int EastAsiaRunCount;
            public int HAnsiRunCount;

            public List<string> Languages;

            public FormattingMetrics()
            {
                Languages = new List<string>();
            }
        }

        private static void FontAndCharSetAnalysis(WordprocessingDocument wDoc, List<XElement> metrics, List<string> notes)
        {
            FormattingAssemblerSettings settings = new FormattingAssemblerSettings
            {
                RemoveStyleNamesFromParagraphAndRunProperties = false,
                ClearStyles = true,
                RestrictToSupportedNumberingFormats = false,
                RestrictToSupportedLanguages = false,
            };
            FormattingAssembler.AssembleFormatting(wDoc, settings);
            var formattingMetrics = new FormattingMetrics();

            foreach (var part in wDoc.ContentParts())
            {
                var xDoc = part.GetXDocument();
                foreach (var run in xDoc.Descendants(W.r))
                {
                    formattingMetrics.RunCount++;
                    AnalyzeRun(run, metrics, notes, formattingMetrics, part.Uri.ToString());
                }
            }

            metrics.Add(new XElement(H.RunCount, new XAttribute(H.Val, formattingMetrics.RunCount)));
            if (formattingMetrics.RunWithoutRprCount > 0)
                metrics.Add(new XElement(H.RunWithoutRprCount, new XAttribute(H.Val, formattingMetrics.RunWithoutRprCount)));
            if (formattingMetrics.ZeroLengthText > 0)
                metrics.Add(new XElement(H.ZeroLengthText, new XAttribute(H.Val, formattingMetrics.ZeroLengthText)));
            if (formattingMetrics.MultiFontRun > 0)
                metrics.Add(new XElement(H.MultiFontRun, new XAttribute(H.Val, formattingMetrics.MultiFontRun)));
            if (formattingMetrics.AsciiCharCount > 0)
                metrics.Add(new XElement(H.AsciiCharCount, new XAttribute(H.Val, formattingMetrics.AsciiCharCount)));
            if (formattingMetrics.CSCharCount > 0)
                metrics.Add(new XElement(H.CSCharCount, new XAttribute(H.Val, formattingMetrics.CSCharCount)));
            if (formattingMetrics.EastAsiaCharCount > 0)
                metrics.Add(new XElement(H.EastAsiaCharCount, new XAttribute(H.Val, formattingMetrics.EastAsiaCharCount)));
            if (formattingMetrics.HAnsiCharCount > 0)
                metrics.Add(new XElement(H.HAnsiCharCount, new XAttribute(H.Val, formattingMetrics.HAnsiCharCount)));
            if (formattingMetrics.AsciiRunCount > 0)
                metrics.Add(new XElement(H.AsciiRunCount, new XAttribute(H.Val, formattingMetrics.AsciiRunCount)));
            if (formattingMetrics.CSRunCount > 0)
                metrics.Add(new XElement(H.CSRunCount, new XAttribute(H.Val, formattingMetrics.CSRunCount)));
            if (formattingMetrics.EastAsiaRunCount > 0)
                metrics.Add(new XElement(H.EastAsiaRunCount, new XAttribute(H.Val, formattingMetrics.EastAsiaRunCount)));
            if (formattingMetrics.HAnsiRunCount > 0)
                metrics.Add(new XElement(H.HAnsiRunCount, new XAttribute(H.Val, formattingMetrics.HAnsiRunCount)));

            if (formattingMetrics.Languages.Any())
            {
                var uls = formattingMetrics.Languages.StringConcatenate(s => s + ",").TrimEnd(',');
                metrics.Add(new XElement(H.Languages, new XAttribute(H.Val, PtUtils.MakeValidXml(uls))));
            }
        }

        private static void AnalyzeRun(XElement run, List<XElement> attList, List<string> notes, FormattingMetrics formattingMetrics, string uri)
        {
            var runText = run.Elements()
                .Where(e => e.Name == W.t || e.Name == W.delText)
                .Select(t => (string)t)
                .StringConcatenate();
            if (runText.Length == 0)
            {
                formattingMetrics.ZeroLengthText++;
                return;
            }
            var rPr = run.Element(W.rPr);
            if (rPr == null)
            {
                formattingMetrics.RunWithoutRprCount++;
                notes.Add(PtUtils.MakeValidXml(string.Format("Error in part {0}: run without rPr at {1}", uri, run.GetXPath())));
                rPr = new XElement(W.rPr);
            }
            FormattingAssembler.CharStyleAttributes csa = new FormattingAssembler.CharStyleAttributes(null, rPr);
            var fontTypeArray = runText
                .Select(ch => FormattingAssembler.DetermineFontTypeFromCharacter(ch, csa))
                .ToArray();
            var distinctFontTypeArray = fontTypeArray
                .Distinct()
                .ToArray();
            var distinctFonts = distinctFontTypeArray
                .Select(ft =>
                {
                    return GetFontFromFontType(csa, ft);
                })
                .Distinct();
            var languages = distinctFontTypeArray
                .Select(ft =>
                {
                    if (ft == FormattingAssembler.FontType.Ascii)
                        return csa.LatinLang;
                    if (ft == FormattingAssembler.FontType.CS)
                        return csa.BidiLang;
                    if (ft == FormattingAssembler.FontType.EastAsia)
                        return csa.EastAsiaLang;
                    //if (ft == FormattingAssembler.FontType.HAnsi)
                    return csa.LatinLang;
                })
                .Select(l =>
                {
                    if (l == "" || l == null)
                        return /* "Dflt:" + */ CultureInfo.CurrentCulture.Name;
                    return l;
                })
                //.Where(l => l != null && l != "")
                .Distinct();
            if (languages.Any(l => !formattingMetrics.Languages.Contains(l)))
                formattingMetrics.Languages = formattingMetrics.Languages.Concat(languages).Distinct().ToList();
            var multiFontRun = distinctFonts.Count() > 1;
            if (multiFontRun)
            {
                formattingMetrics.MultiFontRun++;

                formattingMetrics.AsciiCharCount += fontTypeArray.Where(ft => ft == FormattingAssembler.FontType.Ascii).Count();
                formattingMetrics.CSCharCount += fontTypeArray.Where(ft => ft == FormattingAssembler.FontType.CS).Count();
                formattingMetrics.EastAsiaCharCount += fontTypeArray.Where(ft => ft == FormattingAssembler.FontType.EastAsia).Count();
                formattingMetrics.HAnsiCharCount += fontTypeArray.Where(ft => ft == FormattingAssembler.FontType.HAnsi).Count();
            }
            else
            {
                switch (fontTypeArray[0])
                {
                    case FormattingAssembler.FontType.Ascii:
                        formattingMetrics.AsciiCharCount += runText.Length;
                        formattingMetrics.AsciiRunCount++;
                        break;
                    case FormattingAssembler.FontType.CS:
                        formattingMetrics.CSCharCount += runText.Length;
                        formattingMetrics.CSRunCount++;
                        break;
                    case FormattingAssembler.FontType.EastAsia:
                        formattingMetrics.EastAsiaCharCount += runText.Length;
                        formattingMetrics.EastAsiaRunCount++;
                        break;
                    case FormattingAssembler.FontType.HAnsi:
                        formattingMetrics.HAnsiCharCount += runText.Length;
                        formattingMetrics.HAnsiRunCount++;
                        break;
                }
            }
        }

        private static string GetFontFromFontType(FormattingAssembler.CharStyleAttributes csa, FormattingAssembler.FontType ft)
        {
            switch (ft)
            {
                case FormattingAssembler.FontType.Ascii:
                    return csa.AsciiFont;
                case FormattingAssembler.FontType.CS:
                    return csa.CsFont;
                case FormattingAssembler.FontType.EastAsia:
                    return csa.EastAsiaFont;
                case FormattingAssembler.FontType.HAnsi:
                    return csa.HAnsiFont;
                default: // dummy
                    return csa.AsciiFont;
            }
        }

        private static object? GetStyleHierarchy(WordprocessingDocument document)
        {
            var stylePart = document.MainDocumentPart!.StyleDefinitionsPart;
            if (stylePart == null)
                return null;
            var xd = stylePart.GetXDocument();
            var stylesWithPath = xd.Root!
                .Elements(W.style)
                .Select(s =>
                {
                    // w:styleId is required by the WordprocessingML schema on every w:style.
                    var styleString = (string)s.Attribute(W.styleId)!;
                    var thisStyle = s;
                    while (true)
                    {
                        var baseStyle = (string?)thisStyle.Elements(W.basedOn).Attributes(W.val).FirstOrDefault();
                        if (baseStyle == null)
                            break;
                        styleString = baseStyle + "/" + styleString;
                        thisStyle = xd.Root.Elements(W.style).FirstOrDefault(ts => (string)ts.Attribute(W.styleId)! == baseStyle);
                        if (thisStyle == null)
                            break;
                    }
                    return styleString;
                })
                .OrderBy(n => n)
                .ToList();
            XElement styleHierarchy = new XElement(H.StyleHierarchy);
            foreach (var item in stylesWithPath)
            {
                var styleChain = item.Split('/');
                XElement? elementToAddTo = styleHierarchy;
                foreach (var inChain in styleChain.PtSkipLast(1))
                    elementToAddTo = elementToAddTo!.Elements(H.Style).FirstOrDefault(z => (string)z.Attribute(H.Id)! == inChain);
                var styleToAdd = styleChain.Last();
                elementToAddTo!.Add(
                    new XElement(H.Style,
                        new XAttribute(H.Id, styleChain.Last()),
                        new XAttribute(H.Type, (string?)xd.Root.Elements(W.style).First(z => (string)z.Attribute(W.styleId)! == styleToAdd).Attribute(W.type) ?? "")));
            }
            return styleHierarchy;
        }

        private static XElement? GetMetricsForWmlPart(OpenXmlPart part, MetricsGetterSettings settings)
        {
            XElement? contentControls = null;
            if (part is MainDocumentPart ||
                part is HeaderPart ||
                part is FooterPart ||
                part is FootnotesPart ||
                part is EndnotesPart)
            {
                var xd = part.GetXDocument();
                contentControls = (XElement?)GetContentControlsTransform(xd.Root!, settings);
                if (contentControls?.HasElements != true)
                    contentControls = null;
            }
            var partMetrics = new XElement(H.Part,
                new XAttribute(H.ContentType, part.ContentType),
                new XAttribute(H.Uri, part.Uri.ToString()),
                contentControls);
            if (partMetrics.HasElements)
                return partMetrics;
            return null;
        }

        private static object? GetContentControlsTransform(XNode node, MetricsGetterSettings settings)
        {
            XElement? element = node as XElement;
            if (element != null)
            {
                if (element == element.Document!.Root)
                    return new XElement(H.ContentControls,
                        element.Nodes().Select(n => GetContentControlsTransform(n, settings)));

                if (element.Name == W.sdt)
                {
                    var tag = (string?)element.Elements(W.sdtPr).Elements(W.tag).Attributes(W.val).FirstOrDefault();
                    XAttribute? tagAttr = tag != null ? new XAttribute(H.Tag, tag) : null;

                    var alias = (string?)element.Elements(W.sdtPr).Elements(W.alias).Attributes(W.val).FirstOrDefault();
                    XAttribute? aliasAttr = alias != null ? new XAttribute(H.Alias, alias) : null;

                    // element is never the document root here (that case returns earlier above),
                    // so it always has a parent and GetXPath() always returns non-null for it.
                    var xPathAttr = new XAttribute(H.XPath, element.GetXPath()!);

                    var isText = element.Elements(W.sdtPr).Elements(W.text).Any();
                    var isBibliography = element.Elements(W.sdtPr).Elements(W.bibliography).Any();
                    var isCitation = element.Elements(W.sdtPr).Elements(W.citation).Any();
                    var isComboBox = element.Elements(W.sdtPr).Elements(W.comboBox).Any();
                    var isDate = element.Elements(W.sdtPr).Elements(W.date).Any();
                    var isDocPartList = element.Elements(W.sdtPr).Elements(W.docPartList).Any();
                    var isDocPartObj = element.Elements(W.sdtPr).Elements(W.docPartObj).Any();
                    var isDropDownList = element.Elements(W.sdtPr).Elements(W.dropDownList).Any();
                    var isEquation = element.Elements(W.sdtPr).Elements(W.equation).Any();
                    var isGroup = element.Elements(W.sdtPr).Elements(W.group).Any();
                    var isPicture = element.Elements(W.sdtPr).Elements(W.picture).Any();
                    var isRichText = element.Elements(W.sdtPr).Elements(W.richText).Any() ||
                        (! isText && 
                        ! isBibliography && 
                        ! isCitation && 
                        ! isComboBox && 
                        ! isDate && 
                        ! isDocPartList && 
                        ! isDocPartObj && 
                        ! isDropDownList && 
                        ! isEquation && 
                        ! isGroup && 
                        ! isPicture);
                    string? type = null;
                    if (isText        ) type = "Text";
                    if (isBibliography) type = "Bibliography";
                    if (isCitation    ) type = "Citation";
                    if (isComboBox    ) type = "ComboBox";
                    if (isDate        ) type = "Date";
                    if (isDocPartList ) type = "DocPartList";
                    if (isDocPartObj  ) type = "DocPartObj";
                    if (isDropDownList) type = "DropDownList";
                    if (isEquation    ) type = "Equation";
                    if (isGroup       ) type = "Group";
                    if (isPicture     ) type = "Picture";
                    if (isRichText    ) type = "RichText";
                    // isRichText's own fallback disjunct (all other is* flags false) guarantees
                    // exactly one of the twelve branches above always fires.
                    var typeAttr = new XAttribute(H.Type, type!);

                    return new XElement(H.ContentControl,
                        typeAttr,
                        tagAttr,
                        aliasAttr,
                        xPathAttr,
                        element.Nodes().Select(n => GetContentControlsTransform(n, settings)));
                }

                return element.Nodes().Select(n => GetContentControlsTransform(n, settings));
            }
            if (settings.IncludeTextInContentControls)
                return node;
            return null;
        }
    }

    public static class H
    {
        public static XName ActiveX = "ActiveX";
        public static XName Alias = "Alias";
        public static XName AltChunk = "AltChunk";
        public static XName Arguments = "Arguments";
        public static XName AsciiCharCount = "AsciiCharCount";
        public static XName AsciiRunCount = "AsciiRunCount";
        public static XName AverageParagraphLength = "AverageParagraphLength";
        public static XName BaselineReport = "BaselineReport";
        public static XName Batch = "Batch";
        public static XName BatchName = "BatchName";
        public static XName BatchSelector = "BatchSelector";
        public static XName CSCharCount = "CSCharCount";
        public static XName CSRunCount = "CSRunCount";
        public static XName Catalog = "Catalog";
        public static XName CatalogList = "CatalogList";
        public static XName CatalogListFile = "CatalogListFile";
        public static XName CaughtException = "CaughtException";
        public static XName Cell = "Cell";
        public static XName Column = "Column";
        public static XName Columns = "Columns";
        public static XName ComplexField = "ComplexField";
        public static XName Computer = "Computer";
        public static XName Computers = "Computers";
        public static XName ContentControl = "ContentControl";
        public static XName ContentControls = "ContentControls";
        public static XName ContentType = "ContentType";
        public static XName ContentTypes = "ContentTypes";
        public static XName CustomXmlMarkup = "CustomXmlMarkup";
        public static XName DLL = "DLL";
        public static XName DefaultDialogValuesFile = "DefaultDialogValuesFile";
        public static XName DefaultValues = "DefaultValues";
        public static XName Dependencies = "Dependencies";
        public static XName DestinationDir = "DestinationDir";
        public static XName Directory = "Directory";
        public static XName DirectoryPattern = "DirectoryPattern";
        public static XName DisplayName = "DisplayName";
        public static XName DoJobQueueName = "DoJobQueueName";
        public static XName Document = "Document";
        public static XName DocumentProtection = "DocumentProtection";
        public static XName DocumentSelector = "DocumentSelector";
        public static XName DocumentType = "DocumentType";
        public static XName Documents = "Documents";
        public static XName EastAsiaCharCount = "EastAsiaCharCount";
        public static XName EastAsiaRunCount = "EastAsiaRunCount";
        public static XName ElementCount = "ElementCount";
        public static XName EmbeddedXlsx = "EmbeddedXlsx";
        public static XName Error = "Error";
        public static XName Exception = "Exception";
        public static XName Exe = "Exe";
        public static XName ExeRoot = "ExeRoot";
        public static XName Extension = "Extension";
        public static XName File = "File";
        public static XName FileLength = "FileLength";
        public static XName FileName = "FileName";
        public static XName FilePattern = "FilePattern";
        public static XName FileType = "FileType";
        public static XName Guid = "Guid";
        public static XName HAnsiCharCount = "HAnsiCharCount";
        public static XName HAnsiRunCount = "HAnsiRunCount";
        public static XName RevisionTracking = "RevisionTracking";
        public static XName Hyperlink = "Hyperlink";
        public static XName IPAddress = "IPAddress";
        public static XName Id = "Id";
        public static XName Invalid = "Invalid";
        public static XName InvalidHyperlink = "InvalidHyperlink";
        public static XName InvalidHyperlinkException = "InvalidHyperlinkException";
        public static XName InvalidSaveThroughXslt = "InvalidSaveThroughXslt";
        public static XName JobComplete = "JobComplete";
        public static XName JobExe = "JobExe";
        public static XName JobName = "JobName";
        public static XName JobSpec = "JobSpec";
        public static XName Languages = "Languages";
        public static XName LegacyFrame = "LegacyFrame";
        public static XName LocalDoJobQueue = "LocalDoJobQueue";
        public static XName MachineName = "MachineName";
        public static XName MaxConcurrentJobs = "MaxConcurrentJobs";
        public static XName MaxDocumentsInJob = "MaxDocumentsInJob";
        public static XName MaxParagraphLength = "MaxParagraphLength";
        public static XName Message = "Message";
        public static XName Metrics = "Metrics";
        public static XName MultiDirectory = "MultiDirectory";
        public static XName MultiFontRun = "MultiFontRun";
        public static XName MultiServerQueue = "MultiServerQueue";
        public static XName Name = "Name";
        public static XName Namespaces = "Namespaces";
        public static XName Namespace = "Namespace";
        public static XName NamespaceName = "NamespaceName";
        public static XName NamespacePrefix = "NamespacePrefix";
        public static XName Note = "Note";
        public static XName NumberingFormatList = "NumberingFormatList";
        public static XName ObjectDisposedException = "ObjectDisposedException";
        public static XName ParagraphCount = "ParagraphCount";
        public static XName Part = "Part";
        public static XName Parts = "Parts";
        public static XName PassedDocuments = "PassedDocuments";
        public static XName Path = "Path";
        public static XName ProduceCatalog = "ProduceCatalog";
        public static XName ReferenceToNullImage = "ReferenceToNullImage";
        public static XName Report = "Report";
        public static XName Root = "Root";
        public static XName RootDirectory = "RootDirectory";
        public static XName Row = "Row";
        public static XName RunCount = "RunCount";
        public static XName RunWithoutRprCount = "RunWithoutRprCount";
        public static XName SdkValidationError = "SdkValidationError";
        public static XName SdkValidationError2007 = "SdkValidationError2007";
        public static XName SdkValidationError2010 = "SdkValidationError2010";
        public static XName SdkValidationError2013 = "SdkValidationError2013";
        public static XName Sheet = "Sheet";
        public static XName Sheets = "Sheets";
        public static XName SimpleField = "SimpleField";
        public static XName Skip = "Skip";
        public static XName SmartTag = "SmartTag";
        public static XName SourceRootDir = "SourceRootDir";
        public static XName SpawnerJobExeLocation = "SpawnerJobExeLocation";
        public static XName SpawnerReady = "SpawnerReady";
        public static XName Style = "Style";
        public static XName StyleHierarchy = "StyleHierarchy";
        public static XName SubDocument = "SubDocument";
        public static XName Table = "Table";
        public static XName TableData = "TableData";
        public static XName Tag = "Tag";
        public static XName Take = "Take";
        public static XName TextBox = "TextBox";
        public static XName TrackRevisionsEnabled = "TrackRevisionsEnabled";
        public static XName Type = "Type";
        public static XName Uri = "Uri";
        public static XName Val = "Val";
        public static XName Valid = "Valid";
        public static XName WindowStyle = "WindowStyle";
        public static XName XPath = "XPath";
        public static XName ZeroLengthText = "ZeroLengthText";
        public static XName custDataLst = "custDataLst";
        public static XName custShowLst = "custShowLst";
        public static XName kinsoku = "kinsoku";
        public static XName modifyVerifier = "modifyVerifier";
        public static XName photoAlbum = "photoAlbum";
    }
}
