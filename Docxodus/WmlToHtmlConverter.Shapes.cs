// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus
{
    public static partial class WmlToHtmlConverter
    {
        private static readonly XName WpgGroup = WPG.wpg + "wgp";
        private static readonly XName WpgNestedGroup = WPG.wpg + "grpSp";
        private static readonly XName WpgGroupProperties = WPG.wpg + "grpSpPr";
        private static readonly XName WpcCanvas = WPC.wpc + "wpc";

        /// <summary>
        /// Draws a DrawingML drawing that is neither a picture, a chart nor a text box — a
        /// <c>wps:wsp</c> shape, a <c>wpg:wgp</c> group or a <c>wpc:wpc</c> canvas — as SVG sized
        /// to its <c>wp:extent</c>. Rectangles, rounded rectangles, ellipses and lines are drawn
        /// with their fill and outline; a group maps its children through its child coordinate
        /// space (<c>a:chOff</c>/<c>a:chExt</c>).
        /// <para>
        /// Any inline drawing nothing could draw — other geometry, SmartArt, a chart type that is
        /// not projected, a picture whose image is missing — keeps its extent as an
        /// <see cref="UnrenderedDrawing">unrendered box</see> rather than vanishing. The
        /// empty-paragraph pass counts every inline <c>w:drawing</c> as paragraph content, so a
        /// drawing that rendered to nothing left its paragraph with no line box at all (issue
        /// #1270). The one exception is a picture when the caller passed no image handler: that
        /// caller asked for pictures to be left out. (A handler that returns null for one picture
        /// declines only that picture, whose space is kept.)
        /// </para>
        /// <para>
        /// Floating (<c>wp:anchor</c>) shapes keep their SVG and anchor geometry for the host,
        /// while their paragraph mark remains separate from the drawing.
        /// </para>
        /// </summary>
        private static XElement? ProcessDrawingShape(WordprocessingDocument wordDoc, XElement drawing,
            bool rendersPictures, WmlToHtmlConverterSettings? settings)
        {
            var container = drawing.Elements().FirstOrDefault(e => e.Name == WP.inline || e.Name == WP.anchor);
            if (container == null)
                return null;

            var graphicData = container.Element(A.graphic)?.Element(A.graphicData);
            var content = graphicData?.Elements().FirstOrDefault();
            if (content?.Name == Pic._pic && !rendersPictures)
                return null;

            var extent = container.Element(WP.extent);
            var width = Math.Max(0L, (long?)extent?.Attribute(NoNamespace.cx) ?? 0L);
            var height = Math.Max(0L, (long?)extent?.Attribute(NoNamespace.cy) ?? 0L);
            var context = new ShapeContext(wordDoc);
            var frame = new ShapeBox(0, 0, width, height);
            // Coordinates are EMUs, scaled to the SVG's CSS pixels (9525 EMUs each) by this group.
            var drawingSpace = DrawingSpace();
            if (content?.Name == WPS.wsp)
                AddShapeElement(drawingSpace, content, frame, null, context);
            else if (content?.Name == WpgGroup)
                AddGroup(drawingSpace, content, content.Element(WpgGroupProperties), frame, null, context);
            else if (content?.Name == WpcCanvas)
                AddGroupChildren(drawingSpace, content, ShapeMap.Identity, null, context);
            else
                context.Undrawn++;
            if (context.Drawn == 0 && context.Undrawn > 0)
                return PlaceDrawingShape(UnrenderedDrawing(settings, drawing, content, width, height), container, settings);

            var docPr = container.Element(WP.docPr);
            var altText = (string?)docPr?.Attribute(NoNamespace.descr) ?? (string?)docPr?.Attribute(NoNamespace.name);
            var svg = DrawingSvg(width, height, drawingSpace);
            if (string.IsNullOrEmpty(altText))
                svg.Add(new XAttribute("aria-hidden", "true"));
            else
                svg.Add(new XAttribute("role", "img"), new XAttribute("aria-label", altText));

            // An inline-block whose only content is a block-level SVG has no line box, so its
            // baseline is its bottom edge: the drawing sits on the text baseline, as Word places
            // an inline object.
            var wrapper = new XElement(Xhtml.span, svg);
            if (context.Undrawn > 0)
            {
                // Part of the drawing could not be drawn. Say so where a consumer can see it: the
                // PDF export inventories data-content-type as unsupported content.
                wrapper.Add(new XAttribute("data-docx-shape", "partial"));
                if (settings is { RenderUnsupportedContentPlaceholders: true, IncludeUnsupportedContentMetadata: true })
                    wrapper.Add(new XAttribute("data-content-type", nameof(UnsupportedContentType.Other)),
                        new XAttribute("data-element-name", content!.Name.LocalName),
                        new XAttribute("title", "Part of this drawing is not rendered in HTML output"));
            }
            var style = new Dictionary<string, string> { { "display", "inline-block" } };
            AddEmuDimensions(style, width, height);
            wrapper.AddAnnotation(style);
            return PlaceDrawingShape(wrapper, container, settings);
        }

        private static XElement PlaceDrawingShape(XElement wrapper, XElement container,
            WmlToHtmlConverterSettings? settings)
        {
            if (container.Name != WP.anchor) return wrapper;
            var options = settings ?? new WmlToHtmlConverterSettings();
            var style = wrapper.Annotation<Dictionary<string, string>>()!;
            AddDrawingFrameStyle(style, wrapper, container, options, autoFit: false);
            wrapper.SetAttributeValue("data-docx-shape-anchor", "true");
            style["position"] = "absolute";
            PrepareRelativeShapeSvg(wrapper, options);
            if (container.Ancestors(W.tc).Any())
                ReportShapeLayoutLoss(wrapper, options, "Drawing placement within a table cell is approximated");

            var wrap = container.Elements().FirstOrDefault(e => e.Name == WP.wrapTopAndBottom ||
                e.Name == WP.wrapSquare || e.Name == WP.wrapTight || e.Name == WP.wrapThrough);
            var kind = wrap?.Name.LocalName ?? "wrapNone";
            wrapper.SetAttributeValue("data-docx-anchor-wrap", kind);
            var topDistance = ShapeAnchorNumber(wrapper, "wrap-top");
            if (wrap != null && topDistance > 0)
                ReportShapeLayoutLoss(wrapper, options, "Clearance above this drawing's anchor paragraph is not fully represented");
            var horizontal = container.Element(WP.positionH);
            var vertical = container.Element(WP.positionV);
            var paragraphRelative = !IsTrueAttribute(container, "simplePos") && (string?)vertical?.Attribute(NoNamespace.relativeFrom) == "paragraph" &&
                vertical.Element(WP.posOffset) != null;

            if (wrap != null && wrap.Name != WP.wrapTopAndBottom)
            {
                // A column-edge exclusion has an exact CSS float equivalent. An interior
                // object with text on both sides cannot be represented by a single float.
                var edge = (string?)horizontal?.Attribute(NoNamespace.relativeFrom) == "column" &&
                    (horizontal.Element(WP.align)?.Value is "left" or "right" ||
                        (long?)horizontal.Element(WP.posOffset) == 0);
                if (edge && paragraphRelative && (long?)vertical!.Element(WP.posOffset) == 0)
                {
                    var exclusion = ApplyAnchorTextWrap(wrapper, container)!;
                    var exclusionStyle = exclusion.Annotation<Dictionary<string, string>>()!;
                    AddEmuDimensions(exclusionStyle, (long?)container.Element(WP.extent)?.Attribute(NoNamespace.cx),
                        (long?)container.Element(WP.extent)?.Attribute(NoNamespace.cy));
                    exclusionStyle["position"] = "relative";
                    exclusion.SetAttributeValue("data-docx-shape-exclusion", kind);
                    // The float owns the edge position in continuous HTML. Paginated output
                    // resolves the original anchor independently before promoting the SVG.
                    style["left"] = "0pt";
                    style["top"] = options.RenderPagination == PaginationMode.Paginated ? "0pt"
                        : (-topDistance).ToString(CultureInfo.InvariantCulture) + "pt";
                    return exclusion;
                }
                ReportShapeLayoutLoss(wrapper, options, "This drawing's interior text wrapping is approximated by top/bottom clearance");
                wrapper.SetAttributeValue("data-docx-anchor-wrap", "wrapTopAndBottom");
            }
            if (wrap != null && !paragraphRelative)
                ReportShapeLayoutLoss(wrapper, options, "This drawing's page-relative text clearance cannot be resolved in continuous flow");
            return wrapper;
        }

        private static void PrepareRelativeShapeSvg(XElement wrapper, WmlToHtmlConverterSettings settings)
        {
            var relativeWidth = ShapeAnchorNumber(wrapper, "width-percent") > 0;
            var relativeHeight = ShapeAnchorNumber(wrapper, "height-percent") > 0;
            if (!relativeWidth && !relativeHeight) return;
            var svg = wrapper.Element(Svg + "svg");
            var width = ShapeAnchorNumber(wrapper, "extent-width") * 96 / 72;
            var height = ShapeAnchorNumber(wrapper, "extent-height") * 96 / 72;
            if (svg != null && width > 0 && height > 0)
            {
                svg.SetAttributeValue("viewBox", "0 0 " + width.ToString(CultureInfo.InvariantCulture) + " " + height.ToString(CultureInfo.InvariantCulture));
                svg.SetAttributeValue("preserveAspectRatio", "none");
                svg.SetAttributeValue("style", "display: block; width: 100%; height: 100%;");
                foreach (var element in svg.Descendants().Where(e => e.Attribute("stroke-width") != null))
                {
                    var stroke = (double)element.Attribute("stroke-width")!;
                    // Non-scaling strokes ignore the root EMU transform too; store CSS pixels.
                    element.SetAttributeValue("stroke-width", Format(stroke / 9525));
                    element.SetAttributeValue("vector-effect", "non-scaling-stroke");
                }
            }
            else ReportShapeLayoutLoss(wrapper, settings, "This drawing's relative size cannot be applied to its stored geometry");
            if (settings.RenderPagination != PaginationMode.Paginated ||
                relativeWidth && (string?)wrapper.Attribute("data-docx-anchor-width-relative") is not ("page" or "margin") ||
                relativeHeight && (string?)wrapper.Attribute("data-docx-anchor-height-relative") is not ("page" or "margin"))
                ReportShapeLayoutLoss(wrapper, settings, "This drawing's relative sizing needs a physical page or page-dependent margin context");
        }

        private static void ReportShapeLayoutLoss(XElement anchor, WmlToHtmlConverterSettings settings, string message,
            string elementName = "anchor")
        {
            anchor.SetAttributeValue("data-docx-shape", "partial");
            if (settings is { RenderUnsupportedContentPlaceholders: true, IncludeUnsupportedContentMetadata: true })
            {
                anchor.SetAttributeValue("data-content-type", nameof(UnsupportedContentType.Other));
                anchor.SetAttributeValue("data-element-name", elementName);
                anchor.SetAttributeValue("title", message);
            }
        }

        /// <summary>
        /// A floating drawing does not own a text line. Keep its anchor paragraph's mark, and
        /// represent text clearance separately from the SVG. Overlapping top/bottom drawings
        /// reserve their union's bottom, rather than summing their heights in source order.
        /// </summary>
        private static void FinalizeShapeAnchors(XElement html, WmlToHtmlConverterSettings settings)
        {
            var anchors = html.Descendants().Where(e => (string?)e.Attribute("data-docx-shape-anchor") == "true").ToList();
            var clearances = new Dictionary<XElement, decimal>();
            foreach (var anchor in anchors)
            {
                var paragraph = anchor.Ancestors().FirstOrDefault(e => ParagraphElementNames.Contains(e.Name));
                if (paragraph == null) continue;
                var style = anchor.Annotation<Dictionary<string, string>>()!;
                var inFloat = anchor.Parent?.Attribute("data-docx-shape-exclusion") != null;
                if (settings.RenderPagination != PaginationMode.Paginated && !inFloat)
                {
                    var paragraphStyle = paragraph.Annotation<Dictionary<string, string>>();
                    if (paragraphStyle == null)
                    {
                        paragraphStyle = new Dictionary<string, string>();
                        paragraph.AddAnnotation(paragraphStyle);
                    }
                    paragraphStyle["position"] = "relative";
                    var horizontal = (string?)anchor.Attribute("data-docx-anchor-h-relative");
                    var vertical = (string?)anchor.Attribute("data-docx-anchor-v-relative");
                    if (horizontal != "column" || vertical != "paragraph")
                        ReportShapeLayoutLoss(anchor, settings, "This drawing's page or margin position is approximated relative to its paragraph in continuous HTML");
                    if (horizontal == "column" && new[] { "margin-left", "margin-right" }.Any(property =>
                        paragraphStyle.TryGetValue(property, out var value) && value is not ("0" or "0pt" or "0in")))
                        ReportShapeLayoutLoss(anchor, settings, "This drawing's column position is approximated beside an indented paragraph");
                    style["left"] = ShapeAnchorNumber(anchor, "h-offset").ToString(CultureInfo.InvariantCulture) + "pt";
                    style["top"] = ShapeAnchorNumber(anchor, "v-offset").ToString(CultureInfo.InvariantCulture) + "pt";
                    var align = (string?)anchor.Attribute("data-docx-anchor-h-align");
                    if (align == "center") { style["left"] = "50%"; style["transform"] = "translateX(-50%)"; }
                    else if (align == "right") { style.Remove("left"); style["right"] = "0pt"; }
                }
                if ((string?)anchor.Attribute("data-docx-anchor-wrap") == "wrapTopAndBottom")
                {
                    var bottom = ShapeAnchorNumber(anchor, "extent-height") +
                        ShapeAnchorNumber(anchor, "wrap-bottom") +
                        ((string?)anchor.Attribute("data-docx-anchor-v-relative") == "paragraph"
                            ? ShapeAnchorNumber(anchor, "v-offset") : 0);
                    clearances[paragraph] = Math.Max(clearances.GetValueOrDefault(paragraph), bottom);
                }
            }
            foreach (var pair in clearances)
            {
                var spacer = new XElement(Xhtml.span, new XAttribute("data-docx-shape-clearance", "true"),
                    new XAttribute("aria-hidden", "true"), new XAttribute("data-page-map-exclude", "true"));
                spacer.AddAnnotation(new Dictionary<string, string> { { "float", "left" }, { "width", "100%" },
                    { "height", Math.Max(0, pair.Value).ToString(CultureInfo.InvariantCulture) + "pt" } });
                pair.Key.AddFirst(spacer);
            }
            if (settings.RenderPagination != PaginationMode.Paginated && anchors.Count > 0)
            {
                var body = html.Descendants(Xhtml.body).FirstOrDefault();
                if (body != null)
                {
                    var bodyStyle = body.Annotation<Dictionary<string, string>>();
                    if (bodyStyle == null) { bodyStyle = new Dictionary<string, string>(); body.AddAnnotation(bodyStyle); }
                    bodyStyle["isolation"] = "isolate";
                }
                var ordered = anchors.OrderBy(e => ShapeAnchorNumber(e, "order")).ToList();
                for (var index = 0; index < ordered.Count; index++)
                    ordered[index].Annotation<Dictionary<string, string>>()!["z-index"] =
                        ((string?)ordered[index].Attribute("data-docx-anchor-behind") == "true"
                            ? index - ordered.Count : index + 1).ToString(CultureInfo.InvariantCulture);
            }
        }

        private static decimal ShapeAnchorNumber(XElement anchor, string suffix) =>
            decimal.TryParse((string?)anchor.Attribute("data-docx-anchor-" + suffix), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var number) ? number : 0;

        private static bool IsVmlShape(XElement element) => element.Name.Namespace == VML.vml &&
            element.Name.LocalName is "shape" or "rect" or "roundrect" or "oval" or "line" or "polyline" or "arc";

        // A pict is a drawing container, not one text story. Keep its independently authored
        // members in source order and map group coordinates without scaling fonts or outlines.
        private static XElement ProcessVmlMembers(WordprocessingDocument wordDoc, XElement source,
            Func<ImageInfo, XElement?>? imageHandler, WmlToHtmlConverterSettings? settings, ShapeMap? map = null)
        {
            var wrapper = new XElement(Xhtml.span);
            foreach (var child in source.Elements())
            {
                if (child.Name == MC.AlternateContent)
                {
                    if (SelectedMarkupBranch(child) is { } branch)
                        wrapper.Add(ProcessVmlMembers(wordDoc, branch, imageHandler, settings, map));
                    continue;
                }
                if (child.Name != VML.group && !IsVmlShape(child))
                    continue;
                var frame = VmlMemberFrame(child, map);
                XElement? member;
                if (child.Name == VML.group)
                {
                    var origin = VmlCoordinatePair(child, "coordorigin", 0);
                    // VML groups default to a 1000 by 1000 child coordinate space. The
                    // shape element's omitted coordsize rule does not apply to a group.
                    var size = VmlCoordinatePair(child, "coordsize", 1000);
                    ShapeMap? childMap = frame is { } bounds && origin != null && size is { X: > 0, Y: > 0 }
                        ? new ShapeMap(bounds.Width / size.Value.X, bounds.Height / size.Value.Y,
                            -origin.Value.X * bounds.Width / size.Value.X, -origin.Value.Y * bounds.Height / size.Value.Y)
                        : null;
                    member = ProcessVmlMembers(wordDoc, child, imageHandler, settings, childMap);
                    var style = new Dictionary<string, string>
                    {
                        { "display", "inline-block" }, { "position", "relative" },
                        { "vertical-align", "top" }, { "line-height", "normal" },
                    };
                    if (frame is { } groupFrame)
                    {
                        style["width"] = VmlPointLength(groupFrame.Width);
                        style["height"] = VmlPointLength(groupFrame.Height);
                    }
                    if (map == null && settings?.RenderPagination == PaginationMode.Paginated &&
                        CreateVmlDrawingAnchor(child) is { } anchor)
                        AddDrawingFrameStyle(style, member, anchor, settings, false);
                    member.AddAnnotation(style);
                    var sourceStyle = ParseVmlStyle((string?)child.Attribute("style"));
                    var rotation = sourceStyle.GetValueOrDefault("rotation", "0");
                    var unsupported = childMap == null
                        || !double.TryParse(rotation, NumberStyles.Float, CultureInfo.InvariantCulture, out var degrees)
                        || !double.IsFinite(degrees) || degrees % 360 != 0
                        || sourceStyle.GetValueOrDefault("flip", "") is not ("" or "none")
                        || sourceStyle.ContainsKey("center-x") || sourceStyle.ContainsKey("center-y");
                    if (unsupported)
                    {
                        if (settings != null)
                            ReportShapeLayoutLoss(member, settings, "Unsupported VML group coordinate system or transform");
                        // Child frames are approximate too: they cannot prove that source text
                        // was intentionally clipped or placed outside the physical page.
                        member.Descendants().Attributes().Where(a => a.Name.LocalName is
                            "data-docx-textbox-source-clip" or "data-docx-textbox-source-position").Remove();
                    }
                }
                else
                {
                    member = imageHandler == null ? null : ProcessPictureOrObject(wordDoc, child, imageHandler, settings);
                    if (member != null) RecordImageRepresentation(settings, child);
                    member ??= ProcessTextBox(wordDoc, child, settings, map == null ? null : frame)
                        ?? ProcessVmlShape(child, settings);
                }
                if (member == null) continue;
                if (map != null && frame is { } position)
                {
                    var style = member.Annotation<Dictionary<string, string>>();
                    if (style == null)
                    {
                        style = ParseVmlStyle((string?)member.Attribute("style"));
                        member.Attribute("style")?.Remove();
                        member.AddAnnotation(style);
                    }
                    style["position"] = "absolute";
                    style["left"] = VmlPointLength(position.X);
                    style["top"] = VmlPointLength(position.Y);
                    style["width"] = VmlPointLength(position.Width);
                    if (!HasVmlTextBoxAutoFit(child)) style["height"] = VmlPointLength(position.Height);
                }
                wrapper.Add(member);
            }
            return wrapper;
        }

        private static string VmlPointLength(double value) => value.ToString("0.####", CultureInfo.InvariantCulture) + "pt";

        private static (double X, double Y)? VmlCoordinatePair(XElement element, string attribute, double fallback)
        {
            var value = (string?)element.Attribute(attribute);
            if (value == null) return (fallback, fallback);
            var pair = value.Split(new[] { ',', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            return pair.Length == 2 && double.TryParse(pair[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var x)
                && double.TryParse(pair[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var y)
                && double.IsFinite(x) && double.IsFinite(y) ? (x, y) : null;
        }

        private static ShapeBox? VmlMemberFrame(XElement element, ShapeMap? map)
        {
            var style = ParseVmlStyle((string?)element.Attribute("style"));
            double? Coordinate(string name, double scale, bool optional = false)
            {
                var value = style.GetValueOrDefault(name);
                if (value == null) return optional ? 0 : null;
                if (TryVmlPointEmu(value, out var emu)) return emu / 12700d * scale;
                return map != null && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var units)
                    && double.IsFinite(units * scale) ? units * scale : null;
            }
            var width = Coordinate("width", map?.ScaleX ?? 1);
            var height = Coordinate("height", map?.ScaleY ?? 1);
            var left = Coordinate("left", map?.ScaleX ?? 1, true);
            var top = Coordinate("top", map?.ScaleY ?? 1, true);
            return width is > 0 && height is > 0 && left != null && top != null
                ? new ShapeBox(left.Value + (map?.OffsetX ?? 0), top.Value + (map?.OffsetY ?? 0), width.Value, height.Value)
                : null;
        }

        private static XElement? ProcessVmlShape(XElement element, WmlToHtmlConverterSettings? settings)
        {
            // Image and textbox paths already had the opportunity to render their content.
            // Do not turn an explicitly omitted picture into an unsupported shape.
            if (element.Descendants(VML.imagedata).Any()) return null;
            var shapes = element.DescendantsAndSelf().Where(IsVmlShape).ToList();
            var shape = shapes.FirstOrDefault();
            if (shape == null) return null;
            if (shapes.Count == 1 && (shape == element || shape.Parent == element) && TryVmlRectangle(shape) is { } rectangle)
                return rectangle;
            var placeholder = (settings == null ? null
                : CreateUnsupportedContentPlaceholder(settings, shape, UnsupportedContentType.Other, "[SHAPE]"))
                ?? new XElement(Xhtml.span, new XAttribute("aria-hidden", "true"));
            placeholder.SetAttributeValue("data-docx-shape", "unrendered");
            var style = new Dictionary<string, string> { { "display", "inline-block" },
                { "box-sizing", "border-box" }, { "font-family", "inherit" } };
            // Keeping an undrawn oval or line's extent must not invent a filled rectangle.
            AddVmlDimensions(style, shape);
            placeholder.AddAnnotation(style);
            if (settings?.RenderPagination == PaginationMode.Paginated && shapes.Count == 1 && (shape == element || shape.Parent == element) &&
                CreateVmlDrawingAnchor(shape) is XElement anchor)
                return PlaceDrawingShape(placeholder, anchor, settings);
            return placeholder;
        }

        private static XElement? TryVmlRectangle(XElement shape)
        {
            if (shape.Name.LocalName != "rect" || shape.HasElements || shape.Attributes().Any(a =>
                !a.IsNamespaceDeclaration && a.Name.Namespace != PtOpenXml.pt && !(a.Name.Namespace == XNamespace.None && (a.Name.LocalName is
                    "style" or "id" or "title" or "alt" or "fillcolor" or "filled" or "strokecolor" or "stroked" or "strokeweight") ||
                    a.Name.NamespaceName == "urn:schemas-microsoft-com:office:office" && a.Name.LocalName == "spid")))
                return null;
            var tokens = ((string?)shape.Attribute("style") ?? "").Split(';')
                .Select(token => token.Trim()).Where(token => token.Length > 0).ToList();
            // Positioned/grouped VML and effects need more than an inline rectangle. Preserve
            // their unsupported fallback rather than silently discarding those instructions.
            if (tokens.Any(token => !token.StartsWith("width:", StringComparison.Ordinal) &&
                !token.StartsWith("height:", StringComparison.Ordinal))) return null;
            var width = WidthInPoints(tokens);
            var height = HeightInPoints(tokens);
            if (width is not > 0 || height is not > 0 || !float.IsFinite(width.Value) || !float.IsFinite(height.Value) ||
                width.Value > long.MaxValue / 12700d || height.Value > long.MaxValue / 12700d)
                return null;
            var filled = VmlFlag(shape, "filled");
            var stroked = VmlFlag(shape, "stroked");
            var fill = NormalizeVmlCssColor((string?)shape.Attribute("fillcolor") ?? "#ffffff");
            var stroke = NormalizeVmlCssColor((string?)shape.Attribute("strokecolor") ?? "#000000");
            var weight = (string?)shape.Attribute("strokeweight");
            // VML's default is one CSS pixel. A unitless attribute is measured in EMUs.
            double? strokeWidth = weight == null ? 9525 :
                double.TryParse(weight, NumberStyles.Float, CultureInfo.InvariantCulture, out var emu) ? emu :
                weight.EndsWith("px", StringComparison.Ordinal) && double.TryParse(weight[..^2], NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var pixels) ? pixels * 9525 :
                WidthInPoints(new[] { "width:" + weight }) is { } points ? points * 12700d : null;
            if (filled == null || stroked == null || filled == true && fill == null ||
                stroked == true && (stroke == null || strokeWidth is not >= 0 || !double.IsFinite(strokeWidth.Value)))
                return null;
            var widthEmu = (long)Math.Round(width.Value * 12700d);
            var heightEmu = (long)Math.Round(height.Value * 12700d);
            var rect = ShapeRectangle(new ShapeBox(0, 0, widthEmu, heightEmu));
            rect.Add(new XAttribute("fill", filled == true ? fill! : "none"),
                new XAttribute("stroke", stroked == true ? stroke! : "none"));
            if (stroked == true) rect.Add(new XAttribute("stroke-width", Format(strokeWidth!.Value)));
            var space = DrawingSpace();
            space.Add(rect);
            var svg = DrawingSvg(widthEmu, heightEmu, space);
            var alt = (string?)shape.Attribute("alt") ?? (string?)shape.Attribute("title");
            if (string.IsNullOrEmpty(alt)) svg.SetAttributeValue("aria-hidden", "true");
            else svg.Add(new XAttribute("role", "img"), new XAttribute("aria-label", alt));
            var wrapper = new XElement(Xhtml.span, svg);
            var style = new Dictionary<string, string> { { "display", "inline-block" } };
            AddEmuDimensions(style, widthEmu, heightEmu);
            wrapper.AddAnnotation(style);
            return wrapper;
        }

        private static bool? VmlFlag(XElement shape, string name) => ((string?)shape.Attribute(name))?.ToLowerInvariant() switch
        {
            null or "t" or "true" or "1" or "on" => true,
            "f" or "false" or "0" or "off" => false,
            _ => null,
        };

        private static XElement DrawingSpace() => new XElement(Svg + "g", new XAttribute("transform",
            "scale(" + (1.0 / 9525).ToString("G9", CultureInfo.InvariantCulture) + ")"));

        /// <summary>
        /// The SVG for a drawing's extent. It has no viewBox: a zero extent (a horizontal or vertical
        /// line) would make one degenerate, which disables rendering. Its explicit dimensions, at
        /// least a pixel each, are what the PDF export requires of an inline SVG before it will lay
        /// one out; overflow keeps outlines whole.
        /// </summary>
        private static XElement DrawingSvg(long width, long height, object? content) =>
            new XElement(Svg + "svg",
                new XAttribute("width", Format(Math.Max(width / 9525.0, 1))),
                new XAttribute("height", Format(Math.Max(height / 9525.0, 1))),
                new XAttribute("overflow", "visible"),
                new XAttribute("style", "display: block;"),
                content);

        /// <summary>
        /// An inline drawing nothing could draw, as a box of its extent. With unsupported-content
        /// placeholders on (as the PDF export always converts), it is the standard placeholder at
        /// that size, so the export reports it, and rejects it under its strict policy. Otherwise
        /// it is a blank box marked <c>data-docx-shape="unrendered"</c> and hidden from assistive
        /// technology, which has nothing to announce.
        /// </summary>
        private static XElement UnrenderedDrawing(WmlToHtmlConverterSettings? settings, XElement drawing,
            XElement? content, long width, long height)
        {
            var style = new Dictionary<string, string> { { "display", "inline-block" } };
            AddEmuDimensions(style, width, height);
            var uri = (string?)content?.Parent?.Attribute("uri") ?? "";
            var label = content?.Name == Pic._pic ? "[UNSUPPORTED IMAGE]"
                : content?.Name == C.chart ? "[CHART]"
                : uri.EndsWith("/diagram", StringComparison.Ordinal) ? "[SMARTART]"
                : content?.Name == WPS.wsp || content?.Name == WpgGroup || content?.Name == WpcCanvas ? "[SHAPE]"
                : "[DRAWING]";
            var placeholder = settings == null ? null
                : CreateUnsupportedContentPlaceholder(settings, content ?? drawing, UnsupportedContentType.Other, label);
            if (placeholder != null)
            {
                placeholder.Add(new XAttribute("data-docx-shape", "unrendered"));
                // A drawing label is UI, not a new document font. Introducing generic monospace
                // would make strict-font export fail before its content policy can report the loss.
                style["font-family"] = "inherit";
                style["box-sizing"] = "border-box";
                style["overflow"] = "hidden";
                placeholder.AddAnnotation(style);
                return placeholder;
            }

            var box = new XElement(Xhtml.span,
                DrawingSvg(width, height, null),
                new XAttribute("data-docx-shape", "unrendered"),
                new XAttribute("aria-hidden", "true"));
            box.AddAnnotation(style);
            return box;
        }

        private readonly record struct ShapeBox(double X, double Y, double Width, double Height);

        /// <summary>
        /// A group's child coordinate space mapped onto the drawing's: scale, then translate. It is
        /// applied to geometry rather than emitted as an SVG transform, so a group whose child
        /// space is not the size of its frame does not scale its children's outlines (Word and
        /// LibreOffice keep a grouped 1pt line at 1pt).
        /// </summary>
        private readonly record struct ShapeMap(double ScaleX, double ScaleY, double OffsetX, double OffsetY)
        {
            public static ShapeMap Identity => new(1, 1, 0, 0);

            public ShapeBox Apply(ShapeBox box) =>
                new(box.X * ScaleX + OffsetX, box.Y * ScaleY + OffsetY, box.Width * ScaleX, box.Height * ScaleY);
        }

        /// <summary>The document's theme, read once per drawing and only if a color names it.</summary>
        private sealed class ShapeContext
        {
            private readonly WordprocessingDocument _wordDoc;
            private ThemeColorScheme? _theme;

            public ShapeContext(WordprocessingDocument wordDoc) => _wordDoc = wordDoc;

            public ThemeColorScheme Theme => _theme ??= LoadThemeColorScheme(_wordDoc);

            /// <summary>Shapes drawn, and drawing parts (geometry, pictures, text) left undrawn.</summary>
            public int Drawn;
            public int Undrawn;

            /// <summary>Width of the theme line style an <c>a:lnRef</c> index names, in EMUs.</summary>
            public long? ThemeLineWidth(int index) =>
                (long?)_wordDoc.MainDocumentPart?.ThemePart?.GetXDocument().Root?
                    .Element(A.themeElements)?.Element(A.fmtScheme)?.Element(A.lnStyleLst)?
                    .Elements(A.ln).ElementAtOrDefault(index - 1)?.Attribute(NoNamespace.w);
        }

        /// <summary>
        /// Adds a group whose frame in drawing coordinates is <paramref name="frame"/>: its children
        /// are mapped from its child coordinate space (<c>a:chOff</c>/<c>a:chExt</c>) onto that
        /// frame, and the group's own rotation and flips wrap them.
        /// </summary>
        private static void AddGroup(XElement target, XElement group, XElement? groupProperties, ShapeBox frame,
            XElement? inheritedFill, ShapeContext context)
        {
            var transform = groupProperties?.Element(A.xfrm);
            var childOffset = transform?.Element(A.chOff);
            var childExtent = transform?.Element(A.chExt);
            var childX = (double?)(long?)childOffset?.Attribute(NoNamespace.x) ?? 0;
            var childY = (double?)(long?)childOffset?.Attribute(NoNamespace.y) ?? 0;
            var childWidth = (double?)(long?)childExtent?.Attribute(NoNamespace.cx) ?? frame.Width;
            var childHeight = (double?)(long?)childExtent?.Attribute(NoNamespace.cy) ?? frame.Height;
            // An axis with no child extent (a group of only horizontal lines, say) is not scaled.
            var scaleX = childWidth > 0 ? frame.Width / childWidth : 1;
            var scaleY = childHeight > 0 ? frame.Height / childHeight : 1;
            var map = new ShapeMap(scaleX, scaleY, frame.X - childX * scaleX, frame.Y - childY * scaleY);

            // A child with a:grpFill takes the nearest group fill that is not itself a:grpFill.
            var groupFill = groupProperties?.Elements().FirstOrDefault(IsFillChoice);
            if (groupFill == null || groupFill.Name == A.grpFill)
                groupFill = inheritedFill;

            var orientation = ShapeOrientation(transform, frame);
            var children = orientation == null ? target : new XElement(Svg + "g", new XAttribute("transform", orientation));
            AddGroupChildren(children, group, map, groupFill, context);
            if (children != target)
                target.Add(children);
        }

        /// <summary>Adds the shapes and nested groups of a group or canvas, in z-order.</summary>
        private static void AddGroupChildren(XElement target, XElement group, ShapeMap map, XElement? groupFill,
            ShapeContext context)
        {
            foreach (var child in group.Elements())
            {
                if (child.Name == WPS.wsp)
                {
                    var frame = ChildFrame(child.Element(WPS.spPr)?.Element(A.xfrm));
                    if (frame == null)
                        context.Undrawn++;
                    else
                        AddShapeElement(target, child, map.Apply(frame.Value), groupFill, context);
                }
                else if (child.Name == WpgNestedGroup || child.Name == WpgGroup)
                {
                    var properties = child.Element(WpgGroupProperties);
                    var frame = ChildFrame(properties?.Element(A.xfrm));
                    if (frame == null)
                        context.Undrawn++;
                    else
                        AddGroup(target, child, properties, map.Apply(frame.Value), groupFill, context);
                }
                else if (child.Name == Pic._pic || child.Name == WPG.wpg + "graphicFrame" ||
                         child.Name == WPC.wpc + "graphicFrame" || child.Name == W14.w14 + "contentPart")
                {
                    // Pictures, charts and ink inside a group are not drawn.
                    context.Undrawn++;
                }
            }
        }

        private static ShapeBox? ChildFrame(XElement? transform)
        {
            var offset = transform?.Element(A.off);
            var size = transform?.Element(A.ext);
            if (size == null)
                return null;
            return new ShapeBox(
                (long?)offset?.Attribute(NoNamespace.x) ?? 0,
                (long?)offset?.Attribute(NoNamespace.y) ?? 0,
                Math.Max(0L, (long?)size.Attribute(NoNamespace.cx) ?? 0L),
                Math.Max(0L, (long?)size.Attribute(NoNamespace.cy) ?? 0L));
        }

        /// <summary>
        /// Draws one shape's preset geometry in <paramref name="frame"/>, in drawing coordinates.
        /// Other geometry, and the text of a text box inside a group, count as undrawn.
        /// </summary>
        private static void AddShapeElement(XElement target, XElement shape, ShapeBox frame, XElement? groupFill,
            ShapeContext context)
        {
            var properties = shape.Element(WPS.spPr);
            var shapeStyle = shape.Element(WPS.style);
            var geometry = (string?)properties?.Element(A.prstGeom)?.Attribute(NoNamespace.prst);
            XElement? element = geometry switch
            {
                "rect" => ShapeRectangle(frame),
                "roundRect" => RoundedRectangle(properties!.Element(A.prstGeom)!, frame),
                "hexagon" => Hexagon(properties!, frame),
                "ellipse" => new XElement(Svg + "ellipse",
                    new XAttribute("cx", Format(frame.X + frame.Width / 2)),
                    new XAttribute("cy", Format(frame.Y + frame.Height / 2)),
                    new XAttribute("rx", Format(frame.Width / 2)), new XAttribute("ry", Format(frame.Height / 2))),
                "line" or "straightConnector1" => new XElement(Svg + "line",
                    new XAttribute("x1", Format(frame.X)), new XAttribute("y1", Format(frame.Y)),
                    new XAttribute("x2", Format(frame.X + frame.Width)),
                    new XAttribute("y2", Format(frame.Y + frame.Height))),
                _ => null,
            };
            if (shape.Descendants(W.txbxContent).Any(text => text.Descendants(W.t).Any()))
                context.Undrawn++;
            if (element == null)
            {
                context.Undrawn++;
                return;
            }

            context.Drawn++;
            var isLine = element.Name.LocalName == "line";
            var fill = isLine ? null : ShapeFill(properties, shapeStyle, groupFill, context);
            element.Add(new XAttribute("fill", fill?.Color ?? "none"));
            if (fill?.Opacity is { } fillOpacity)
                element.Add(new XAttribute("fill-opacity", Format(fillOpacity)));
            var (stroke, strokeWidth) = ShapeOutline(properties, shapeStyle, groupFill, context);
            element.Add(new XAttribute("stroke", stroke?.Color ?? "none"));
            if (stroke != null)
            {
                element.Add(new XAttribute("stroke-width", Format(strokeWidth)));
                if (stroke.Value.Opacity is { } strokeOpacity)
                    element.Add(new XAttribute("stroke-opacity", Format(strokeOpacity)));
            }
            var orientation = ShapeOrientation(properties?.Element(A.xfrm), frame);
            if (orientation != null)
                element.Add(new XAttribute("transform", orientation));
            target.Add(element);
        }

        private static XElement RoundedRectangle(XElement presetGeometry, ShapeBox frame)
        {
            // ECMA-376 presetShapeDefinitions: the corner radius is adj/100000 of the shorter
            // side, adj defaulting to 16667 and clamped to [0, 50000].
            var adjustment = PresetAdjustment(presetGeometry, "adj", 16667) ?? 16667;
            var radius = Math.Min(frame.Width, frame.Height) * Math.Clamp(adjustment, 0, 50000) / 100000;
            return new XElement(Svg + "rect",
                new XAttribute("x", Format(frame.X)), new XAttribute("y", Format(frame.Y)),
                new XAttribute("width", Format(frame.Width)), new XAttribute("height", Format(frame.Height)),
                new XAttribute("rx", Format(radius)), new XAttribute("ry", Format(radius)));
        }

        private static XElement ShapeRectangle(ShapeBox frame) => new XElement(Svg + "rect",
            new XAttribute("x", Format(frame.X)), new XAttribute("y", Format(frame.Y)),
            new XAttribute("width", Format(frame.Width)), new XAttribute("height", Format(frame.Height)));

        private static double? PresetAdjustment(XElement preset, string name, double fallback)
        {
            var formula = (string?)preset.Element(A.avLst)?.Elements(A.gd)
                .FirstOrDefault(guide => (string?)guide.Attribute(NoNamespace.name) == name)?.Attribute(NoNamespace.fmla);
            if (formula == null) return fallback;
            return formula.StartsWith("val ", StringComparison.Ordinal) &&
                double.TryParse(formula[4..], NumberStyles.Float, CultureInfo.InvariantCulture, out var value) && double.IsFinite(value)
                ? value : null;
        }

        private static XElement? Hexagon(XElement properties, ShapeBox frame)
        {
            var preset = properties.Element(A.prstGeom)!;
            var adjustment = PresetAdjustment(preset, "adj", 25000);
            var verticalFactor = PresetAdjustment(preset, "vf", 115470);
            if (adjustment == null || verticalFactor == null) return null;
            // ECMA presetShapeDefinitions: x1=ss*pin(0,adj,50000*w/ss)/100000;
            // y1=vc-sin(hd2*vf/100000,60deg). Map source geometry before group scaling.
            var sourceWidth = (long?)properties.Element(A.xfrm)?.Element(A.ext)?.Attribute(NoNamespace.cx) ?? frame.Width;
            var sourceHeight = (long?)properties.Element(A.xfrm)?.Element(A.ext)?.Attribute(NoNamespace.cy) ?? frame.Height;
            var inset = sourceWidth > 0 ? Math.Clamp(Math.Min(sourceWidth, sourceHeight) * adjustment.Value / 100000,
                0, sourceWidth / 2) * frame.Width / sourceWidth : 0;
            var deltaY = frame.Height / 2 * verticalFactor.Value / 100000 * Math.Sin(Math.PI / 3);
            if (!double.IsFinite(deltaY)) return null;
            var centerY = frame.Y + frame.Height / 2;
            var points = new[] { (frame.X, centerY), (frame.X + inset, centerY - deltaY),
                (frame.X + frame.Width - inset, centerY - deltaY), (frame.X + frame.Width, centerY),
                (frame.X + frame.Width - inset, centerY + deltaY), (frame.X + inset, centerY + deltaY) };
            return new XElement(Svg + "polygon", new XAttribute("points",
                string.Join(" ", points.Select(point => Format(point.Item1) + "," + Format(point.Item2)))));
        }

        /// <summary>
        /// The SVG transform for an <c>a:xfrm</c>'s flips and rotation about the frame's center,
        /// or null when it has neither. DrawingML flips a shape within its frame before rotating it.
        /// </summary>
        private static string? ShapeOrientation(XElement? transform, ShapeBox frame)
        {
            var degrees = ((long?)transform?.Attribute(NoNamespace.rot) ?? 0) / 60000.0;
            var flipX = IsTrueAttribute(transform, "flipH") ? -1 : 1;
            var flipY = IsTrueAttribute(transform, "flipV") ? -1 : 1;
            if (degrees == 0 && flipX == 1 && flipY == 1)
                return null;

            var centerX = frame.X + frame.Width / 2;
            var centerY = frame.Y + frame.Height / 2;
            var parts = new List<string>();
            if (degrees != 0)
                parts.Add($"rotate({Format(degrees)} {Format(centerX)} {Format(centerY)})");
            if (flipX == -1 || flipY == -1)
                parts.Add($"matrix({flipX} 0 0 {flipY} {Format(centerX * (1 - flipX))} {Format(centerY * (1 - flipY))})");
            return string.Join(" ", parts);
        }

        private static bool IsTrueAttribute(XElement? element, string name) =>
            (string?)element?.Attribute(name) is "1" or "true";

        /// <summary>A paint: a CSS color and, when the DrawingML color is translucent, its opacity.</summary>
        private readonly record struct ShapePaint(string Color, double? Opacity);

        /// <summary>A shape's fill, or null when it is unfilled.</summary>
        private static ShapePaint? ShapeFill(XElement? properties, XElement? shapeStyle, XElement? groupFill,
            ShapeContext context)
        {
            var fill = properties?.Elements().FirstOrDefault(IsFillChoice);
            if (fill != null)
                return FillPaint(fill, groupFill, context);

            // No explicit fill: the shape style's fill reference applies; index 0 means no fill.
            var reference = shapeStyle?.Element(A.fillRef);
            return ReadIndex(reference) > 0 ? Paint(reference, context) : null;
        }

        /// <summary>A shape's outline and its width in EMUs; a null paint means no outline.</summary>
        private static (ShapePaint? Paint, double Width) ShapeOutline(XElement? properties, XElement? shapeStyle,
            XElement? groupFill, ShapeContext context)
        {
            var line = properties?.Element(A.ln);
            var reference = shapeStyle?.Element(A.lnRef);
            var referenceIndex = ReadIndex(reference);
            var fill = line?.Elements().FirstOrDefault(IsFillChoice);
            var paint = fill != null
                ? FillPaint(fill, groupFill, context)
                : referenceIndex > 0 ? Paint(reference, context) : null;
            // An outline with no width of its own takes the referenced theme line style's width;
            // otherwise it is DrawingML's thinnest line, drawn here as one CSS pixel (9525 EMUs).
            var width = (long?)line?.Attribute(NoNamespace.w)
                ?? (referenceIndex > 0 ? context.ThemeLineWidth(referenceIndex) : null)
                ?? 9525L;
            return (paint, Math.Max(width, 9525L));
        }

        private static bool IsFillChoice(XElement element) =>
            element.Name == A.noFill || element.Name == A.solidFill || element.Name == A.gradFill ||
            element.Name == A.pattFill || element.Name == A.blipFill || element.Name == A.grpFill;

        /// <summary>
        /// The paint of a fill choice. A gradient approximates to its first stop, a pattern to its
        /// foreground and a group fill to the enclosing group's; a picture fill is not drawn.
        /// </summary>
        private static ShapePaint? FillPaint(XElement fill, XElement? groupFill, ShapeContext context)
        {
            if (fill.Name == A.grpFill)
                return groupFill == null || groupFill.Name == A.grpFill ? null : FillPaint(groupFill, null, context);
            if (fill.Name == A.solidFill)
                return Paint(fill, context);
            if (fill.Name == A.gradFill)
                return Paint(fill.Element(A.gsLst)?.Elements(A.gs).FirstOrDefault(), context);
            if (fill.Name == A.pattFill)
                return Paint(fill.Element(A.fgClr), context);
            if (fill.Name == A.blipFill)
                context.Undrawn++;
            return null;
        }

        private static ShapePaint? Paint(XElement? colorParent, ShapeContext context)
        {
            var color = ResolveDrawingColor(colorParent, context.Theme);
            if (color == null)
            {
                if (colorParent?.HasElements == true)
                    context.Undrawn++;
                return null;
            }
            var alpha = ReadDoubleAttribute(colorParent!.Elements().First().Element(A.alpha), "val") / 100000;
            return new ShapePaint(color, alpha is < 1 ? Math.Max(alpha.Value, 0) : null);
        }

        private static int ReadIndex(XElement? reference) =>
            int.TryParse((string?)reference?.Attribute(NoNamespace.idx), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var index) ? index : 0;

        /// <summary>
        /// Resolves the DrawingML color element under <paramref name="colorParent"/> — an sRGB,
        /// theme or system color, with its luminance, shade and tint modifiers — to a CSS
        /// <c>#RRGGBB</c> value, or null when it names nothing this converter can resolve.
        /// </summary>
        private static string? ResolveDrawingColor(XElement? colorParent, ThemeColorScheme? theme)
        {
            var color = colorParent?.Elements().FirstOrDefault(e => e.Name == A.srgbClr || e.Name == A.schemeClr ||
                e.Name == A.sysClr || e.Name == A.scrgbClr || e.Name == A.hslClr);
            if (color == null)
                return null;

            string? rgb = null;
            if (color.Name == A.srgbClr)
                rgb = (string?)color.Attribute(NoNamespace.val);
            else if (color.Name == A.scrgbClr)
            {
                // Linear RGB in thousandths of a percent.
                string Channel(string name) => ((int)Math.Round(FromLinear(
                    (ReadDoubleAttribute(color, name) ?? 0) / 100000))).ToString("X2", CultureInfo.InvariantCulture);
                rgb = Channel("r") + Channel("g") + Channel("b");
            }
            else if (color.Name == A.hslClr)
                rgb = HslToRgb((ReadDoubleAttribute(color, "hue") ?? 0) / 60000,
                    (ReadDoubleAttribute(color, "sat") ?? 0) / 100000, (ReadDoubleAttribute(color, "lum") ?? 0) / 100000);
            else if (color.Name == A.sysClr)
                rgb = (string?)color.Attribute("lastClr")
                    ?? ((string?)color.Attribute(NoNamespace.val) is { } system ? MapSystemColor(system) : null);
            else if ((string?)color.Attribute(NoNamespace.val) is { } name)
            {
                // Word's default color mapping (w:clrSchemeMapping) sends text/background slots
                // to the theme's dark/light colors.
                name = name switch { "tx1" => "dk1", "bg1" => "lt1", "tx2" => "dk2", "bg2" => "lt2", _ => name };
                theme?.Colors.TryGetValue(name, out rgb);
            }

            return IsHexColor(rgb) ? "#" + ApplyDrawingColorModifiers(rgb, color).ToUpperInvariant() : null;
        }

        private static string ApplyDrawingColorModifiers(string color, XElement colorElement)
        {
            var channels = new[]
            {
                int.Parse(color.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                int.Parse(color.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                int.Parse(color.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            }.Select(channel => (double)channel).ToArray();
            double? Modifier(XName name) => ReadDoubleAttribute(colorElement.Element(name), "val") / 100000;

            // Shade darkens toward black and tint lightens toward white, both in linear RGB as
            // Word applies them (accent1 4472C4 at 50% shade is Word's familiar 2F528F outline).
            if (Modifier(A.shade) is { } shade)
                channels = channels.Select(channel => FromLinear(ToLinear(channel) * shade)).ToArray();
            if (Modifier(A.tint) is { } tint)
                channels = channels.Select(channel => FromLinear(1 - (1 - ToLinear(channel)) * tint)).ToArray();
            // Luminance modulation and offset follow, approximated per channel as the chart
            // projection always has.
            var lumMod = Modifier(A.lumMod) ?? 1;
            var lumOff = Modifier(A.lumOff) ?? 0;
            return string.Concat(channels.Select(channel =>
                ((int)Math.Round(Math.Clamp(channel * lumMod + 255 * lumOff, 0, 255))).ToString("X2", CultureInfo.InvariantCulture)));
        }

        private static string HslToRgb(double hue, double saturation, double luminance)
        {
            saturation = Math.Clamp(saturation, 0, 1);
            luminance = Math.Clamp(luminance, 0, 1);
            var chroma = (1 - Math.Abs(2 * luminance - 1)) * saturation;
            var sector = (hue % 360 + 360) % 360 / 60;
            var second = chroma * (1 - Math.Abs(sector % 2 - 1));
            var (r, g, b) = (int)sector switch
            {
                0 => (chroma, second, 0.0),
                1 => (second, chroma, 0.0),
                2 => (0.0, chroma, second),
                3 => (0.0, second, chroma),
                4 => (second, 0.0, chroma),
                _ => (chroma, 0.0, second),
            };
            var lightness = luminance - chroma / 2;
            return string.Concat(new[] { r, g, b }.Select(channel =>
                ((int)Math.Round((channel + lightness) * 255)).ToString("X2", CultureInfo.InvariantCulture)));
        }

        private static double ToLinear(double channel)
        {
            var value = channel / 255;
            return value <= 0.04045 ? value / 12.92 : Math.Pow((value + 0.055) / 1.055, 2.4);
        }

        private static double FromLinear(double linear)
        {
            linear = Math.Clamp(linear, 0, 1);
            return 255 * (linear <= 0.0031308 ? linear * 12.92 : 1.055 * Math.Pow(linear, 1 / 2.4) - 0.055);
        }
    }
}
