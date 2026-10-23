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
        /// Draws an inline DrawingML drawing that is neither a picture, a chart nor a text box — a
        /// <c>wps:wsp</c> shape, a <c>wpg:wgp</c> group or a <c>wpc:wpc</c> canvas — as SVG sized
        /// to its <c>wp:extent</c>. Rectangles, rounded rectangles, ellipses and lines are drawn
        /// with their fill and outline; a group maps its children through its child coordinate
        /// space (<c>a:chOff</c>/<c>a:chExt</c>). Geometry this projection does not draw still
        /// occupies its extent.
        /// <para>
        /// Any other inline drawing nothing else rendered also gets an empty box of its extent. The
        /// empty-paragraph pass counts every inline <c>w:drawing</c> as paragraph content, so a
        /// drawing that rendered to nothing left its paragraph with no line box at all (issue
        /// #1270). The one exception is a picture when the caller passed no image handler: that
        /// caller asked for pictures to be left out.
        /// </para>
        /// <para>
        /// Floating (<c>wp:anchor</c>) shapes are not drawn. They are not counted as paragraph
        /// content, and drawing them faithfully needs their stacking (<c>behindDoc</c>) and text
        /// wrapping, which this projection does not model.
        /// </para>
        /// </summary>
        private static XElement? ProcessDrawingShape(WordprocessingDocument wordDoc, XElement drawing,
            bool rendersPictures)
        {
            var container = drawing.Element(WP.inline);
            if (container == null)
                return null;

            var graphicData = container.Element(A.graphic)?.Element(A.graphicData);
            var payload = graphicData?.Elements()
                .FirstOrDefault(e => e.Name == WPS.wsp || e.Name == WpgGroup || e.Name == WpcCanvas);
            if (payload == null && !rendersPictures && graphicData?.Element(Pic._pic) != null)
                return null;

            var extent = container.Element(WP.extent);
            var width = Math.Max(0L, (long?)extent?.Attribute(NoNamespace.cx) ?? 0L);
            var height = Math.Max(0L, (long?)extent?.Attribute(NoNamespace.cy) ?? 0L);
            var docPr = container.Element(WP.docPr);
            var altText = (string?)docPr?.Attribute(NoNamespace.descr) ?? (string?)docPr?.Attribute(NoNamespace.name) ?? "";

            // The SVG has no viewBox: a zero extent (a horizontal or vertical line) would make one
            // degenerate, which disables rendering. Drawing coordinates are EMUs, scaled to the
            // SVG's CSS pixels (9525 EMUs each) by the root group; overflow keeps outlines whole.
            var svg = new XElement(Svg + "svg",
                new XAttribute("overflow", "visible"),
                new XAttribute("role", "img"),
                new XAttribute("aria-label", altText),
                new XAttribute("style", "display: block; width: 100%; height: 100%;"));
            if (payload == null)
                svg.Add(new XAttribute("data-docx-shape", "unrendered"));
            else
            {
                var context = new ShapeContext(wordDoc);
                var frame = new ShapeBox(0, 0, width, height);
                var drawingSpace = new XElement(Svg + "g",
                    new XAttribute("transform", "scale(" + (1.0 / 9525).ToString("G9", CultureInfo.InvariantCulture) + ")"));
                if (payload.Name == WPS.wsp)
                    AddShapeElement(drawingSpace, payload, frame, context);
                else if (payload.Name == WpgGroup)
                    drawingSpace.Add(GroupElement(payload, payload.Element(WpgGroupProperties), frame, context));
                else
                    AddGroupChildren(drawingSpace, payload, context);
                svg.Add(drawingSpace);
            }

            // An inline-block whose only content is a block-level SVG has no line box, so its
            // baseline is its bottom edge: the drawing sits on the text baseline, as Word places
            // an inline object.
            var wrapper = new XElement(Xhtml.span, svg);
            var style = new Dictionary<string, string> { { "display", "inline-block" } };
            AddEmuDimensions(style, width, height);
            wrapper.AddAnnotation(style);
            return wrapper;
        }

        private readonly record struct ShapeBox(double X, double Y, double Width, double Height);

        /// <summary>The document's theme, read once per drawing and only if a color names it.</summary>
        private sealed class ShapeContext
        {
            private readonly WordprocessingDocument _wordDoc;
            private ThemeColorScheme? _theme;

            public ShapeContext(WordprocessingDocument wordDoc) => _wordDoc = wordDoc;

            public ThemeColorScheme Theme => _theme ??= LoadThemeColorScheme(_wordDoc);

            /// <summary>Width of the theme line style an <c>a:lnRef</c> index names, in EMUs.</summary>
            public long? ThemeLineWidth(int index) =>
                (long?)_wordDoc.MainDocumentPart?.ThemePart?.GetXDocument().Root?
                    .Element(A.themeElements)?.Element(A.fmtScheme)?.Element(A.lnStyleLst)?
                    .Elements(A.ln).ElementAtOrDefault(index - 1)?.Attribute(NoNamespace.w);
        }

        /// <summary>
        /// A group as an SVG group whose transform maps its child coordinate space
        /// (<c>a:chOff</c>/<c>a:chExt</c>) onto its frame in the parent's coordinates, so each
        /// child's own <c>a:xfrm</c> lands where the group puts it.
        /// </summary>
        private static XElement GroupElement(XElement group, XElement? groupProperties, ShapeBox frame,
            ShapeContext context)
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

            var children = new XElement(Svg + "g", new XAttribute("transform",
                $"matrix({Format(scaleX)} 0 0 {Format(scaleY)} {Format(frame.X - childX * scaleX)} {Format(frame.Y - childY * scaleY)})"));
            AddGroupChildren(children, group, context);

            var orientation = ShapeOrientation(transform, frame);
            return orientation == null
                ? children
                : new XElement(Svg + "g", new XAttribute("transform", orientation), children);
        }

        /// <summary>Adds the shapes and nested groups of a group or canvas, in z-order.</summary>
        private static void AddGroupChildren(XElement target, XElement group, ShapeContext context)
        {
            foreach (var child in group.Elements())
            {
                if (child.Name == WPS.wsp)
                {
                    var frame = ChildFrame(child.Element(WPS.spPr)?.Element(A.xfrm));
                    if (frame != null)
                        AddShapeElement(target, child, frame.Value, context);
                }
                else if (child.Name == WpgNestedGroup || child.Name == WpgGroup)
                {
                    var properties = child.Element(WpgGroupProperties);
                    var frame = ChildFrame(properties?.Element(A.xfrm));
                    if (frame != null)
                        target.Add(GroupElement(child, properties, frame.Value, context));
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
        /// Draws one shape's preset geometry in <paramref name="frame"/>. Other geometry draws
        /// nothing: the drawing's extent is still reserved by the enclosing SVG.
        /// </summary>
        private static void AddShapeElement(XElement target, XElement shape, ShapeBox frame, ShapeContext context)
        {
            var properties = shape.Element(WPS.spPr);
            var shapeStyle = shape.Element(WPS.style);
            var geometry = (string?)properties?.Element(A.prstGeom)?.Attribute(NoNamespace.prst);
            XElement? element = geometry switch
            {
                "rect" => new XElement(Svg + "rect",
                    new XAttribute("x", Format(frame.X)), new XAttribute("y", Format(frame.Y)),
                    new XAttribute("width", Format(frame.Width)), new XAttribute("height", Format(frame.Height))),
                "roundRect" => RoundedRectangle(properties!.Element(A.prstGeom)!, frame),
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
            if (element == null)
                return;

            var isLine = element.Name.LocalName == "line";
            element.Add(new XAttribute("fill",
                isLine ? "none" : ShapeFill(properties, shapeStyle, context) ?? "none"));
            var (stroke, strokeWidth) = ShapeOutline(properties, shapeStyle, context);
            element.Add(new XAttribute("stroke", stroke ?? "none"));
            if (stroke != null)
                element.Add(new XAttribute("stroke-width", Format(strokeWidth)));
            var orientation = ShapeOrientation(properties?.Element(A.xfrm), frame);
            if (orientation != null)
                element.Add(new XAttribute("transform", orientation));
            target.Add(element);
        }

        private static XElement RoundedRectangle(XElement presetGeometry, ShapeBox frame)
        {
            // ECMA-376 presetShapeDefinitions: the corner radius is adj/100000 of the shorter
            // side, adj defaulting to 16667 and clamped to [0, 50000].
            var adjustment = presetGeometry.Element(A.avLst)?.Elements(A.gd)
                .Where(guide => (string?)guide.Attribute(NoNamespace.name) == "adj")
                .Select(guide => (string?)guide.Attribute(NoNamespace.fmla))
                .Select(formula => formula != null && formula.StartsWith("val ", StringComparison.Ordinal) &&
                    double.TryParse(formula.Substring(4), NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
                        ? (double?)value
                        : null)
                .FirstOrDefault() ?? 16667;
            var radius = Math.Min(frame.Width, frame.Height) * Math.Clamp(adjustment, 0, 50000) / 100000;
            return new XElement(Svg + "rect",
                new XAttribute("x", Format(frame.X)), new XAttribute("y", Format(frame.Y)),
                new XAttribute("width", Format(frame.Width)), new XAttribute("height", Format(frame.Height)),
                new XAttribute("rx", Format(radius)), new XAttribute("ry", Format(radius)));
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

        /// <summary>A shape's fill color, or null when it is unfilled.</summary>
        private static string? ShapeFill(XElement? properties, XElement? shapeStyle, ShapeContext context)
        {
            var fill = properties?.Elements().FirstOrDefault(IsFillChoice);
            if (fill != null)
                return FillColor(fill, context);

            // No explicit fill: the shape style's fill reference applies; index 0 means no fill.
            var reference = shapeStyle?.Element(A.fillRef);
            return ReadIndex(reference) > 0 ? ResolveDrawingColor(reference, context.Theme) : null;
        }

        /// <summary>A shape's outline color and width in EMUs; a null color means no outline.</summary>
        private static (string? Color, double Width) ShapeOutline(XElement? properties, XElement? shapeStyle,
            ShapeContext context)
        {
            var line = properties?.Element(A.ln);
            var reference = shapeStyle?.Element(A.lnRef);
            var referenceIndex = ReadIndex(reference);
            var fill = line?.Elements().FirstOrDefault(IsFillChoice);
            var color = fill != null
                ? FillColor(fill, context)
                : referenceIndex > 0 ? ResolveDrawingColor(reference, context.Theme) : null;
            // An outline with no width of its own takes the referenced theme line style's width;
            // otherwise it is DrawingML's thinnest line, drawn here as one CSS pixel (9525 EMUs).
            var width = (long?)line?.Attribute(NoNamespace.w)
                ?? (referenceIndex > 0 ? context.ThemeLineWidth(referenceIndex) : null)
                ?? 9525L;
            return (color, Math.Max(width, 9525L));
        }

        private static bool IsFillChoice(XElement element) =>
            element.Name == A.noFill || element.Name == A.solidFill || element.Name == A.gradFill ||
            element.Name == A.pattFill || element.Name == A.blipFill || element.Name == A.grpFill;

        /// <summary>
        /// The color a fill choice paints with. A gradient approximates to its first stop and a
        /// pattern to its foreground; picture and group fills are not drawn.
        /// </summary>
        private static string? FillColor(XElement fill, ShapeContext context)
        {
            if (fill.Name == A.solidFill)
                return ResolveDrawingColor(fill, context.Theme);
            if (fill.Name == A.gradFill)
                return ResolveDrawingColor(fill.Element(A.gsLst)?.Elements(A.gs).FirstOrDefault(), context.Theme);
            if (fill.Name == A.pattFill)
                return ResolveDrawingColor(fill.Element(A.fgClr), context.Theme);
            return null;
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
            var color = colorParent?.Elements()
                .FirstOrDefault(e => e.Name == A.srgbClr || e.Name == A.schemeClr || e.Name == A.sysClr);
            if (color == null)
                return null;

            string? rgb = null;
            if (color.Name == A.srgbClr)
                rgb = (string?)color.Attribute(NoNamespace.val);
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
