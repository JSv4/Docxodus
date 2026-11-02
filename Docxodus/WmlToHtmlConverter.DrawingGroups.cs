using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

public static partial class WmlToHtmlConverter
{
    private static XElement? ProcessDrawingTextGroup(WordprocessingDocument document, XElement drawing,
        Func<ImageInfo, XElement?>? imageHandler, WmlToHtmlConverterSettings? settings)
    {
        var container = drawing.Elements().FirstOrDefault(e => e.Name == WP.inline || e.Name == WP.anchor);
        var source = container?.Element(A.graphic)?.Element(A.graphicData)?.Elements().FirstOrDefault();
        if (source?.Name != WpgGroup
            || !source.Descendants().Any(e => e.Name == W.txbxContent || e.Name == WPS.wps + "linkedTxbx")) return null;

        var extent = container!.Element(WP.extent);
        var width = Math.Max(0L, (long?)extent?.Attribute(NoNamespace.cx) ?? 0);
        var height = Math.Max(0L, (long?)extent?.Attribute(NoNamespace.cy) ?? 0);
        var context = new ShapeContext(document);
        var wrapper = DrawingTextGroupMembers(document, source, new ShapeBox(0, 0, width, height),
            null, imageHandler, settings, context, new DrawingTextOrientation(0, 1, 1));
        var style = wrapper.Annotation<Dictionary<string, string>>()!;
        style["position"] = "relative";
        style["display"] = "inline-block";
        style["vertical-align"] = "baseline";
        var placed = PlaceDrawingShape(wrapper, container, settings);
        if (context.Undrawn > 0)
            ReportShapeLayoutLoss(wrapper, settings ?? new WmlToHtmlConverterSettings(),
                "Part of this drawing group is not rendered", source.Name.LocalName);
        if ((string?)wrapper.Attribute("data-docx-shape") == "partial")
            wrapper.Descendants().Attributes().Where(a => a.Name.LocalName is
                "data-docx-textbox-source-clip" or "data-docx-textbox-source-position").Remove();
        return placed;
    }

    private static XElement DrawingTextGroupMembers(WordprocessingDocument document, XElement source, ShapeBox frame,
        XElement? inheritedFill, Func<ImageInfo, XElement?>? imageHandler, WmlToHtmlConverterSettings? settings,
        ShapeContext context, DrawingTextOrientation parentOrientation)
    {
        var properties = source.Element(WpgGroupProperties);
        var transform = properties?.Element(A.xfrm);
        var orientation = parentOrientation.Then(transform);
        var offset = transform?.Element(A.chOff);
        var extent = transform?.Element(A.chExt);
        var childWidth = (long?)extent?.Attribute(NoNamespace.cx) ?? frame.Width;
        var childHeight = (long?)extent?.Attribute(NoNamespace.cy) ?? frame.Height;
        var scaleX = childWidth > 0 ? frame.Width / childWidth : 1;
        var scaleY = childHeight > 0 ? frame.Height / childHeight : 1;
        var map = new ShapeMap(scaleX, scaleY,
            -((long?)offset?.Attribute(NoNamespace.x) ?? 0) * scaleX,
            -((long?)offset?.Attribute(NoNamespace.y) ?? 0) * scaleY);
        var fill = properties?.Elements().FirstOrDefault(IsFillChoice);
        if (fill == null || fill.Name == A.grpFill) fill = inheritedFill;
        var group = new XElement(Xhtml.span, new XAttribute("data-docx-drawing-group", "true"));
        group.AddAnnotation(DrawingMemberStyle(frame, transform));
        foreach (var child in source.Elements())
        {
            XElement? member = null;
            var childProperties = child.Element(WPS.spPr) ?? child.Element(Pic.spPr) ?? child.Element(WpgGroupProperties);
            var childTransform = childProperties?.Element(A.xfrm);
            var childFrame = ChildFrame(childTransform);
            if (child.Name == WPS.wsp || child.Name == WpgNestedGroup || child.Name == WpgGroup || child.Name == Pic._pic)
            {
                // An unknown frame invalidates placement, but must not delete a text story.
                if (childFrame == null) context.Undrawn++;
                var bounds = childFrame == null ? new ShapeBox(0, 0, frame.Width, frame.Height) : map.Apply(childFrame.Value);
                if (child.Name == WpgNestedGroup || child.Name == WpgGroup)
                    member = DrawingTextGroupMembers(document, child, bounds, fill, imageHandler, settings, context, orientation);
                else if (child.Name == WPS.wsp)
                {
                    member = new XElement(Xhtml.span, new XAttribute("data-docx-drawing-member", "true"));
                    member.AddAnnotation(DrawingMemberStyle(bounds, childTransform));
                    var localFrame = new ShapeBox(0, 0, bounds.Width, bounds.Height);
                    var textBox = ProcessTextBox(document, child, settings, drawingFrame: localFrame);
                    var space = DrawingSpace();
                    AddShapeElement(space, child, localFrame, fill, context, textBox != null, applyOrientation: false);
                    var svg = DrawingSvg((long)Math.Round(bounds.Width), (long)Math.Round(bounds.Height), space);
                    svg.SetAttributeValue("style", "display:block;position:absolute;left:0;top:0;");
                    svg.SetAttributeValue("aria-hidden", "true");
                    member.Add(svg);
                    if (textBox != null)
                    {
                        var textStyle = textBox.Annotation<Dictionary<string, string>>()!;
                        textStyle["position"] = "relative";
                        var textOrientation = orientation.Then(childTransform);
                        var upright = IsTrueAttribute(child.Element(WPS.bodyPr), "upright");
                        // Reflections move the frame and its geometry, but never mirror its glyphs.
                        // Upright text additionally cancels the composed rotation about its own center.
                        if (textOrientation.FlipX != 1 || textOrientation.FlipY != 1
                            || (upright && textOrientation.Rotation != 0))
                        {
                            textStyle["transform"] = $"scale({textOrientation.FlipX},{textOrientation.FlipY})"
                                + (upright ? $" rotate({Format(-textOrientation.Rotation)}deg)" : "");
                            textStyle["transform-origin"] = "50% 50%";
                        }
                        member.Add(textBox);
                    }
                }
                else if (imageHandler != null)
                {
                    var id = (string?)child.Element(Pic.blipFill)?.Element(A.blip)?.Attribute(R.embed);
                    var owner = GetOwningPart(document, child);
                    var part = owner.Parts.FirstOrDefault(p => p.RelationshipId == id).OpenXmlPart as ImagePart;
                    if (part != null)
                        member = RenderImagePart(part, child, imageHandler, settings,
                            alt: (string?)child.Element(Pic.nvPicPr)?.Element(Pic.cNvPr)?.Attribute(NoNamespace.descr));
                    if (member == null) context.Undrawn++;
                    else
                    {
                        RecordImageRepresentation(settings, child);
                        member.Attribute("style")?.Remove();
                        member.AddAnnotation(DrawingMemberStyle(bounds, childTransform));
                        if (child.Element(Pic.blipFill)?.Element(A.srcRect)?.Attributes().Any(a => a.Value != "0") == true)
                            context.Undrawn++;
                    }
                }
            }
            else if (child.Name == WPG.wpg + "graphicFrame" || child.Name == WPC.wpc + "graphicFrame"
                || child.Name == W14.w14 + "contentPart" || child.Name == MC.AlternateContent)
                context.Undrawn++;
            if (member != null) group.Add(member);
        }
        return group;
    }

    private readonly record struct DrawingTextOrientation(double Rotation, int FlipX, int FlipY)
    {
        public DrawingTextOrientation Then(XElement? transform) => new(
            Rotation + FlipX * FlipY * ((long?)transform?.Attribute(NoNamespace.rot) ?? 0) / 60000.0,
            FlipX * (IsTrueAttribute(transform, "flipH") ? -1 : 1),
            FlipY * (IsTrueAttribute(transform, "flipV") ? -1 : 1));
    }

    // Map dimensions, not the whole subtree: grouping must not scale type or outline widths.
    private static Dictionary<string, string> DrawingMemberStyle(ShapeBox frame, XElement? transform)
    {
        var style = new Dictionary<string, string>
        {
            ["display"] = "inline-block", ["position"] = "absolute", ["line-height"] = "normal",
            ["left"] = VmlPointLength(frame.X / 12700), ["top"] = VmlPointLength(frame.Y / 12700),
            ["width"] = VmlPointLength(frame.Width / 12700), ["height"] = VmlPointLength(frame.Height / 12700),
        };
        var degrees = ((long?)transform?.Attribute(NoNamespace.rot) ?? 0) / 60000.0;
        var flipX = IsTrueAttribute(transform, "flipH") ? -1 : 1;
        var flipY = IsTrueAttribute(transform, "flipV") ? -1 : 1;
        if (degrees != 0 || flipX != 1 || flipY != 1)
        {
            style["transform"] = $"rotate({Format(degrees)}deg) scale({flipX},{flipY})";
            style["transform-origin"] = "50% 50%";
        }
        return style;
    }
}
