using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

public static partial class WmlToHtmlConverter
{
    // Resolve these asymmetric defaults before formatting assembly replaces direct run properties:
    // hps wins over annotation run sizes; an explicit base run size wins over hpsBaseText.
    private static void ResolveRubySizes(WordprocessingDocument document)
    {
        foreach (var part in document.ContentParts())
        foreach (var ruby in part.GetXDocument().Descendants(W.ruby))
        foreach (var content in ruby.Elements().Where(e => e.Name == W.rt || e.Name == W.rubyBase))
        {
            var annotation = content.Name == W.rt;
            var size = ruby.Element(W.rubyPr)?.Element(annotation ? W.hps : W.hpsBaseText)?.Attribute(W.val);
            if (size == null) continue;
            foreach (var run in content.Descendants(W.r).Where(run => run.Ancestors(W.ruby).First() == ruby))
            {
                var properties = run.Element(W.rPr);
                if (properties == null) run.AddFirst(properties = new XElement(W.rPr));
                foreach (var name in new[] { W.sz, W.szCs })
                {
                    if (!annotation && properties.Element(name) != null) continue;
                    properties.Elements(name).Remove();
                    properties.Add(new XElement(name, new XAttribute(W.val, size.Value)));
                }
            }
        }
    }

    private static XElement ProcessRuby(WordprocessingDocument document, WmlToHtmlConverterSettings settings, XElement source)
    {
        var properties = source.Element(W.rubyPr);
        var baseText = source.Element(W.rubyBase);
        var annotation = source.Element(W.rt);
        var alignment = (string?)properties?.Element(W.rubyAlign)?.Attribute(W.val) switch
        {
            "center" => "center",
            "left" => "start",
            "distributeLetter" => "space-between",
            "distributeSpace" => "space-around",
            _ => null,
        };
        var size = (string?)properties?.Element(W.hps)?.Attribute(W.val);
        var raise = (string?)properties?.Element(W.hpsRaise)?.Attribute(W.val);
        bool SimpleRuns(XElement? content) => content != null && content.HasElements
            && content.Elements().All(run => run.Name == W.r && run.Elements().All(e => e.Name == W.rPr || e.Name == W.t))
            && !content.Descendants().Any(e => e.Name == W.position || e.Name == W.vertAlign);
        object? Convert(XElement? content) => content?.Elements()
            .Select(e => ConvertToHtmlTransform(document, settings, e, false, 0m));
        if (alignment == null || !SimpleRuns(baseText) || !SimpleRuns(annotation)
            || !decimal.TryParse(size, NumberStyles.Number, CultureInfo.InvariantCulture, out var halfPoints) || halfPoints <= 0
            || !decimal.TryParse(raise, NumberStyles.Number, CultureInfo.InvariantCulture, out var halfRaise) || halfRaise < 0
            || source.Ancestors(W.r).Elements(W.rPr).Elements().Any(e => e.Name == W.position || e.Name == W.vertAlign)
            || source.Ancestors(W.p).Any(p => IsBidi(p) || p.Descendants(W.tab).Any()
                || (string?)p.Element(W.pPr)?.Element(W.spacing)?.Attribute(W.lineRule) == "exact"))
        {
            var fallback = CreateUnsupportedContentPlaceholder(settings, source, UnsupportedContentType.RubyAnnotation, "")
                ?? new XElement(Xhtml.span);
            fallback.Add(Convert(baseText), " (", Convert(annotation), ")");
            // The enclosing source run already owns exact-line baseline correction. Its nested
            // ruby runs become ordinary inline fallback text, not additional correction owners.
            foreach (var run in fallback.Descendants().Where(e => e.Attribute("data-docx-exact-run") != null))
            {
                run.Attribute("data-docx-exact-run")!.Remove();
                if (run.Annotation<Dictionary<string, string>>() is { } style)
                {
                    style.Remove("vertical-align");
                    style["line-height"] = "normal";
                }
            }
            return fallback;
        }

        var rt = new XElement(Xhtml.xhtml + "rt", Convert(annotation));
        var language = (string?)properties?.Element(W.lid)?.Attribute(W.val);
        if (!string.IsNullOrEmpty(language)) rt.SetAttributeValue("lang", language);
        rt.AddAnnotation(new Dictionary<string, string> { ["font-size"] = (halfPoints / 2).ToString(CultureInfo.InvariantCulture) + "pt" });
        var ruby = new XElement(Xhtml.xhtml + "ruby",
            new XElement(Xhtml.span, new XAttribute("data-docx-ruby-base", "true"), Convert(baseText)),
            new XElement(Xhtml.xhtml + "rp", "("), rt, new XElement(Xhtml.xhtml + "rp", ")"));
        ruby.AddAnnotation(new Dictionary<string, string> { ["ruby-align"] = alignment });
        var wrapper = new XElement(Xhtml.span, new XAttribute("data-docx-ruby", "true"),
            new XAttribute("data-docx-ruby-raise", (halfRaise / 2).ToString(CultureInfo.InvariantCulture)), ruby);
        wrapper.AddAnnotation(new Dictionary<string, string> { ["display"] = "inline-block" });
        return wrapper;
    }
}
