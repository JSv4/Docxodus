#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// Resolves <c>mc:AlternateContent</c> the way Word does on open, for the shapes where keeping the
/// wrapper visibly diverges from Word's own compare output (which re-serializes the RESOLVED
/// content). Two oracle-proven rules, both conservative:
/// <list type="bullet">
/// <item>A <c>mc:Choice</c> requiring only VML namespaces (Word's strict-save watermark shape,
/// <c>Requires="v"</c>) is unwrapped to its bare <c>w:pict</c> payload — LibreOffice does not
/// render the wrapped form, Word Compare emits it bare.</item>
/// <item>When NO choice is understood (e.g. the obsolete Office 2008/6/28 draft
/// wordprocessingShape namespace), the <c>mc:Fallback</c> content is inlined — Word renders the
/// fallback VML; LibreOffice renders nothing for the original.</item>
/// </list>
/// Modern DrawingML choices (canonical 2010 wps/wpg/wpc) keep their wrapper — every reader
/// understands them and Word Compare preserves them. Untouched documents are returned as the same
/// instance (no copy).
/// </summary>
internal static class MarkupCompatibilityNormalizer
{
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private static readonly HashSet<string> VmlNamespaces = new(StringComparer.Ordinal)
    {
        "urn:schemas-microsoft-com:vml",
        "urn:schemas-microsoft-com:office:office",
        "urn:schemas-microsoft-com:office:word",
    };

    /// <summary>Namespaces a modern Word build understands in a <c>Requires</c> list. Anything
    /// outside this set (notably pre-release draft namespaces) makes the choice unreadable.</summary>
    private static readonly HashSet<string> UnderstoodNamespaces = new(StringComparer.Ordinal)
    {
        "urn:schemas-microsoft-com:vml",
        "urn:schemas-microsoft-com:office:office",
        "urn:schemas-microsoft-com:office:word",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        "http://schemas.microsoft.com/office/word/2010/wordml",
        "http://schemas.microsoft.com/office/word/2012/wordml",
        "http://schemas.microsoft.com/office/word/2018/wordml",
        "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        "http://schemas.microsoft.com/office/drawing/2010/main",
        "http://schemas.microsoft.com/office/drawing/2014/main",
    };

    internal static WmlDocument Normalize(WmlDocument doc)
    {
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        var anyChanged = false;
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries.ToList())
            {
                if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    continue;

                string text;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    text = reader.ReadToEnd();
                if (!text.Contains("AlternateContent", StringComparison.Ordinal))
                    continue;

                var rewritten = ResolveAlternateContent(text);
                if (rewritten is null)
                    continue;

                anyChanged = true;
                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.BaseStream.SetLength(0);
                writer.Write(rewritten);
            }
        }
        return anyChanged ? new WmlDocument(doc.FileName, ms.ToArray()) : doc;
    }

    /// <summary>Returns the rewritten part XML, or null when nothing needed resolving.</summary>
    private static string? ResolveAlternateContent(string xml)
    {
        XDocument doc;
        try
        {
            doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        }
        catch (System.Xml.XmlException)
        {
            return null;
        }

        var changed = false;
        foreach (var ac in doc.Descendants(Mc + "AlternateContent").ToList())
        {
            var selected = ac.Elements(Mc + "Choice")
                .FirstOrDefault(c => RequiredNamespaces(c).All(UnderstoodNamespaces.Contains));
            if (selected is not null)
            {
                // Only VML-only choices are unwrapped; modern DrawingML wrappers stay.
                var required = RequiredNamespaces(selected).ToList();
                if (required.Count == 0 || !required.All(VmlNamespaces.Contains))
                    continue;
                ac.ReplaceWith(selected.Nodes());
                changed = true;
            }
            else
            {
                var fallback = ac.Element(Mc + "Fallback");
                if (fallback is null)
                    continue;
                ac.ReplaceWith(fallback.Nodes());
                changed = true;
            }
        }
        if (!changed)
            return null;

        using var sw = new Utf8StringWriter();
        doc.Save(sw, SaveOptions.DisableFormatting);
        return sw.ToString();
    }

    /// <summary>The namespaces a choice's <c>Requires</c> prefix list resolves to in scope.
    /// An unresolvable prefix yields an empty marker that never matches the understood set.</summary>
    private static IEnumerable<string> RequiredNamespaces(XElement choice)
    {
        var requires = (string?)choice.Attribute("Requires");
        if (string.IsNullOrWhiteSpace(requires))
            yield break;
        foreach (var prefix in requires.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            yield return choice.GetNamespaceOfPrefix(prefix)?.NamespaceName ?? string.Empty;
    }

    private sealed class Utf8StringWriter : StringWriter
    {
        public override Encoding Encoding => Encoding.UTF8;
    }
}
