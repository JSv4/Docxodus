// Deterministic edited variants of a source .docx.
//
// Every variant is produced by editing the source package's XML directly rather than by
// driving DocxSession, so the generator's own cost never lands in the measured numbers and
// the same input always yields byte-identical variants (the seeded RNG makes the "random"
// selections reproducible run to run).

using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;

namespace Docxodus.Stress;

internal sealed record Variant(string Name, string Description, byte[] Bytes);

internal static class Variants
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static List<Variant> Build(byte[] source) =>
    [
        new("identical", "no change at all (fast-path floor)", source),
        new("light", "8 scattered word-level edits", EditEveryNthTextNode(source, nth: 0, count: 8)),
        new("heavy", "an edit in every fifth paragraph", EditParagraphs(source, nth: 5)),
        new("churn", "an edit in every second text node", EditEveryNthTextNode(source, nth: 2, count: int.MaxValue)),
        new("reorder", "24 blocks relocated across the body", Reorder(source, blocks: 24)),
        new("structural", "20 paragraphs deleted, 20 inserted", Structural(source, count: 20)),
        new("footnotes", "an edit in every second footnote paragraph", EditFootnotes(source)),
        new("rewrite", "every paragraph's text replaced", Rewrite(source)),
    ];

    // ---- variant builders ----

    private static byte[] EditEveryNthTextNode(byte[] source, int nth, int count) =>
        Transform(source, body =>
        {
            var texts = body.Descendants(W + "t").Where(t => t.Value.Trim().Length > 3).ToList();
            var done = 0;
            if (nth == 0)
            {
                // Spread `count` edits evenly over the document instead of clustering them.
                var stride = Math.Max(1, texts.Count / Math.Max(1, count));
                for (var i = 0; i < texts.Count && done < count; i += stride)
                {
                    texts[i].Value = MutateWord(texts[i].Value, i);
                    done++;
                }
                return;
            }

            for (var i = 0; i < texts.Count && done < count; i++)
            {
                if (i % nth != 0) continue;
                texts[i].Value = MutateWord(texts[i].Value, i);
                done++;
            }
        });

    private static byte[] EditParagraphs(byte[] source, int nth) =>
        Transform(source, body =>
        {
            var paras = body.Elements(W + "p").ToList();
            for (var i = 0; i < paras.Count; i++)
            {
                if (i % nth != 0) continue;
                var t = paras[i].Descendants(W + "t").FirstOrDefault(x => x.Value.Trim().Length > 3);
                if (t != null) t.Value = MutateWord(t.Value, i);
            }
        });

    private static byte[] Rewrite(byte[] source) =>
        Transform(source, body =>
        {
            var i = 0;
            foreach (var t in body.Descendants(W + "t").ToList())
            {
                if (t.Value.Trim().Length > 0) t.Value = MutateWord(t.Value, i);
                i++;
            }
        });

    // Relocation stress: take blocks from the second quarter of the body and re-insert them
    // in the last quarter, unchanged, so the aligner has to recover them as moves.
    private static byte[] Reorder(byte[] source, int blocks) =>
        Transform(source, body =>
        {
            var all = body.Elements().Where(e => e.Name == W + "p" || e.Name == W + "tbl").ToList();
            if (all.Count < blocks * 4) blocks = Math.Max(1, all.Count / 4);
            var take = all.Skip(all.Count / 4).Take(blocks).ToList();
            var anchor = all[Math.Min(all.Count - 1, all.Count * 3 / 4)];
            foreach (var b in take) b.Remove();
            anchor.AddAfterSelf(take);
        });

    private static byte[] Structural(byte[] source, int count) =>
        Transform(source, body =>
        {
            var paras = body.Elements(W + "p")
                .Where(p => p.Descendants(W + "t").Any(t => t.Value.Trim().Length > 20))
                .ToList();
            if (paras.Count == 0) return;

            var stride = Math.Max(1, paras.Count / Math.Max(1, count * 2));
            var deleted = 0;
            var inserted = 0;
            for (var i = 0; i < paras.Count && (deleted < count || inserted < count); i += stride)
            {
                if (deleted < count && i % 2 == 0)
                {
                    paras[i].Remove();
                    deleted++;
                }
                else if (inserted < count)
                {
                    var clone = new XElement(paras[i]);
                    foreach (var t in clone.Descendants(W + "t").ToList())
                        t.Value = $"Negotiated addition {inserted}: {t.Value}";
                    // Drop bookmark/comment anchors so the insertion is genuinely new content.
                    foreach (var e in clone.Descendants().Where(IsAnchorish).ToList()) e.Remove();
                    paras[i].AddAfterSelf(clone);
                    inserted++;
                }
            }
        });

    private static bool IsAnchorish(XElement e) =>
        e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd" ||
        e.Name == W + "commentRangeStart" || e.Name == W + "commentRangeEnd" ||
        e.Name == W + "commentReference";

    private static byte[] EditFootnotes(byte[] source) =>
        TransformPart(source, doc => doc.MainDocumentPart?.FootnotesPart, root =>
        {
            var texts = root.Descendants(W + "t").Where(t => t.Value.Trim().Length > 3).ToList();
            for (var i = 0; i < texts.Count; i += 2)
                texts[i].Value = MutateWord(texts[i].Value, i);
        });

    // ---- plumbing ----

    private static byte[] Transform(byte[] source, Action<XElement> editBody) =>
        TransformPart(source, doc => doc.MainDocumentPart, root =>
        {
            var body = root.Element(W + "body");
            if (body != null) editBody(body);
        });

    private static byte[] TransformPart(
        byte[] source, Func<WordprocessingDocument, OpenXmlPart?> pick, Action<XElement> edit)
    {
        var wml = new WmlDocument("in.docx", source);
        using var streamDoc = new OpenXmlMemoryStreamDocument(wml);
        using (var doc = streamDoc.GetWordprocessingDocument())
        {
            var part = pick(doc);
            if (part != null)
            {
                var xdoc = part.GetXDocument();
                if (xdoc.Root != null)
                {
                    edit(xdoc.Root);
                    part.PutXDocument();
                }
            }
        }

        return streamDoc.GetModifiedWmlDocument().DocumentByteArray;
    }

    // A word-level substitution, not a whole-string replacement: this is what a counsel edit
    // looks like to the tokenizer, and it keeps the surrounding tokens Equal so the token
    // differ has real work to do rather than a trivially total replacement.
    private static string MutateWord(string value, int seed)
    {
        var words = value.Split(' ');
        var target = -1;
        for (var i = 0; i < words.Length; i++)
        {
            if (words[i].Trim().Length <= 2) continue;
            target = i;
            if (i >= seed % Math.Max(1, words.Length)) break;
        }

        if (target < 0) return value + " (amended)";
        words[target] = Replacements[seed % Replacements.Length];
        return string.Join(' ', words);
    }

    private static readonly string[] Replacements =
    [
        "amended", "restated", "supplemental", "conditional", "irrevocable",
        "notwithstanding", "pari-passu", "as-converted", "post-money", "pre-money",
        "liquidation", "redemption", "conversion", "protective", "cumulative",
    ];
}
