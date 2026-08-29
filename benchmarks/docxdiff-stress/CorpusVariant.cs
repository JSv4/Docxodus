// A deterministic edit that works on an arbitrary .docx, however odd its contents.
//
// The generated variants in Variants.cs assume a document shaped like the reference legal form. A
// corpus of 678 fixtures contains documents with no body text, no main part, deliberately malformed
// XML, and packages that are not really Wordprocessing at all — so this one is written to degrade
// rather than throw: anything it cannot edit comes back unchanged, which still exercises the
// identical-bytes paths for that document.

using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Docxodus.Stress;

internal static class CorpusVariant
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// Edit every fourth non-trivial <c>w:t</c> in the main document part, by word substitution so
    /// the surrounding tokens stay Equal and the token differ has real work to do. Operates on the
    /// zip entry directly rather than through the SDK, so a package the SDK rejects still produces a
    /// usable right-hand side instead of aborting the document's whole row.
    /// </summary>
    public static byte[] Edit(byte[] source)
    {
        using var ms = new MemoryStream();
        ms.Write(source, 0, source.Length);

        try
        {
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
            {
                var entry = zip.GetEntry("word/document.xml");
                if (entry is null) return source;

                string xml;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    xml = reader.ReadToEnd();

                XDocument doc;
                try { doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace); }
                catch (System.Xml.XmlException) { return source; }

                var texts = doc.Descendants(W + "t").Where(t => t.Value.Trim().Length > 3).ToList();
                if (texts.Count == 0) return source;

                var edits = 0;
                for (var i = 0; i < texts.Count; i += 4)
                {
                    texts[i].Value = MutateWord(texts[i].Value, i);
                    edits++;
                }

                if (edits == 0) return source;

                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.BaseStream.SetLength(0);
                writer.Write(doc.ToString(SaveOptions.DisableFormatting));
            }

            return ms.ToArray();
        }
        catch (InvalidDataException)
        {
            // Not a readable zip. Comparing it against itself is still a meaningful row.
            return source;
        }
    }

    private static string MutateWord(string value, int seed)
    {
        var words = value.Split(' ');
        for (var i = 0; i < words.Length; i++)
        {
            if (words[i].Trim().Length <= 2) continue;
            words[i] = Replacements[seed % Replacements.Length];
            return string.Join(' ', words);
        }

        return value + " " + Replacements[seed % Replacements.Length];
    }

    private static readonly string[] Replacements =
    [
        "amended", "restated", "supplemental", "conditional", "irrevocable",
        "notwithstanding", "reconciled", "superseded", "novated", "assigned",
    ];
}
