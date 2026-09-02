#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #668: a single tracked text replacement rewrote the XML payload of 23 of the NVCA
/// charter's 44 parts — every header and footer, both note parts, styles and settings. Their XML
/// was unchanged; what changed was three bytes at the front. <see cref="DocxSession.Save"/>
/// serializes every projected part, and the writer it went through imposed a UTF-8 byte-order mark
/// on parts that had none.
///
/// <para>The invariant these tests pin: a part outside the mutation's contribution set keeps its
/// decompressed payload byte for byte. That is what makes a package diff show the edit rather than
/// bury it, keeps content-addressed stores and part digests stable, and makes it possible to prove
/// what an automated edit did <em>not</em> touch.</para>
///
/// <para>The ZIP container itself is deliberately not part of the invariant: compression level and
/// entry ordering stay implementation-defined, so every comparison here is over decompressed part
/// contents.</para>
/// </summary>
public class DocxSessionPartPayloadIdentityTests
{
    private static readonly string NvcaPath =
        Path.Combine("../../../../TestFiles/", "NVCA-Model-COI.docx");

    private static Dictionary<string, byte[]> PartPayloads(byte[] docx)
    {
        var payloads = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using var ms = new MemoryStream(docx);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        foreach (var entry in zip.Entries)
        {
            using var stream = entry.Open();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            payloads[entry.FullName] = buffer.ToArray();
        }

        return payloads;
    }

    private static List<string> ChangedParts(byte[] before, byte[] after)
    {
        var b = PartPayloads(before);
        var a = PartPayloads(after);
        return b.Keys
            .Where(name => a.TryGetValue(name, out var payload)
                && !payload.AsSpan().SequenceEqual(b[name]))
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToList();
    }

    private static string Digest(byte[] payload) => Convert.ToHexString(SHA256.HashData(payload));

    /// <summary>Replace the charter's first `[specify percentage]` placeholder, tracked.</summary>
    private static byte[] OneTrackedReplacement(byte[] source)
    {
        using var session = new DocxSession(source, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Payload Identity",
        });
        var hit = session.Grep(Regex.Escape("[specify percentage]")).First();
        Assert.True(session.ReplaceMatch(hit, "a majority").Success);
        return session.Save();
    }

    [Fact]
    public void ATrackedReplacementRewritesOnlyTheStoryAndTheTrackedChangesSetting()
    {
        var before = File.ReadAllBytes(NvcaPath);
        var after = OneTrackedReplacement(before);

        // word/document.xml carries the edit. word/settings.xml gains <w:trackRevisions/>, which
        // the tracked-change mode genuinely requires. Nothing else may move.
        Assert.Equal(
            new[] { "word/document.xml", "word/settings.xml" },
            ChangedParts(before, after));
    }

    [Fact]
    public void EveryHeaderFooterNoteAndDefinitionPartKeepsItsExactPayload()
    {
        var before = File.ReadAllBytes(NvcaPath);
        var after = OneTrackedReplacement(before);

        var b = PartPayloads(before);
        var a = PartPayloads(after);

        // Named explicitly rather than derived, so this fails loudly if the fixture ever loses a
        // running story instead of quietly asserting over a smaller set.
        var untouched = b.Keys
            .Where(name => name.StartsWith("word/header", StringComparison.Ordinal)
                || name.StartsWith("word/footer", StringComparison.Ordinal)
                || name is "word/footnotes.xml" or "word/endnotes.xml"
                    or "word/styles.xml" or "word/numbering.xml" or "word/fontTable.xml"
                    or "word/webSettings.xml" or "word/theme/theme1.xml")
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToList();

        // The charter carries 8 headers and 10 footers; if that ever stops being true the
        // assertion below would be checking far less than it claims to.
        Assert.Equal(8, untouched.Count(n => n.StartsWith("word/header", StringComparison.Ordinal)));
        Assert.Equal(10, untouched.Count(n => n.StartsWith("word/footer", StringComparison.Ordinal)));

        foreach (var name in untouched)
        {
            Assert.True(a.ContainsKey(name), $"{name} disappeared from the saved package");
            Assert.Equal(Digest(b[name]), Digest(a[name]));
        }
    }

    [Fact]
    public void TheSavedPackageKeepsItsPartInventoryAndValidationDelta()
    {
        var before = File.ReadAllBytes(NvcaPath);
        var after = OneTrackedReplacement(before);

        Assert.Equal(
            PartPayloads(before).Keys.OrderBy(n => n, StringComparer.Ordinal).ToList(),
            PartPayloads(after).Keys.OrderBy(n => n, StringComparer.Ordinal).ToList());

        static int SchemaFindings(byte[] docx)
        {
            using var ms = new MemoryStream(docx);
            using var word = WordprocessingDocument.Open(ms, false);
            return new OpenXmlValidator().Validate(word).Count();
        }

        // Skipping a write must not be able to leave a package that validates worse than its input.
        Assert.True(SchemaFindings(after) <= SchemaFindings(before));
    }

    [Fact]
    public void AnEditToAHeaderRewritesThatHeaderAndNothingElse()
    {
        // The counterpart risk of writing less: an edit that lands OUTSIDE word/document.xml must
        // still reach the package, or "nothing changed" would be satisfied by losing the edit.
        // It is also the sharper form of the invariant — the body is untouched here, so it must
        // keep its bytes exactly the way the headers do when the body is the one being edited.
        var before = File.ReadAllBytes(NvcaPath);

        byte[] after;
        using (var session = new DocxSession(before))
        {
            var hit = session.Grep(
                    Regex.Escape("This sample document is the work product"),
                    scope: ProjectionScopes.Headers)
                .FirstOrDefault();
            Assert.NotNull(hit);
            Assert.True(session.ReplaceMatch(hit!, "This redlined document").Success);
            after = session.Save();
        }

        var changed = ChangedParts(before, after);
        Assert.True(
            changed.Any(name => name.StartsWith("word/header", StringComparison.Ordinal)),
            "the edited header was not written; changed: " + string.Join(", ", changed));
        Assert.DoesNotContain("word/document.xml", changed);
    }

    [Fact]
    public void MarkupDetailOutsideTheXmlInfosetIsNormalizedRatherThanPreserved()
    {
        // The boundary of the invariant, stated rather than assumed. A part is preserved because
        // the round-trip through LINQ-to-XML reproduces it, not because the write is skipped — so
        // anything the XML infoset does not carry is normalized. Empty-element spelling IS carried
        // (`<w:x></w:x>` survives as long form), which covers what OOXML producers actually emit;
        // whitespace INSIDE a tag is not.
        var before = File.ReadAllBytes(NvcaPath);

        var doctored = WithRewrittenHeader(before, xml =>
            new Regex(@"<(w:[A-Za-z]+)([^>/]*?)\s*/>").Replace(
                xml, m => $"<{m.Groups[1].Value}{m.Groups[2].Value}></{m.Groups[1].Value}>", 1));
        byte[] afterLongForm;
        using (var session = new DocxSession(doctored)) afterLongForm = session.Save();

        // Long-form empty elements survive: this is the case a naive reading of "reserialized"
        // would expect to break, and it does not.
        Assert.DoesNotContain("word/header1.xml", ChangedParts(doctored, afterLongForm));

        var spaced = WithRewrittenHeader(before, xml =>
            new Regex(@"<(w:[A-Za-z]+)([^>/]*?)\s*/>").Replace(
                xml, m => $"<{m.Groups[1].Value}{m.Groups[2].Value}  />", 1));
        byte[] afterSpaced;
        using (var session = new DocxSession(spaced)) afterSpaced = session.Save();

        // Insignificant whitespace inside a tag is not part of the infoset and does not survive.
        // No OOXML producer emits it; recording the limit is the point.
        Assert.Contains("word/header1.xml", ChangedParts(spaced, afterSpaced));
    }

    private static byte[] WithRewrittenHeader(byte[] docx, Func<string, string> rewrite)
    {
        using var output = new MemoryStream();
        using (var source = new MemoryStream(docx))
        using (var sourceZip = new ZipArchive(source, ZipArchiveMode.Read))
        using (var targetZip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in sourceZip.Entries)
            {
                using var entryStream = entry.Open();
                using var buffer = new MemoryStream();
                entryStream.CopyTo(buffer);
                var payload = buffer.ToArray();
                if (entry.FullName == "word/header1.xml")
                    payload = Encoding.UTF8.GetBytes(rewrite(Encoding.UTF8.GetString(payload)));
                using var target = targetZip.CreateEntry(entry.FullName).Open();
                target.Write(payload, 0, payload.Length);
            }
        }

        return output.ToArray();
    }

    [Fact]
    public void ANoEditOpenAndSaveLeavesEveryPartPayloadAlone()
    {
        var before = File.ReadAllBytes(NvcaPath);
        byte[] after;
        using (var session = new DocxSession(before)) after = session.Save();

        Assert.Empty(ChangedParts(before, after));
    }
}
