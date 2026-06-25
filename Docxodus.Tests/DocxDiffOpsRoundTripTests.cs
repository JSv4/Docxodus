#nullable enable

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Headless CI guard for the byte-in / byte-out accept/reject surface added to
/// <see cref="DocxDiffOps"/> — the primitive both the WASM/npm and stdio/docx-scalpel clients route through to
/// verify a redline's round-trip contract. This is the .NET-level oracle behind the client round-trip tests
/// (<c>npm/tests/docx-diff.spec.ts</c>, <c>python/tests/test_docx_diff.py</c>): if the Ops surface itself were
/// wrong, every client would be too. Asserts the actual contract — accept(compare(left,right)) ≡ right and
/// reject ≡ left at the body-text level — not the shape of the result.
/// </summary>
public class DocxDiffOpsRoundTripTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");
    private static byte[] Wc(string name) => File.ReadAllBytes(Path.Combine(TestFilesDir.FullName, "WC", name));

    private static string BodyText(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        return body is null ? "" : string.Concat(body.Descendants<Text>().Select(t => t.Text));
    }

    [Theory]
    [InlineData("WC001-Digits.docx", "WC001-Digits-Mod.docx")]
    [InlineData("WC004-Large.docx", "WC004-Large-Mod.docx")]
    public void AcceptRejectRoundTrip_MaterializesRightAndLeft(string leftName, string rightName)
    {
        var left = Wc(leftName);
        var right = Wc(rightName);

        var redline = DocxDiffOps.Compare(left, right, null);
        var accepted = DocxDiffOps.AcceptRevisions(redline);
        var rejected = DocxDiffOps.RejectRevisions(redline);

        Assert.NotEqual(BodyText(left), BodyText(right));     // the pair genuinely differs
        Assert.Equal(BodyText(right), BodyText(accepted));    // accept ≡ right
        Assert.Equal(BodyText(left), BodyText(rejected));     // reject ≡ left
    }

    [Fact]
    public void AcceptOrReject_EmptyInput_Throws()
    {
        Assert.Throws<ArgumentException>(() => DocxDiffOps.AcceptRevisions(Array.Empty<byte>()));
        Assert.Throws<ArgumentException>(() => DocxDiffOps.RejectRevisions(Array.Empty<byte>()));
    }
}
