#nullable enable

using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Parity tests for M1.3 list resolution: the <see cref="IrListInfo"/> facts the
/// <see cref="IrReader"/> resolves through the numbering registry must agree, paragraph for
/// paragraph, with what <see cref="DocxSession.GetListMembership"/> reports for the SAME bytes.
/// Both code paths derive anchors from the same deterministic Unid grammar
/// (<c>kind:scope:unid</c>), so an IR paragraph's <see cref="IrAnchor.ToString"/> is a valid
/// <c>GetListMembership</c> anchor id with no translation.
///
/// <para>Normalization note: <see cref="IrListInfo.NumberFormat"/> is the raw <c>w:numFmt/@w:val</c>
/// string (e.g. "decimal", "lowerLetter", or "" when the level is missing), whereas
/// <see cref="ListMembership.Format"/> is the <see cref="NumberFormat"/> enum, which collapses every
/// unrecognized/absent value to <see cref="NumberFormat.Decimal"/>. The test maps the IR string
/// through the same parse the session uses (<see cref="MapFormat"/>) before comparing — documenting
/// that the two surfaces are equal modulo that lossy enum projection, not byte-for-byte.</para>
/// </summary>
public class IrListParityTests
{
    // A self-contained fixture exercising all three resolution paths:
    //   (a) direct w:numPr list items at two levels (numId 1 → abstractNum 0, decimal/lowerLetter),
    //   (b) a style-inherited list item (pStyle "MyList" → basedOn "MyListBase" carries numPr,
    //       numId 2 → abstractNum 1, upperRoman),
    //   (c) a list item whose w:num (numId 3 → abstractNum 0) applies a startOverride at ilvl 0.
    private const string Numbering =
        "<w:abstractNum w:abstractNumId=\"0\">" +
        "  <w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"decimal\"/><w:start w:val=\"1\"/></w:lvl>" +
        "  <w:lvl w:ilvl=\"1\"><w:numFmt w:val=\"lowerLetter\"/><w:start w:val=\"1\"/></w:lvl>" +
        "</w:abstractNum>" +
        "<w:abstractNum w:abstractNumId=\"1\">" +
        "  <w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"upperRoman\"/><w:start w:val=\"1\"/></w:lvl>" +
        "</w:abstractNum>" +
        "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>" +
        "<w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num>" +
        "<w:num w:numId=\"3\"><w:abstractNumId w:val=\"0\"/>" +
        "  <w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"5\"/></w:lvlOverride>" +
        "</w:num>";

    private const string Styles =
        "<w:style w:type=\"paragraph\" w:styleId=\"MyListBase\">" +
        "  <w:name w:val=\"My List Base\"/>" +
        "  <w:pPr><w:numPr><w:numId w:val=\"2\"/></w:numPr></w:pPr>" +
        "</w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"MyList\">" +
        "  <w:name w:val=\"My List\"/>" +
        "  <w:basedOn w:val=\"MyListBase\"/>" +
        "</w:style>";

    private const string Body =
        // (a) direct numPr, level 0, decimal.
        "<w:p><w:pPr><w:numPr><w:numId w:val=\"1\"/><w:ilvl w:val=\"0\"/></w:numPr></w:pPr>" +
        "  <w:r><w:t>direct level 0</w:t></w:r></w:p>" +
        // (a) direct numPr, level 1, lowerLetter.
        "<w:p><w:pPr><w:numPr><w:numId w:val=\"1\"/><w:ilvl w:val=\"1\"/></w:numPr></w:pPr>" +
        "  <w:r><w:t>direct level 1</w:t></w:r></w:p>" +
        // (b) style-inherited numPr via MyList → MyListBase, level 0, upperRoman.
        "<w:p><w:pPr><w:pStyle w:val=\"MyList\"/></w:pPr>" +
        "  <w:r><w:t>style inherited</w:t></w:r></w:p>" +
        // (c) direct numPr on numId 3, which carries a startOverride at ilvl 0.
        "<w:p><w:pPr><w:numPr><w:numId w:val=\"3\"/><w:ilvl w:val=\"0\"/></w:numPr></w:pPr>" +
        "  <w:r><w:t>start override</w:t></w:r></w:p>" +
        // A plain paragraph (no list) to confirm null on both surfaces.
        "<w:p><w:r><w:t>not a list</w:t></w:r></w:p>";

    [Fact]
    public void IrListInfo_MatchesGetListMembership_PerParagraph()
    {
        var doc = IrTestDocuments.FromParts(Body, Styles, Numbering);

        var ir = IrReader.Read(doc);
        using var session = new DocxSession(doc.DocumentByteArray);

        var paras = ir.Body.Blocks.OfType<IrParagraph>().ToList();
        Assert.Equal(5, paras.Count);

        // Sanity: four list items + one plain paragraph, as the fixture intends.
        Assert.Equal(4, paras.Count(p => p.List is not null));

        foreach (var p in paras)
        {
            var anchorId = p.Anchor.ToString();
            ListMembership? session_ = session.GetListMembership(anchorId);

            if (p.List is null)
            {
                Assert.Null(session_);
                continue;
            }

            Assert.NotNull(session_);
            var ir_ = p.List;
            var s = session_!;

            Assert.Equal(s.NumId, ir_.NumId);
            Assert.Equal(s.AbstractNumId, ir_.AbstractNumId);
            Assert.Equal(s.Level, ir_.Ilvl);
            Assert.Equal(s.StartOverride, ir_.StartOverride);
            Assert.Equal(s.FromStyle, ir_.FromStyle);
            // Normalize the IR raw numFmt string to the session's NumberFormat enum (see class doc).
            Assert.Equal(s.Format, MapFormat(ir_.NumberFormat));
        }
    }

    [Fact]
    public void IrListInfo_ResolvesExpectedFacts()
    {
        var doc = IrTestDocuments.FromParts(Body, Styles, Numbering);
        var ir = IrReader.Read(doc);
        var lists = ir.Body.Blocks.OfType<IrParagraph>()
            .Select(p => p.List)
            .Where(l => l is not null)
            .Select(l => l!)
            .ToList();

        // (a) direct level 0 → numId 1, abs 0, decimal, no override, not from style.
        Assert.Equal(new IrListInfo(1, 0, 0, "decimal", null, false), lists[0]);
        // (a) direct level 1 → lowerLetter.
        Assert.Equal(new IrListInfo(1, 0, 1, "lowerLetter", null, false), lists[1]);
        // (b) style-inherited → numId 2, abs 1, upperRoman, ilvl 0 (no ilvl in numPr), FromStyle.
        Assert.Equal(new IrListInfo(2, 1, 0, "upperRoman", null, true), lists[2]);
        // (c) startOverride 5 on numId 3 / abs 0 / level 0 (decimal).
        Assert.Equal(new IrListInfo(3, 0, 0, "decimal", 5, false), lists[3]);
    }

    /// <summary>
    /// Map IR's raw <c>w:numFmt</c> string to the public <see cref="NumberFormat"/> enum, mirroring
    /// <c>BlockMetadataOps.ParseNumberFormat</c> — any unrecognized/empty value (including the ""
    /// the reader emits for a missing level) collapses to <see cref="NumberFormat.Decimal"/>.
    /// </summary>
    private static NumberFormat MapFormat(string raw) => raw switch
    {
        "bullet" => NumberFormat.Bullet,
        "upperLetter" => NumberFormat.UpperLetter,
        "lowerLetter" => NumberFormat.LowerLetter,
        "upperRoman" => NumberFormat.UpperRoman,
        "lowerRoman" => NumberFormat.LowerRoman,
        _ => NumberFormat.Decimal,
    };
}
