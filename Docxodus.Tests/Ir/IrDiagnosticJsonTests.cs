#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

public class IrDiagnosticJsonTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    private static WmlDocument Fixture(string name) =>
        new(Path.Combine(TestFilesDir.FullName, name));

    [Fact]
    public void DiagnosticJson_TwoReads_ByteIdentical()
    {
        var doc = Fixture("HC031-Complicated-Document.docx");

        var json1 = IrDiagnosticJson.Write(IrReader.Read(doc));
        var json2 = IrDiagnosticJson.Write(IrReader.Read(doc));

        Assert.Equal(json1, json2);
    }

    [Fact]
    public void DiagnosticJson_SimpleDocument_ContainsExpectedStructure()
    {
        var doc = IrTestDocuments.Create("Hello world", "Second line");

        var json = IrDiagnosticJson.Write(IrReader.Read(doc));

        // Valid JSON.
        using var parsed = JsonDocument.Parse(json);
        var root = parsed.RootElement;

        // Document level is now {"scopes":[…]} with the body entry first.
        var bodyScope = root.GetProperty("scopes")[0];
        Assert.Equal("body", bodyScope.GetProperty("scope").GetString());
        var blocks = bodyScope.GetProperty("blocks");
        Assert.Equal(JsonValueKind.Array, blocks.ValueKind);

        // Two paragraphs, each carrying an anchor, content hash, and the right text.
        Assert.Equal(2, blocks.GetArrayLength());
        var first = blocks[0];
        Assert.Equal("paragraph", first.GetProperty("type").GetString());
        Assert.Matches("^p:body:[0-9a-f]{32}$", first.GetProperty("anchor").GetString());
        Assert.Matches("^[0-9a-f]{64}$", first.GetProperty("contentHash").GetString());
        Assert.Matches("^[0-9a-f]{64}$", first.GetProperty("formatFingerprint").GetString());

        var firstInlines = first.GetProperty("inlines");
        Assert.Equal("text", firstInlines[0].GetProperty("kind").GetString());
        Assert.Equal("Hello world", firstInlines[0].GetProperty("text").GetString());

        var second = blocks[1];
        Assert.Equal("Second line", second.GetProperty("inlines")[0].GetProperty("text").GetString());
    }

    [Fact]
    public void DiagnosticJson_IsValidJson_ForComplexFixture()
    {
        var doc = Fixture("HC001-5DayTourPlanTemplate.docx");

        var json = IrDiagnosticJson.Write(IrReader.Read(doc));

        // Parsing throws if the output is not well-formed JSON.
        using var parsed = JsonDocument.Parse(json);
        var bodyScope = parsed.RootElement.GetProperty("scopes")[0];
        Assert.Equal("body", bodyScope.GetProperty("scope").GetString());
        Assert.True(bodyScope.GetProperty("blocks").GetArrayLength() > 0);
    }

    // --- writer/reader lockstep completeness guard ------------------------
    //
    // The diagnostic writer keeps an "unsupported" fallback for kinds the reader can't yet emit.
    // These guards enforce that EVERY concrete IrInline / IrBlock subtype has a real writer branch:
    // a new model kind cannot ship without the writer being extended in the same change. (This is
    // the lockstep enforcement the M1.1 final review demanded.) We discover the concrete types via
    // reflection, construct a minimal instance of each, serialize, and assert no "unsupported" kind
    // string appears anywhere in the output.

    private static readonly IrHash ZeroHash = default;
    private static readonly IrRunFormat EmptyRunFormat = new() { UnmodeledDigest = ZeroHash };
    private static readonly IrParaFormat EmptyParaFormat = new() { UnmodeledDigest = ZeroHash };

    private static IrAnchor Anchor(IrAnchorKind kind) =>
        new(kind, "body", new string('0', 32));

    /// <summary>Wrap a single inline in a minimal one-paragraph document and render it.</summary>
    private static string RenderInline(IrInline inline)
    {
        var para = new IrParagraph
        {
            Anchor = Anchor(IrAnchorKind.P),
            ContentHash = ZeroHash,
            FormatFingerprint = ZeroHash,
            Format = EmptyParaFormat,
            Inlines = IrNodeList.From(new[] { inline }),
        };
        return IrDiagnosticJson.Write(WrapBlock(para));
    }

    /// <summary>Wrap a single block in a minimal document and render it.</summary>
    private static string RenderBlock(IrBlock block) => IrDiagnosticJson.Write(WrapBlock(block));

    private static IrDocument WrapBlock(IrBlock block) => new()
    {
        Body = new IrScope("body", IrNodeList.From(new[] { block })),
        Footnotes = IrNoteStore.Empty,
        Endnotes = IrNoteStore.Empty,
        Comments = IrCommentStore.Empty,
        Styles = IrStyleRegistry.Empty,
        Numbering = IrNumberingRegistry.Empty,
        ThemeFonts = IrThemeFonts.Empty,
        AnchorIndex = new Dictionary<string, IrBlock>(),
        Sources = new Dictionary<Uri, XDocument>(),
    };

    private static IEnumerable<Type> ConcreteSubtypes<TBase>() =>
        typeof(IrInline).Assembly.GetTypes()
            .Where(t => !t.IsAbstract && typeof(TBase).IsAssignableFrom(t));

    /// <summary>A minimal instance of each concrete <see cref="IrInline"/> kind for the guard.</summary>
    private static IrInline MinimalInline(Type t) => t switch
    {
        _ when t == typeof(IrTextRun) => new IrTextRun("x", EmptyRunFormat),
        _ when t == typeof(IrTab) => new IrTab(EmptyRunFormat),
        _ when t == typeof(IrBreak) => new IrBreak(IrBreakKind.Line),
        _ when t == typeof(IrHyperlink) =>
            new IrHyperlink("https://example.com", null, IrNodeList.Empty<IrInline>()),
        _ when t == typeof(IrFieldRun) => new IrFieldRun("PAGE", IrNodeList.Empty<IrInline>()),
        _ when t == typeof(IrNoteRef) => new IrNoteRef(IrNoteKind.Footnote, "1"),
        _ when t == typeof(IrInlineImage) =>
            new IrInlineImage(new Uri("/word/media/i.png", UriKind.Relative), ZeroHash, 1, 2, null),
        _ when t == typeof(IrOpaqueInline) =>
            new IrOpaqueInline(XName.Get("thing", "urn:x"), ZeroHash),
        _ when t == typeof(IrTextbox) =>
            new IrTextbox(IrNodeList.From(new IrBlock[]
            {
                new IrParagraph
                {
                    Anchor = Anchor(IrAnchorKind.P),
                    ContentHash = ZeroHash,
                    FormatFingerprint = ZeroHash,
                    Format = EmptyParaFormat,
                    Inlines = IrNodeList.From(new IrInline[] { new IrTextRun("inside", EmptyRunFormat) }),
                },
            })),
        _ => throw new Xunit.Sdk.XunitException(
            $"No minimal-instance factory for inline kind '{t.Name}'. Add one to MinimalInline " +
            "and a writer branch to IrDiagnosticJson, then this guard will pass."),
    };

    /// <summary>A minimal instance of each concrete <see cref="IrBlock"/> kind for the guard.</summary>
    private static IrBlock MinimalBlock(Type t) => t switch
    {
        _ when t == typeof(IrParagraph) => new IrParagraph
        {
            Anchor = Anchor(IrAnchorKind.P),
            ContentHash = ZeroHash,
            FormatFingerprint = ZeroHash,
            Format = EmptyParaFormat,
            Inlines = IrNodeList.Empty<IrInline>(),
        },
        _ when t == typeof(IrTable) => new IrTable
        {
            Anchor = Anchor(IrAnchorKind.Tbl),
            ContentHash = ZeroHash,
            FormatFingerprint = ZeroHash,
            Rows = IrNodeList.Empty<IrRow>(),
            UnmodeledTablePropsDigest = ZeroHash,
        },
        _ when t == typeof(IrSectionBreak) => new IrSectionBreak
        {
            Anchor = Anchor(IrAnchorKind.Sec),
            ContentHash = ZeroHash,
            FormatFingerprint = ZeroHash,
            Format = new IrSectionFormat { UnmodeledDigest = ZeroHash },
        },
        _ when t == typeof(IrOpaqueBlock) => new IrOpaqueBlock
        {
            Anchor = Anchor(IrAnchorKind.Unk),
            ContentHash = ZeroHash,
            FormatFingerprint = ZeroHash,
            ElementName = XName.Get("thing", "urn:x"),
        },
        _ => throw new Xunit.Sdk.XunitException(
            $"No minimal-instance factory for block kind '{t.Name}'. Add one to MinimalBlock " +
            "and a writer branch to IrDiagnosticJson, then this guard will pass."),
    };

    private static void AssertNoUnsupportedKind(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        foreach (var kind in AllKindAndTypeStrings(parsed.RootElement))
            Assert.NotEqual("unsupported", kind);
    }

    // Collect every "kind" (inline) and "type" (block) string value anywhere in the tree.
    private static IEnumerable<string> AllKindAndTypeStrings(JsonElement el)
    {
        switch (el.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var prop in el.EnumerateObject())
                {
                    if ((prop.Name == "kind" || prop.Name == "type")
                        && prop.Value.ValueKind == JsonValueKind.String)
                        yield return prop.Value.GetString()!;
                    foreach (var s in AllKindAndTypeStrings(prop.Value))
                        yield return s;
                }
                break;
            case JsonValueKind.Array:
                foreach (var item in el.EnumerateArray())
                    foreach (var s in AllKindAndTypeStrings(item))
                        yield return s;
                break;
        }
    }

    [Fact]
    public void DiagnosticJson_EveryConcreteInlineKind_SerializesToKnownKind()
    {
        var inlineTypes = ConcreteSubtypes<IrInline>().ToList();
        Assert.NotEmpty(inlineTypes);

        foreach (var t in inlineTypes)
            AssertNoUnsupportedKind(RenderInline(MinimalInline(t)));
    }

    [Fact]
    public void DiagnosticJson_EveryConcreteBlockKind_SerializesToKnownType()
    {
        var blockTypes = ConcreteSubtypes<IrBlock>().ToList();
        Assert.NotEmpty(blockTypes);

        foreach (var t in blockTypes)
            AssertNoUnsupportedKind(RenderBlock(MinimalBlock(t)));
    }

    [Fact]
    public void DiagnosticJson_InlineImage_PartUriIsRelative_NoFilesystemPathLeaks()
    {
        var img = new IrInlineImage(
            new Uri("/word/media/image1.png", UriKind.Relative), ZeroHash, 100, 200, "alt");
        var json = RenderInline(img);

        using var parsed = JsonDocument.Parse(json);
        var image = parsed.RootElement.GetProperty("scopes")[0].GetProperty("blocks")[0]
            .GetProperty("inlines")[0];
        Assert.Equal("image", image.GetProperty("kind").GetString());
        var partUri = image.GetProperty("partUri").GetString()!;
        Assert.Equal("/word/media/image1.png", partUri);
        // No absolute filesystem path may leak (no drive letter, no file:// scheme, no home dir).
        Assert.DoesNotContain("file:", partUri);
        Assert.DoesNotContain(":\\", partUri);
        Assert.DoesNotContain("/home/", partUri);
    }
}
