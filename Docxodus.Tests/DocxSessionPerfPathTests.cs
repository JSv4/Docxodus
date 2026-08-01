#nullable enable

using System.Linq;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The mutation fast path: <see cref="DocxSessionSettings.EmitMarkdownPatch"/> (skip the
/// per-op scope re-projection for clients that re-render from HTML) and the index-only
/// anchor lookup (mutations must not rebuild the full markdown projection just to
/// resolve an anchor id).
/// </summary>
public class DocxSessionPerfPathTests
{
    private static DocxSession Open(bool emitPatch) =>
        new DocxSession(DocxSession.CreateBlankDocxBytes(),
            new DocxSessionSettings { EmitMarkdownPatch = emitPatch });

    [Fact]
    public void DS300_EmitMarkdownPatchFalse_SkipsPatch_KeepsEverythingElse()
    {
        using var withPatch = Open(true);
        using var noPatch = Open(false);
        var anchorA = withPatch.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:body:"));
        var anchorB = noPatch.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:body:"));

        var a = withPatch.ReplaceText(anchorA, "Hello world.");
        var b = noPatch.ReplaceText(anchorB, "Hello world.");

        Assert.True(a.Success);
        Assert.True(b.Success);
        Assert.NotNull(a.Patch);
        Assert.Null(b.Patch);
        Assert.Equal(a.Modified.Count, b.Modified.Count);
        Assert.Equal(a.Modified[0].Kind, b.Modified[0].Kind);
    }

    [Fact]
    public void DS301_EmitMarkdownPatch_ParsesFromJson()
    {
        var s = Docxodus.Internal.DocxSessionJson.ParseSettings("{\"emitMarkdownPatch\":false}");
        Assert.False(s.EmitMarkdownPatch);
        Assert.True(Docxodus.Internal.DocxSessionJson.ParseSettings("{}").EmitMarkdownPatch);
    }
}
