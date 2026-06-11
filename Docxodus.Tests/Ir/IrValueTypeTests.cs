#nullable enable

using System;
using System.Xml.Linq;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

public class IrValueTypeTests
{
    [Fact]
    public void IrHash_Compute_MatchesKnownSha256Vector()
    {
        var hash = IrHash.Compute("abc");
        Assert.Equal("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", hash.ToHex());
    }

    [Fact]
    public void IrHash_Equality_ByValue()
    {
        var a = IrHash.Compute("hello world");
        var b = IrHash.Compute("hello world");
        var c = IrHash.Compute("goodbye world");

        Assert.Equal(a, b);
        Assert.True(a == b);
        Assert.Equal(a.GetHashCode(), b.GetHashCode());

        Assert.NotEqual(a, c);
        Assert.True(a != c);
    }

    [Fact]
    public void IrAnchor_ToString_MatchesProjectionGrammar()
    {
        const string unid = "0123456789abcdef0123456789abcdef";
        var anchor = new IrAnchor(IrAnchorKind.P, "body", unid);
        Assert.Equal($"p:body:{unid}", anchor.ToString());
    }

    [Fact]
    public void IrAnchor_KindTokens_RoundTrip()
    {
        foreach (IrAnchorKind kind in Enum.GetValues(typeof(IrAnchorKind)))
        {
            var token = IrAnchor.KindToken(kind);
            var roundTripped = IrAnchor.KindFromToken(token);
            Assert.Equal(kind, roundTripped);
        }

        Assert.Throws<ArgumentException>(() => IrAnchor.KindFromToken("nope"));
    }

    [Fact]
    public void IrProvenance_NeverAffectsEquality()
    {
        var p1 = new IrProvenance { Element = new XElement("a"), PartUri = new Uri("/word/document.xml", UriKind.Relative) };
        var p2 = new IrProvenance { Element = new XElement("b"), PartUri = new Uri("/word/header1.xml", UriKind.Relative) };

        Assert.True(p1.Equals(p2));
        Assert.Equal(p1.GetHashCode(), p2.GetHashCode());
        Assert.Equal(0, p1.GetHashCode());
    }
}
