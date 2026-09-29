// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="PartNamespaces"/> puts back the namespace declarations an element loses when it is cloned
/// out of one part into another (issue #836). Each test imports a list definition the way the comparison
/// does — <c>new XElement(source)</c> out of a Word-authored numbering part — into a part of its own.
/// </summary>
public class PartNamespacesTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    /// <summary>A numbering part as Word writes it: extension namespaces declared on the root, ignorable.</summary>
    private static readonly XElement WordNumbering = XElement.Parse(
        $"<w:numbering xmlns:mc='{Mc}' xmlns:w='{W}' xmlns:w14='{W14}' xmlns:w15='{W15}' mc:Ignorable='w14 w15'>" +
        "<w:abstractNum w:abstractNumId='1' w15:restartNumberingAfterBreak='0'><w:lvl w:ilvl='0'>" +
        "<mc:AlternateContent><mc:Choice Requires='w14'><w:numFmt w:val='custom' w:format='001, 002, 003, ...'/>" +
        "</mc:Choice><mc:Fallback><w:numFmt w:val='decimal'/></mc:Fallback></mc:AlternateContent>" +
        "</w:lvl></w:abstractNum></w:numbering>");

    private static readonly PartNamespaces Word = PartNamespaces.Of(WordNumbering);

    [Fact]
    public void ImportedExtensionAttribute_IsDeclaredOnTheRootUnderTheSourcePrefixAndListedIgnorable()
    {
        var part = PartWithImportedDefinition($"<w:numbering xmlns:w='{W}'/>");

        Assert.True(Word.DeclareIn(part));

        Assert.Equal(W15, part.GetNamespaceOfPrefix("w15"));
        Assert.Contains("w15", IgnorablePrefixes(part));
    }

    [Fact]
    public void ImportedRequiresPrefix_IsDeclaredOnTheRoot()
    {
        var part = PartWithImportedDefinition($"<w:numbering xmlns:w='{W}'/>");

        Word.DeclareIn(part);

        Assert.Equal(W14, part.GetNamespaceOfPrefix("w14"));
    }

    [Fact]
    public void NamespaceTheRootDeclaresWithoutListingIgnorable_IsListedWhenTheSourceListsIt()
    {
        var part = PartWithImportedDefinition($"<w:numbering xmlns:w='{W}' xmlns:w15='{W15}'/>");

        Assert.True(Word.DeclareIn(part));

        Assert.Contains("w15", IgnorablePrefixes(part));
    }

    [Fact]
    public void PrefixTheRootBindsToAnotherNamespace_IsNeitherReboundNorListedIgnorable()
    {
        var part = PartWithImportedDefinition($"<w:numbering xmlns:w='{W}' xmlns:w15='urn:example:other'/>");

        Word.DeclareIn(part);

        Assert.Equal("urn:example:other", part.GetNamespaceOfPrefix("w15")!.NamespaceName);
        Assert.DoesNotContain("w15", IgnorablePrefixes(part));
    }

    [Fact]
    public void PartWhoseContentAlreadyResolves_IsLeftUntouched()
    {
        var part = new XElement(WordNumbering);
        var before = part.ToString();

        Assert.False(Word.DeclareIn(part));

        Assert.Equal(before, part.ToString());
    }

    private static XElement PartWithImportedDefinition(string rootXml)
    {
        var part = XElement.Parse(rootXml);
        part.Add(new XElement(WordNumbering.Element(W + "abstractNum")!));
        return part;
    }

    private static string[] IgnorablePrefixes(XElement root) =>
        ((string?)root.Attribute(Mc + "Ignorable") ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries);
}
