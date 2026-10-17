using System.Linq;
using System.Xml.Linq;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// The redline's field-context pass (<c>IrMarkupRenderer.NormalizeFieldsInRoot</c>) groups each paragraph's
/// fields by their characters. An <c>end</c> that closes a field opened in an earlier paragraph belongs to no
/// field of this one (issue #1166).
/// </summary>
public class IrFieldNormalizationTests
{
    private static readonly XNamespace Ns = W.w;

    private static XElement Run(XElement child) => new(W.r, child);
    private static XElement FieldChar(string type) => Run(new XElement(W.fldChar, new XAttribute(W.fldCharType, type)));
    private static XElement Text(string text) => Run(new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), text));
    private static XElement Inserted(XElement run) =>
        new(W.ins, new XAttribute(W.id, "1"), new XAttribute(W.author, "A"), run);

    private static XElement Document(params object[] paragraphContent) =>
        new(W.document, new XAttribute(XNamespace.Xmlns + "w", Ns.NamespaceName),
            new XElement(W.body, new XElement(W.p, paragraphContent)));

    /// <summary>A field whose result is unchanged has its plumbing lifted out of an insertion, also when the
    /// paragraph first closes a field opened before it.</summary>
    [Fact]
    public void NormalizeFieldsInRoot_FieldAfterAnEndFromAnEarlierParagraph_LiftsItsPlumbingBare()
    {
        var root = Document(
            FieldChar("end"),
            Text(" text "),
            Inserted(FieldChar("begin")),
            Inserted(Run(new XElement(W.instrText, " PAGE "))),
            Inserted(FieldChar("separate")),
            Text("3"),
            Inserted(FieldChar("end")));

        Assert.True(IrMarkupRenderer.NormalizeFieldsInRoot(root));

        Assert.Empty(root.Descendants(W.ins));
        Assert.Equal(new[] { "end", "begin", "separate", "end" },
            root.Descendants(W.fldChar).Select(f => (string?)f.Attribute(W.fldCharType)));
    }

    /// <summary>The field opened before the paragraph stays as it is: its end is not grouped with anything.</summary>
    [Fact]
    public void NormalizeFieldsInRoot_InsertedEndFromAnEarlierParagraph_StaysInserted()
    {
        var root = Document(Text("Second line"), Inserted(FieldChar("end")));

        Assert.False(IrMarkupRenderer.NormalizeFieldsInRoot(root));

        Assert.Single(root.Descendants(W.ins).Descendants(W.fldChar));
    }
}
