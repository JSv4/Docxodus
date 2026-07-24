#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Reversible-field regressions. Field instructions and begin-state are intentionally transparent to the
/// visible-content hash, so these tests prove they nevertheless reach a structural whole-paragraph fallback.
/// The fixtures are synthetic OOXML, independent of any benchmark corpus.
/// </summary>
public class IrFieldEnvelopeMarkupRendererTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    [Fact]
    public void Read_FieldEnvelopeDigest_DistinguishesCodeAndBeginState_WhileVisibleContentStaysEqual()
    {
        var page = Paragraph(Doc(ComplexField(" PAGE ", "7")));
        var pages = Paragraph(Doc(ComplexField(" NUMPAGES ", "7")));
        var dirty = Paragraph(Doc(ComplexField(" PAGE ", "7", "w:dirty=\"true\"")));

        Assert.Equal(page.ContentHash, pages.ContentHash);
        Assert.Equal(page.ContentHash, dirty.ContentHash);
        Assert.NotEqual(page.FieldEnvelopeDigest, pages.FieldEnvelopeDigest);
        Assert.NotEqual(page.FieldEnvelopeDigest, dirty.FieldEnvelopeDigest);
    }

    [Fact]
    public void Render_ComplexFieldCodeAndStateChange_RoundTripsBothFieldVariants()
    {
        var left = Doc(ComplexField(" PAGE ", "7", "w:dirty=\"true\""));
        var right = Doc(ComplexField(" NUMPAGES ", "7", "w:fldLock=\"true\""));
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        var op = Assert.Single(script.Operations);
        Assert.Equal(IrEditOpKind.ModifyBlock, op.Kind);
        Assert.True(op.RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Equal(Paragraph(right).FieldEnvelopeDigest, Paragraph(accepted).FieldEnvelopeDigest);
        Assert.Equal(Paragraph(left).FieldEnvelopeDigest, Paragraph(rejected).FieldEnvelopeDigest);
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Render_SimpleFieldCodeAndStateChange_PreservesStateDuringRevisionSafeExpansion()
    {
        var left = Doc(SimpleField(" REF old ", "label", "true", "true", "AQID"));
        var right = Doc(SimpleField(" REF new ", "label", "false", "false", "BAUG"));
        Assert.Equal(Paragraph(left).ContentHash, Paragraph(right).ContentHash);
        Assert.NotEqual(Paragraph(left).FieldEnvelopeDigest, Paragraph(right).FieldEnvelopeDigest);
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        AssertExpandedField(accepted, " REF new ", "false", "false", "BAUG");
        AssertExpandedField(rejected, " REF old ", "true", "true", "AQID");
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Render_SimpleFieldCachedResultChange_UsesWholeCarrierFallback()
    {
        var left = Doc(SimpleField(" REF customer ", "Customer A", "false", "false", "AQID"));
        var right = Doc(SimpleField(" REF customer ", "Customer B", "false", "false", "AQID"));
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        var op = Assert.Single(script.Operations);
        Assert.Equal(IrEditOpKind.ModifyBlock, op.Kind);
        Assert.True(op.RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        Assert.Equal("Customer B", string.Concat(MainXml(RevisionProcessor.AcceptRevisions(redline))
            .Descendants(W + "t").Select(element => element.Value)));
        Assert.Equal("Customer A", string.Concat(MainXml(RevisionProcessor.RejectRevisions(redline))
            .Descendants(W + "t").Select(element => element.Value)));
    }

    [Fact]
    public void Render_UnchangedSimpleFieldBeforeEditedSuffix_PreservesFieldAndBothSuffixes()
    {
        var sharedField = SimpleFieldInner(" REF customer ", "Customer", "true", "false", "AQID");
        var left = Doc("<w:p>" + sharedField + "<w:r><w:t> was old</w:t></w:r></w:p>");
        var right = Doc("<w:p>" + sharedField + "<w:r><w:t> was new</w:t></w:r></w:p>");
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        var op = Assert.Single(script.Operations);
        Assert.Equal(IrEditOpKind.ModifyBlock, op.Kind);
        Assert.False(op.RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        AssertDirectSimpleFieldAndText(accepted, " REF customer ", "Customer was new");
        AssertDirectSimpleFieldAndText(rejected, " REF customer ", "Customer was old");
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Render_EmptySimpleFieldBeforeEditedSuffix_PreservesFieldAndBothSuffixes()
    {
        const string sharedField = "<w:fldSimple w:instr=\" REF empty \"/>";
        var left = Doc("<w:p>" + sharedField + "<w:r><w:t> was old</w:t></w:r></w:p>");
        var right = Doc("<w:p>" + sharedField + "<w:r><w:t> was new</w:t></w:r></w:p>");
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        AssertOneFieldInstruction(accepted, " REF empty ");
        AssertOneFieldInstruction(rejected, " REF empty ");
        Assert.Equal(" was new", string.Concat(MainXml(accepted).Descendants(W + "t").Select(element => element.Value)));
        Assert.Equal(" was old", string.Concat(MainXml(rejected).Descendants(W + "t").Select(element => element.Value)));
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Render_ZeroWidthSimpleFieldResultBeforeEditedSuffix_PreservesFieldAndBothSuffixes()
    {
        const string sharedField = "<w:fldSimple w:instr=\" REF tab \"><w:r><w:tab/></w:r></w:fldSimple>";
        var left = Doc("<w:p>" + sharedField + "<w:r><w:t> was old</w:t></w:r></w:p>");
        var right = Doc("<w:p>" + sharedField + "<w:r><w:t> was new</w:t></w:r></w:p>");
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        AssertOneFieldInstruction(accepted, " REF tab ");
        AssertOneFieldInstruction(rejected, " REF tab ");
        Assert.Single(MainXml(accepted).Descendants(W + "tab"));
        Assert.Single(MainXml(rejected).Descendants(W + "tab"));
        Assert.Equal(" was new", string.Concat(MainXml(accepted).Descendants(W + "t").Select(element => element.Value)));
        Assert.Equal(" was old", string.Concat(MainXml(rejected).Descendants(W + "t").Select(element => element.Value)));
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Render_ZeroWidthSimpleFieldBeforeFullyReplacedSuffix_SurvivesBothViews()
    {
        const string sharedField = "<w:fldSimple w:instr=\" REF empty \"/>";
        var left = Doc("<w:p>" + sharedField + "<w:r><w:t>old</w:t></w:r></w:p>");
        var right = Doc("<w:p>" + sharedField + "<w:r><w:t>new</w:t></w:r></w:p>");

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        AssertOneFieldInstruction(accepted, " REF empty ");
        AssertOneFieldInstruction(rejected, " REF empty ");
        Assert.Equal("new", string.Concat(MainXml(accepted).Descendants(W + "t").Select(element => element.Value)));
        Assert.Equal("old", string.Concat(MainXml(rejected).Descendants(W + "t").Select(element => element.Value)));
    }

    [Fact]
    public void Render_SimpleFieldTextboxResultBeforeEditedSuffix_UsesZeroWidthSourceCoordinates()
    {
        const string sharedField =
            "<w:fldSimple w:instr=\" REF textbox \"><w:r><w:pict><v:shape><v:textbox><w:txbxContent>" +
            "<w:p><w:r><w:t>X</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:fldSimple>";
        var left = TextboxFieldDoc(sharedField, " was old");
        var right = TextboxFieldDoc(sharedField, " was new");
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        AssertExpandedTextboxFieldAndText(accepted, " REF textbox ", "X was new");
        AssertExpandedTextboxFieldAndText(rejected, " REF textbox ", "X was old");
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Render_EmptySimpleFieldCodeChange_ExpandsAsOneReversibleCarrier()
    {
        var left = Doc("<w:p><w:fldSimple w:instr=\" REF left \"/></w:p>");
        var right = Doc("<w:p><w:fldSimple w:instr=\" REF right \"/></w:p>");
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        Assert.Empty(MainXml(redline).Descendants(W + "fldSimple"));
        Assert.Equal(" REF right ", FieldInstruction(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(" REF left ", FieldInstruction(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Render_SimpleFieldNoBreakHyphenBeforeEditedSuffix_StaysAligned() =>
        AssertSimpleFieldSpecialVisibleChildBeforeEditedSuffix("noBreakHyphen", "");

    [Fact]
    public void Render_SimpleFieldSoftHyphenBeforeEditedSuffix_StaysAligned() =>
        AssertSimpleFieldSpecialVisibleChildBeforeEditedSuffix("softHyphen", "");

    [Fact]
    public void Render_SimpleFieldValidSymbolBeforeEditedSuffix_StaysAligned() =>
        AssertSimpleFieldSpecialVisibleChildBeforeEditedSuffix("sym", "0041");

    private static void AssertSimpleFieldSpecialVisibleChildBeforeEditedSuffix(string childName, string symbolChar)
    {
        var fieldChild = childName == "sym"
            ? "<w:sym w:font=\"Wingdings\" w:char=\"" + symbolChar + "\"/>"
            : "<w:" + childName + "/>";
        var sharedField = "<w:fldSimple w:instr=\" REF special \"><w:r>" + fieldChild +
            "</w:r></w:fldSimple>";
        var left = Doc("<w:p>" + sharedField + "<w:r><w:t> was old</w:t></w:r></w:p>");
        var right = Doc("<w:p>" + sharedField + "<w:r><w:t> was new</w:t></w:r></w:p>");
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        Assert.False(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        AssertDirectSimpleFieldChildAndText(accepted, " REF special ", childName, symbolChar, " was new");
        AssertDirectSimpleFieldChildAndText(rejected, " REF special ", childName, symbolChar, " was old");
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Render_ComplexFieldCachedResultChange_RemainsFineGrained()
    {
        var left = Doc(ComplexField(" PAGE ", "7"));
        var right = Doc(ComplexField(" PAGE ", "8"));
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        var op = Assert.Single(script.Operations);
        Assert.Equal(IrEditOpKind.ModifyBlock, op.Kind);
        Assert.False(op.RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        Assert.Equal("8", string.Concat(MainXml(RevisionProcessor.AcceptRevisions(redline))
            .Descendants(W + "t").Select(element => element.Value)));
        Assert.Equal("7", string.Concat(MainXml(RevisionProcessor.RejectRevisions(redline))
            .Descendants(W + "t").Select(element => element.Value)));
    }

    [Fact]
    public void Render_NestedSimpleFields_AreExpandedBeforeRevisionWrapping()
    {
        var left = Doc(NestedSimpleFields(" REF outer ", " REF left ", "value"));
        var right = Doc(NestedSimpleFields(" REF outer ", " REF right ", "value"));
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        Assert.Empty(MainXml(redline).Descendants(W + "fldSimple"));

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Equal(" REF outer  REF right ", FieldInstruction(accepted));
        Assert.Equal(" REF outer  REF left ", FieldInstruction(rejected));
        Assert.Equal(2, MainXml(accepted).Descendants(W + "fldChar")
            .Count(element => (string?)element.Attribute(W + "fldCharType") == "begin"));
        Assert.Equal(2, MainXml(rejected).Descendants(W + "fldChar")
            .Count(element => (string?)element.Attribute(W + "fldCharType") == "begin"));
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Render_NestedComplexFieldInstructionChange_UsesWholeCarrierFallback()
    {
        var left = Doc(NestedComplexField(" IF 1 = 1 ", " PAGE ", "7"));
        var right = Doc(NestedComplexField(" IF 1 = 1 ", " NUMPAGES ", "7"));
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        Assert.Equal(" IF 1 = 1  NUMPAGES ", FieldInstruction(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(" IF 1 = 1  PAGE ", FieldInstruction(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Render_InstructionOnlyComplexFieldCodeChange_LeavesOneCodePerRevisionView()
    {
        var left = Doc(InstructionOnlyComplexField(" REF left "));
        var right = Doc(InstructionOnlyComplexField(" REF right "));
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        Assert.Equal(" REF right ", FieldInstruction(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(" REF left ", FieldInstruction(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Read_HyperlinkFieldAndHyperlinkElementRemainCanonicalizedWithoutFieldEnvelope()
    {
        const string target = "https://example.com/";
        var field = IrTestDocuments.FromBodyXmlWithHyperlinks(
            "<w:p><w:fldSimple w:instr=\" HYPERLINK &quot;https://example.com/&quot; \">" +
            "<w:r><w:t>go</w:t></w:r></w:fldSimple></w:p>",
            ("rId1", target));
        var element = IrTestDocuments.FromBodyXmlWithHyperlinks(
            "<w:p><w:hyperlink r:id=\"rId1\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>go</w:t></w:r></w:hyperlink></w:p>",
            ("rId1", target));

        var left = Paragraph(field);
        var right = Paragraph(element);
        Assert.Equal(left.ContentHash, right.ContentHash);
        Assert.Equal(default, left.FieldEnvelopeDigest);
        Assert.Equal(default, right.FieldEnvelopeDigest);

        var op = Assert.Single(IrEditScriptBuilder.Build(IrReader.Read(field), IrReader.Read(element), new IrDiffSettings()).Operations);
        Assert.Equal(IrEditOpKind.EqualBlock, op.Kind);
    }

    [Fact]
    public void Render_SimpleHyperlinkFieldStateChange_PreservesBothFieldStates()
    {
        const string instruction = " HYPERLINK https://example.com/ ";
        var left = Doc(SimpleHyperlinkField(instruction, "go", "w:dirty=\"true\" w:fldLock=\"true\"", "AQID"));
        var right = Doc(SimpleHyperlinkField(instruction, "go", "w:dirty=\"false\" w:fldLock=\"false\"", "BAUG"));
        var leftParagraph = Paragraph(left);
        var rightParagraph = Paragraph(right);

        Assert.Equal(leftParagraph.ContentHash, rightParagraph.ContentHash);
        Assert.NotEqual(leftParagraph.FieldEnvelopeDigest, rightParagraph.FieldEnvelopeDigest);
        var field = Assert.IsType<IrHyperlink>(Assert.Single(leftParagraph.Inlines));
        Assert.True(field.IsFieldHyperlink);
        Assert.True(field.IsSimpleField);

        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        AssertExpandedField(RevisionProcessor.AcceptRevisions(redline), instruction, "false", "false", "BAUG");
        AssertExpandedField(RevisionProcessor.RejectRevisions(redline), instruction, "true", "true", "AQID");
    }

    [Fact]
    public void Render_SimpleHyperlinkFieldCachedResultChange_UsesWholeCarrierFallback()
    {
        const string instruction = " HYPERLINK https://example.com/ ";
        var left = Doc(SimpleHyperlinkField(instruction, "Customer old"));
        var right = Doc(SimpleHyperlinkField(instruction, "Customer new"));

        Assert.Equal(Paragraph(left).FieldEnvelopeDigest, Paragraph(right).FieldEnvelopeDigest);
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        AssertOneFieldInstruction(accepted, instruction);
        AssertOneFieldInstruction(rejected, instruction);
        Assert.Equal("Customer new", string.Concat(MainXml(accepted).Descendants(W + "t").Select(element => element.Value)));
        Assert.Equal("Customer old", string.Concat(MainXml(rejected).Descendants(W + "t").Select(element => element.Value)));
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Render_ComplexHyperlinkFieldSwitchAndCachedResultChange_PreservesBothVariants()
    {
        const string leftInstruction = " HYPERLINK https://example.com/ \\o \"old tooltip\" ";
        const string rightInstruction = " HYPERLINK https://example.com/ \\o \"new tooltip\" ";
        var left = Doc(ComplexHyperlinkField(leftInstruction, "label old"));
        var right = Doc(ComplexHyperlinkField(rightInstruction, "label new"));
        var leftParagraph = Paragraph(left);
        var rightParagraph = Paragraph(right);

        Assert.NotEqual(leftParagraph.ContentHash, rightParagraph.ContentHash);
        Assert.NotEqual(leftParagraph.FieldEnvelopeDigest, rightParagraph.FieldEnvelopeDigest);
        Assert.True(Assert.IsType<IrHyperlink>(Assert.Single(leftParagraph.Inlines)).IsFieldHyperlink);

        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Equal(rightInstruction, FieldInstruction(accepted));
        Assert.Equal(leftInstruction, FieldInstruction(rejected));
        Assert.Equal("label new", string.Concat(MainXml(accepted).Descendants(W + "t").Select(element => element.Value)));
        Assert.Equal("label old", string.Concat(MainXml(rejected).Descendants(W + "t").Select(element => element.Value)));
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Render_ComplexHyperlinkFieldBeginStateChange_PreservesBothFieldStates()
    {
        const string instruction = " HYPERLINK https://example.com/ ";
        var left = Doc(ComplexHyperlinkField(instruction, "go", "w:dirty=\"true\" w:fldLock=\"true\"", "AQID"));
        var right = Doc(ComplexHyperlinkField(instruction, "go", "w:dirty=\"false\" w:fldLock=\"false\"", "BAUG"));
        var leftParagraph = Paragraph(left);
        var rightParagraph = Paragraph(right);

        Assert.Equal(leftParagraph.ContentHash, rightParagraph.ContentHash);
        Assert.NotEqual(leftParagraph.FieldEnvelopeDigest, rightParagraph.FieldEnvelopeDigest);
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        AssertExpandedField(RevisionProcessor.AcceptRevisions(redline), instruction, "false", "false", "BAUG");
        AssertExpandedField(RevisionProcessor.RejectRevisions(redline), instruction, "true", "true", "AQID");
    }

    [Fact]
    public void Render_ComplexHyperlinkFieldUnmodeledSwitchChange_IsNotEqual()
    {
        foreach (var (leftInstruction, rightInstruction) in new[]
        {
            (" HYPERLINK https://example.com/ \\o \"old tooltip\" ",
             " HYPERLINK https://example.com/ \\o \"new tooltip\" "),
            (" HYPERLINK https://example.com/ \\t \"old-frame\" ",
             " HYPERLINK https://example.com/ \\t \"new-frame\" "),
        })
        {
            var left = Doc(ComplexHyperlinkField(leftInstruction, "same label"));
            var right = Doc(ComplexHyperlinkField(rightInstruction, "same label"));
            var leftParagraph = Paragraph(left);
            var rightParagraph = Paragraph(right);

            Assert.Equal(leftParagraph.ContentHash, rightParagraph.ContentHash);
            Assert.NotEqual(leftParagraph.FieldEnvelopeDigest, rightParagraph.FieldEnvelopeDigest);
            var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
            Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

            var redline = DocxDiff.Compare(left, right);
            AssertSchemaValid(redline);
            Assert.Equal(rightInstruction, FieldInstruction(RevisionProcessor.AcceptRevisions(redline)));
            Assert.Equal(leftInstruction, FieldInstruction(RevisionProcessor.RejectRevisions(redline)));
        }
    }

    [Fact]
    public void Render_EmptyHyperlinkFieldsBeforeEditedSuffix_UseWholeCarrierFallback()
    {
        const string instruction = " HYPERLINK https://example.com/ ";
        var simpleLeft = Doc("<w:p>" + SimpleHyperlinkFieldInner(instruction) + "<w:r><w:t>old</w:t></w:r></w:p>");
        var simpleRight = Doc("<w:p>" + SimpleHyperlinkFieldInner(instruction) + "<w:r><w:t>new</w:t></w:r></w:p>");
        var complexLeft = Doc("<w:p>" + ComplexHyperlinkFieldInner(instruction) + "<w:r><w:t>old</w:t></w:r></w:p>");
        var complexRight = Doc("<w:p>" + ComplexHyperlinkFieldInner(instruction) + "<w:r><w:t>new</w:t></w:r></w:p>");

        foreach (var (left, right) in new[] { (simpleLeft, simpleRight), (complexLeft, complexRight) })
        {
            var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
            Assert.True(Assert.Single(script.Operations).RequiresWholeParagraphReplace);

            var redline = DocxDiff.Compare(left, right);
            AssertSchemaValid(redline);
            var accepted = RevisionProcessor.AcceptRevisions(redline);
            var rejected = RevisionProcessor.RejectRevisions(redline);
            AssertOneFieldInstruction(accepted, instruction);
            AssertOneFieldInstruction(rejected, instruction);
            Assert.Equal("new", string.Concat(MainXml(accepted).Descendants(W + "t").Select(element => element.Value)));
            Assert.Equal("old", string.Concat(MainXml(rejected).Descendants(W + "t").Select(element => element.Value)));
            AssertNoRevisionMarkup(accepted);
            AssertNoRevisionMarkup(rejected);
        }
    }

    [Fact]
    public void Render_HeaderFieldBoundaryBeforeEditedText_RoundTripsInBothViews()
    {
        const string body = "<w:p><w:r><w:t>Body.</w:t></w:r></w:p>";
        const string instruction = " REF header ";
        var left = IrTestDocuments.FromBodyAndHeaderXml(
            body,
            "<w:p><w:r><w:t>old </w:t></w:r>" + ComplexFieldInner(instruction, "value") + "</w:p>");
        var right = IrTestDocuments.FromBodyAndHeaderXml(
            body,
            "<w:p><w:r><w:t>new </w:t></w:r>" + ComplexFieldInner(instruction, "value") + "</w:p>");

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        var accepted = PartXml(RevisionProcessor.AcceptRevisions(redline), main => main.HeaderParts.First());
        var rejected = PartXml(RevisionProcessor.RejectRevisions(redline), main => main.HeaderParts.First());
        AssertOneFieldInstruction(accepted, instruction);
        AssertOneFieldInstruction(rejected, instruction);
        Assert.Equal("new value", string.Concat(accepted.Descendants(W + "t").Select(element => element.Value)));
        Assert.Equal("old value", string.Concat(rejected.Descendants(W + "t").Select(element => element.Value)));
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Render_FootnoteFieldBoundaryBeforeEditedText_RoundTripsInBothViews()
    {
        const string body = "<w:p><w:r><w:t>Body.</w:t></w:r></w:p>";
        const string instruction = " REF footnote ";
        var left = IrTestDocuments.FromBodyXmlWithFootnoteParagraph(
            body,
            "<w:r><w:t>old </w:t></w:r>" + ComplexFieldInner(instruction, "value"));
        var right = IrTestDocuments.FromBodyXmlWithFootnoteParagraph(
            body,
            "<w:r><w:t>new </w:t></w:r>" + ComplexFieldInner(instruction, "value"));

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        var accepted = PartXml(RevisionProcessor.AcceptRevisions(redline), main => main.FootnotesPart);
        var rejected = PartXml(RevisionProcessor.RejectRevisions(redline), main => main.FootnotesPart);
        AssertOneFieldInstruction(accepted, instruction);
        AssertOneFieldInstruction(rejected, instruction);
        Assert.Equal("new value", string.Concat(accepted.Descendants(W + "t").Select(element => element.Value)));
        Assert.Equal("old value", string.Concat(rejected.Descendants(W + "t").Select(element => element.Value)));
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    [Fact]
    public void Build_TextboxFieldCodeChange_PropagatesInnerStructuralCarrierToOuterParagraph()
    {
        var left = TextboxDoc(ComplexField(" PAGE ", "7"));
        var right = TextboxDoc(ComplexField(" NUMPAGES ", "7"));
        var leftOuter = Paragraph(left);
        var rightOuter = Paragraph(right);

        // The textbox's visible cached result is unchanged, so ordinary content identity stays stable. Its
        // inner FieldEnvelopeDigest must nevertheless reach the owning paragraph's structural carrier summary.
        Assert.Equal(leftOuter.ContentHash, rightOuter.ContentHash);
        Assert.NotEqual(leftOuter.FieldEnvelopeDigest, rightOuter.FieldEnvelopeDigest);

        var op = Assert.Single(IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings()).Operations);
        Assert.Equal(IrEditOpKind.ModifyBlock, op.Kind);
        Assert.True(op.RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        Assert.Equal(" NUMPAGES ", FieldInstruction(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(" PAGE ", FieldInstruction(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Render_TextboxInlineEnvelopeChange_PropagatesToOwningParagraph()
    {
        var left = TextboxDoc(InlineSdt("left", "same"));
        var right = TextboxDoc(InlineSdt("right", "same"));
        var leftOuter = Paragraph(left);
        var rightOuter = Paragraph(right);

        Assert.Equal(leftOuter.ContentHash, rightOuter.ContentHash);
        Assert.NotEqual(leftOuter.InlineEnvelopeDigest, rightOuter.InlineEnvelopeDigest);

        var op = Assert.Single(IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings()).Operations);
        Assert.Equal(IrEditOpKind.ModifyBlock, op.Kind);
        Assert.True(op.RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        Assert.Equal("right", InlineSdtTag(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("left", InlineSdtTag(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Render_TableCellFieldCodeChange_PropagatesThroughCellRowAndTable()
    {
        var left = TableDoc(ComplexField(" PAGE ", "7"));
        var right = TableDoc(ComplexField(" NUMPAGES ", "7"));
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());

        var op = Assert.Single(script.Operations);
        Assert.Equal(IrEditOpKind.ModifyBlock, op.Kind);
        Assert.NotNull(op.TableDiff);
        var cellOp = Assert.Single(Assert.Single(op.TableDiff!.RowOps).CellOps!);
        Assert.NotNull(cellOp.BlockOps);
        Assert.True(Assert.Single(cellOp.BlockOps!).RequiresWholeParagraphReplace);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        Assert.Equal(" NUMPAGES ", FieldInstruction(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(" PAGE ", FieldInstruction(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Consolidate_DivergentSameVisibleFieldCodes_RecordAConflictInsteadOfFalseConsensus()
    {
        var baseDoc = Doc(ComplexField(" PAGE ", "7"));
        var alice = Doc(ComplexField(" NUMPAGES ", "7"));
        var bob = Doc(ComplexField(" SECTIONPAGES ", "7"));
        var reviewers = new[]
        {
            new DocxDiffReviewer { Author = "Alice", Document = alice },
            new DocxDiffReviewer { Author = "Bob", Document = bob },
        };
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = ConflictResolution.FirstReviewerWins };

        var merged = DocxDiff.Consolidate(baseDoc, reviewers, settings);
        AssertSchemaValid(merged);
        Assert.Single(DocxDiff.GetConflicts(baseDoc, reviewers, settings));
        Assert.Equal(" NUMPAGES ", FieldInstruction(RevisionProcessor.AcceptRevisions(merged)));
        Assert.Equal(" PAGE ", FieldInstruction(RevisionProcessor.RejectRevisions(merged)));
        Assert.Contains("\"requiresWholeParagraphReplace\"", DocxDiff.GetConsolidatedEditScriptJson(baseDoc, reviewers, settings));
    }

    private static WmlDocument Doc(string body) => IrTestDocuments.FromBodyXml(body);

    private static IrParagraph Paragraph(WmlDocument doc) =>
        IrReader.Read(doc).Body.Blocks.OfType<IrParagraph>().Single();

    private static string ComplexField(string instruction, string result, string beginAttributes = "") =>
        "<w:p>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\" " + beginAttributes + "/></w:r>" +
        "<w:r><w:instrText xml:space=\"preserve\">" + instruction + "</w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        "<w:r><w:t>" + result + "</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
        "</w:p>";

    private static string ComplexFieldInner(string instruction, string result) =>
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        "<w:r><w:instrText xml:space=\"preserve\">" + instruction + "</w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        "<w:r><w:t>" + result + "</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    private static string SimpleField(string instruction, string result, string dirty, string locked, string fieldData) =>
        "<w:p>" + SimpleFieldInner(instruction, result, dirty, locked, fieldData) + "</w:p>";

    private static string SimpleFieldInner(string instruction, string result, string dirty, string locked, string fieldData) =>
        "<w:fldSimple w:instr=\"" + instruction + "\" w:dirty=\"" + dirty + "\" w:fldLock=\"" + locked + "\">" +
        "<w:fldData>" + fieldData + "</w:fldData><w:r><w:t>" + result + "</w:t></w:r></w:fldSimple>";

    private static string SimpleHyperlinkField(string instruction, string result, string attributes = "", string fieldData = "") =>
        "<w:p>" + SimpleHyperlinkFieldInner(instruction, result, attributes, fieldData) + "</w:p>";

    private static string SimpleHyperlinkFieldInner(
        string instruction, string? result = null, string attributes = "", string fieldData = "") =>
        "<w:fldSimple w:instr=\"" + XmlAttribute(instruction) + "\"" +
        (string.IsNullOrEmpty(attributes) ? "" : " " + attributes) + ">" +
        (string.IsNullOrEmpty(fieldData) ? "" : "<w:fldData>" + fieldData + "</w:fldData>") +
        (result is null ? "" : "<w:r><w:t>" + result + "</w:t></w:r>") +
        "</w:fldSimple>";

    private static string ComplexHyperlinkField(string instruction, string result, string beginAttributes = "", string fieldData = "") =>
        "<w:p>" + ComplexHyperlinkFieldInner(instruction, result, beginAttributes, fieldData) + "</w:p>";

    private static string ComplexHyperlinkFieldInner(
        string instruction, string? result = null, string beginAttributes = "", string fieldData = "") =>
        "<w:r><w:fldChar w:fldCharType=\"begin\"" +
        (string.IsNullOrEmpty(beginAttributes) ? "" : " " + beginAttributes) + ">" +
        (string.IsNullOrEmpty(fieldData) ? "" : "<w:fldData>" + fieldData + "</w:fldData>") +
        "</w:fldChar></w:r>" +
        "<w:r><w:instrText xml:space=\"preserve\">" + instruction + "</w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        (result is null ? "" : "<w:r><w:t>" + result + "</w:t></w:r>") +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    private static string XmlAttribute(string value) =>
        value.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;");

    private static string NestedSimpleFields(string outerInstruction, string innerInstruction, string result) =>
        "<w:p><w:fldSimple w:instr=\"" + outerInstruction + "\"><w:fldSimple w:instr=\"" +
        innerInstruction + "\"><w:r><w:t>" + result + "</w:t></w:r></w:fldSimple></w:fldSimple></w:p>";

    private static string NestedComplexField(string outerInstruction, string innerInstruction, string result) =>
        "<w:p>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        "<w:r><w:instrText xml:space=\"preserve\">" + outerInstruction + "</w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        "<w:r><w:instrText xml:space=\"preserve\">" + innerInstruction + "</w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        "<w:r><w:t>" + result + "</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
        "</w:p>";

    private static string InstructionOnlyComplexField(string instruction) =>
        "<w:p>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        "<w:r><w:instrText xml:space=\"preserve\">" + instruction + "</w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
        "</w:p>";

    private static string InlineSdt(string tag, string text) =>
        "<w:p><w:sdt><w:sdtPr><w:tag w:val=\"" + tag + "\"/></w:sdtPr><w:sdtContent>" +
        "<w:r><w:t>" + text + "</w:t></w:r></w:sdtContent></w:sdt></w:p>";

    private static WmlDocument TextboxDoc(string innerParagraph) =>
        IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent>" + innerParagraph +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>");

    private static WmlDocument TextboxFieldDoc(string field, string suffix) =>
        IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p>" + field + "<w:r><w:t>" + suffix + "</w:t></w:r></w:p>");

    private static WmlDocument TableDoc(string paragraph) =>
        Doc("<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:tcPr/>" + paragraph + "</w:tc></w:tr></w:tbl>");

    private static void AssertExpandedField(WmlDocument doc, string instruction, string dirty, string locked, string fieldData)
    {
        var main = MainXml(doc);
        Assert.Empty(main.Descendants(W + "fldSimple"));
        Assert.Equal(instruction, FieldInstruction(doc));
        var begin = Assert.Single(main.Descendants(W + "fldChar").Where(element =>
            (string?)element.Attribute(W + "fldCharType") == "begin"));
        Assert.Equal(dirty, (string?)begin.Attribute(W + "dirty"));
        Assert.Equal(locked, (string?)begin.Attribute(W + "fldLock"));
        Assert.Equal(fieldData, (string?)begin.Element(W + "fldData"));
    }

    private static void AssertDirectSimpleFieldAndText(WmlDocument doc, string instruction, string text)
    {
        var main = MainXml(doc);
        var field = Assert.Single(main.Descendants(W + "fldSimple"));
        Assert.Equal(instruction, (string?)field.Attribute(W + "instr"));
        Assert.Equal(text, string.Concat(main.Descendants(W + "t").Select(element => element.Value)));
    }

    private static void AssertDirectSimpleFieldChildAndText(
        WmlDocument doc, string instruction, string childName, string symbolChar, string suffixText)
    {
        var main = MainXml(doc);
        var field = Assert.Single(main.Descendants(W + "fldSimple"));
        Assert.Equal(instruction, (string?)field.Attribute(W + "instr"));
        var child = Assert.Single(field.Descendants(W + childName));
        if (childName == "sym")
            Assert.Equal(symbolChar, (string?)child.Attribute(W + "char"));
        Assert.Equal(suffixText, string.Concat(main.Descendants(W + "t").Select(element => element.Value)));
    }

    private static void AssertExpandedTextboxFieldAndText(WmlDocument doc, string instruction, string text)
    {
        var main = MainXml(doc);
        Assert.Empty(main.Descendants(W + "fldSimple"));
        Assert.Equal(instruction, FieldInstruction(doc));
        Assert.Single(main.Descendants(W + "txbxContent"));
        Assert.Equal(text, string.Concat(main.Descendants(W + "t").Select(element => element.Value)));
    }

    private static void AssertOneFieldInstruction(WmlDocument doc, string instruction)
    {
        AssertOneFieldInstruction(MainXml(doc).Root!, instruction);
    }

    private static void AssertOneFieldInstruction(XElement root, string instruction)
    {
        var simple = root.Descendants(W + "fldSimple").Select(element => (string?)element.Attribute(W + "instr"))
            .Where(value => value is not null).Cast<string>().ToList();
        var complex = root.Descendants(W + "instrText").Select(element => element.Value).ToList();
        Assert.True(
            (simple.Count == 1 && simple[0] == instruction && complex.Count == 0) ||
            (simple.Count == 0 && complex.Count == 1 && complex[0] == instruction),
            "Expected exactly one surviving simple or complex field instruction. Simple=[" +
            string.Join("|", simple) + "] Complex=[" + string.Join("|", complex) + "]");
    }

    private static string FieldInstruction(WmlDocument doc) =>
        string.Concat(MainXml(doc).Descendants(W + "instrText").Select(element => element.Value));

    private static string InlineSdtTag(WmlDocument doc) =>
        (string?)Assert.Single(MainXml(doc).Descendants(W + "tag")).Attribute(W + "val") ?? "";

    private static void AssertSchemaValid(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(wdoc)
            .Select(error => $"{error.Id}@{error.Path?.XPath}: {error.Description}")
            .ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    private static void AssertNoRevisionMarkup(WmlDocument doc)
    {
        AssertNoRevisionMarkup(MainXml(doc).Root!);
    }

    private static void AssertNoRevisionMarkup(XElement root)
    {
        var revisions = new HashSet<XName>
        {
            W + "ins", W + "del", W + "moveFrom", W + "moveTo",
            W + "moveFromRangeStart", W + "moveFromRangeEnd", W + "moveToRangeStart", W + "moveToRangeEnd",
        };
        Assert.DoesNotContain(root.Descendants(), element => revisions.Contains(element.Name));
    }

    private static XElement PartXml(WmlDocument doc, System.Func<MainDocumentPart, OpenXmlPart?> pick)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var part = pick(wdoc.MainDocumentPart!)!;
        using var partStream = part.GetStream();
        return XElement.Load(partStream);
    }

    private static XDocument MainXml(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.GetStream());
        return XDocument.Parse(reader.ReadToEnd());
    }
}
