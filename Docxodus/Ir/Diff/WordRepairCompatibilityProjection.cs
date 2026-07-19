#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Detects the very narrow Word repair cascade for which Microsoft Word Compare invents a visual replay of
/// unchanged content. This is deliberately a raw-package detector: the normal IR correctly ignores regenerated
/// <c>w14:paraId</c>/<c>w:rsid*</c> bookkeeping, while this opt-in compatibility projection must recognize the
/// producer artifact without making the general-purpose aligner identity-sensitive.
/// </summary>
/// <remarks>
/// The guards are intentionally conjunctive and conservative. A pair must have no accepted-content edit in any
/// modeled story, matching pre-existing inline revisions, at least 64 raw para ids with at least 50% positional
/// churn, two or more tables whose only delta is the known shell cleanup, and at least eight style definitions
/// whose only delta is an implicit <c>Normal</c>/<c>DefaultParagraphFont</c> inheritance default. A normal
/// formatting edit, a true body edit, or a routine Word save therefore remains on the ordinary renderer path.
/// </remarks>
internal static class WordRepairCompatibilityProjection
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

    private const int MinimumPairedParaIds = 64;
    private const int MinimumChangedStyles = 8;
    private const int MinimumChangedTables = 2;

    /// <summary>
    /// Returns true only when the body should be rendered as paired whole-block deletion/insertion replacements.
    /// This does not alter <paramref name="script"/>: JSON and alignment consumers retain the semantic IR truth.
    /// </summary>
    internal static bool ShouldRenderWholeBodyReplacement(
        IrEditScript script,
        IrDocument acceptedLeft,
        IrDocument acceptedRight,
        WmlDocument rawLeft,
        WmlDocument rawRight)
    {
        ArgumentNullException.ThrowIfNull(script);
        ArgumentNullException.ThrowIfNull(acceptedLeft);
        ArgumentNullException.ThrowIfNull(acceptedRight);
        ArgumentNullException.ThrowIfNull(rawLeft);
        ArgumentNullException.ThrowIfNull(rawRight);

        if (!HasNoSemanticEdit(script, acceptedLeft, acceptedRight))
            return false;

        using var leftStream = new OpenXmlMemoryStreamDocument(rawLeft);
        using var rightStream = new OpenXmlMemoryStreamDocument(rawRight);
        using var leftDoc = leftStream.GetWordprocessingDocument();
        using var rightDoc = rightStream.GetWordprocessingDocument();

        var leftMain = leftDoc.MainDocumentPart;
        var rightMain = rightDoc.MainDocumentPart;
        var leftBody = leftMain?.GetXDocument().Root?.Element(W + "body");
        var rightBody = rightMain?.GetXDocument().Root?.Element(W + "body");
        if (leftBody == null || rightBody == null)
            return false;

        if (!HasMatchingPreexistingInsertionsAndDeletions(leftBody, rightBody))
            return false;
        if (!HasPervasiveParaIdChurn(leftBody, rightBody))
            return false;
        if (!StylesOnlyMaterializeImplicitDefaults(leftMain?.StyleDefinitionsPart, rightMain?.StyleDefinitionsPart))
            return false;

        return HasOnlyKnownTableShellRepair(leftBody, rightBody);
    }

    private static bool HasNoSemanticEdit(IrEditScript script, IrDocument left, IrDocument right)
    {
        // The body is the only scope this projection rewrites. Any changed note/header/footer story is an
        // ordinary semantic edit and must keep the regular renderer path.
        // The builder normally represents no changed auxiliary stories as null, but callers which construct
        // a script directly can supply an empty immutable list. Treat the two representations identically.
        if (script.NoteOps is { Count: > 0 } || script.HeaderFooterOps is { Count: > 0 })
            return false;

        var leftBlocks = left.Body.Blocks;
        var rightBlocks = right.Body.Blocks;
        if (leftBlocks.Count != rightBlocks.Count || leftBlocks.Count == 0)
            return false;

        for (int i = 0; i < leftBlocks.Count; i++)
        {
            if (leftBlocks[i].GetType() != rightBlocks[i].GetType() ||
                !leftBlocks[i].ContentHash.Equals(rightBlocks[i].ContentHash))
            {
                return false;
            }
        }

        // A genuine edit produces a non-paired body op. The detector allows Equal and FormatOnly only: the
        // latter is exactly the table-shell normalization this compatibility mode projects visibly.
        return script.Operations.All(op =>
            (op.Kind == IrEditOpKind.EqualBlock || op.Kind == IrEditOpKind.FormatOnlyBlock) &&
            op.LeftAnchor != null && op.RightAnchor != null);
    }

    private static bool HasMatchingPreexistingInsertionsAndDeletions(XElement leftBody, XElement rightBody)
    {
        var left = RevisionElements(leftBody).ToList();
        var right = RevisionElements(rightBody).ToList();
        if (left.Count == 0 || left.Count != right.Count ||
            !left.Any(e => e.Name == W + "ins") || !left.Any(e => e.Name == W + "del"))
        {
            return false;
        }

        for (int i = 0; i < left.Count; i++)
        {
            var a = new XElement(left[i]);
            var b = new XElement(right[i]);
            // Revision id/author/date are input facts rather than repair bookkeeping. Require them to
            // match too; only para/text ids and rsids are regenerated Word save metadata.
            StripVolatileIdentity(a, stripRevisionMetadata: false);
            StripVolatileIdentity(b, stripRevisionMetadata: false);
            if (!XNode.DeepEquals(a, b))
                return false;
        }
        return true;
    }

    private static IEnumerable<XElement> RevisionElements(XElement body) =>
        body.DescendantsAndSelf().Where(e => e.Name == W + "ins" || e.Name == W + "del");

    private static bool HasPervasiveParaIdChurn(XElement leftBody, XElement rightBody)
    {
        var leftIds = leftBody.DescendantsAndSelf().Attributes(W14 + "paraId").Select(a => a.Value).ToList();
        var rightIds = rightBody.DescendantsAndSelf().Attributes(W14 + "paraId").Select(a => a.Value).ToList();
        if (leftIds.Count < MinimumPairedParaIds || leftIds.Count != rightIds.Count)
            return false;

        int changed = 0;
        for (int i = 0; i < leftIds.Count; i++)
            if (!string.Equals(leftIds[i], rightIds[i], StringComparison.Ordinal))
                changed++;

        return changed * 2 >= leftIds.Count;
    }

    private static bool StylesOnlyMaterializeImplicitDefaults(
        StyleDefinitionsPart? leftStyles,
        StyleDefinitionsPart? rightStyles)
    {
        var leftRoot = leftStyles?.GetXDocument().Root;
        var rightRoot = rightStyles?.GetXDocument().Root;
        if (leftRoot == null || rightRoot == null)
            return false;

        var leftById = leftRoot.Elements(W + "style")
            .Where(e => e.Attribute(W + "styleId") != null)
            .ToDictionary(e => (string)e.Attribute(W + "styleId")!, StringComparer.Ordinal);
        var rightById = rightRoot.Elements(W + "style")
            .Where(e => e.Attribute(W + "styleId") != null)
            .ToDictionary(e => (string)e.Attribute(W + "styleId")!, StringComparer.Ordinal);
        if (leftById.Count != rightById.Count || !leftById.Keys.OrderBy(x => x, StringComparer.Ordinal)
                .SequenceEqual(rightById.Keys.OrderBy(x => x, StringComparer.Ordinal), StringComparer.Ordinal))
        {
            return false;
        }

        int changed = 0;
        foreach (var styleId in leftById.Keys)
            if (!XNode.DeepEquals(leftById[styleId], rightById[styleId]))
                changed++;
        if (changed < MinimumChangedStyles)
            return false;

        var leftNormalized = new XElement(leftRoot);
        var rightNormalized = new XElement(rightRoot);
        StripImplicitStyleDefaults(leftNormalized);
        StripImplicitStyleDefaults(rightNormalized);
        StripWhitespaceOnlyText(leftNormalized);
        StripWhitespaceOnlyText(rightNormalized);

        // Word's repair pass can reorder otherwise unchanged style definitions (the captured repair pair
        // moves CommentReference ahead of CommentText). Ordering is not semantic, but every non-style root
        // child and every normalized definition still has to match exactly.
        var leftNonStyles = new XElement(leftNormalized);
        var rightNonStyles = new XElement(rightNormalized);
        // Keep an invalid/anonymous direct w:style in this exact comparison rather than silently ignoring it.
        // The keyed comparison below deliberately uses the same styleId filter as the initial maps.
        leftNonStyles.Elements(W + "style").Where(e => e.Attribute(W + "styleId") != null).Remove();
        rightNonStyles.Elements(W + "style").Where(e => e.Attribute(W + "styleId") != null).Remove();
        if (!XNode.DeepEquals(leftNonStyles, rightNonStyles))
            return false;

        var normalizedLeftById = leftNormalized.Elements(W + "style")
            .Where(e => e.Attribute(W + "styleId") != null)
            .ToDictionary(e => (string)e.Attribute(W + "styleId")!, StringComparer.Ordinal);
        var normalizedRightById = rightNormalized.Elements(W + "style")
            .Where(e => e.Attribute(W + "styleId") != null)
            .ToDictionary(e => (string)e.Attribute(W + "styleId")!, StringComparer.Ordinal);
        return leftById.Keys.All(styleId =>
            XNode.DeepEquals(normalizedLeftById[styleId], normalizedRightById[styleId]));
    }

    private static void StripImplicitStyleDefaults(XElement root)
    {
        foreach (var style in root.Elements(W + "style"))
        {
            foreach (var child in style.Elements().Where(e =>
                         (e.Name == W + "basedOn" || e.Name == W + "next") &&
                         IsImplicitDefaultReference(e)).ToList())
            {
                child.Remove();
            }
        }
    }

    private static void StripWhitespaceOnlyText(XElement root)
    {
        // styles.xml has no semantically meaningful text nodes in this comparison; Word may reindent it
        // while repairing a package, and XNode.DeepEquals intentionally treats that serialization trivia as
        // different. Restrict this normalization to the copied styles roots above.
        foreach (var text in root.DescendantNodesAndSelf().OfType<XText>()
                     .Where(text => string.IsNullOrWhiteSpace(text.Value)).ToList())
        {
            text.Remove();
        }
    }

    private static bool IsImplicitDefaultReference(XElement element)
    {
        if (element.Attributes().Any(a => a.Name != W + "val"))
            return false;
        var value = (string?)element.Attribute(W + "val");
        return value is "Normal" or "DefaultParagraphFont";
    }

    private static bool HasOnlyKnownTableShellRepair(XElement rawLeftBody, XElement rawRightBody)
    {
        var left = new XElement(rawLeftBody);
        var right = new XElement(rawRightBody);
        StripVolatileIdentity(left, stripRevisionMetadata: true);
        StripVolatileIdentity(right, stripRevisionMetadata: true);

        var leftTables = left.Descendants(W + "tbl").ToList();
        var rightTables = right.Descendants(W + "tbl").ToList();
        if (leftTables.Count != rightTables.Count)
            return false;

        int changedTables = 0;
        for (int i = 0; i < leftTables.Count; i++)
            if (!XNode.DeepEquals(leftTables[i], rightTables[i]))
                changedTables++;
        if (changedTables < MinimumChangedTables)
            return false;

        StripKnownTableRepairArtifacts(left);
        StripKnownTableRepairArtifacts(right);
        return XNode.DeepEquals(left, right);
    }

    private static void StripVolatileIdentity(XElement root, bool stripRevisionMetadata)
    {
        foreach (var element in root.DescendantsAndSelf())
        {
            foreach (var attr in element.Attributes().ToList())
            {
                if ((attr.Name.Namespace == W14 && (attr.Name.LocalName == "paraId" || attr.Name.LocalName == "textId")) ||
                    (attr.Name.Namespace == W && attr.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)))
                {
                    attr.Remove();
                }
            }

            if (!stripRevisionMetadata || (element.Name != W + "ins" && element.Name != W + "del"))
                continue;
            element.Attribute(W + "id")?.Remove();
            element.Attribute(W + "author")?.Remove();
            element.Attribute(W + "date")?.Remove();
        }
    }

    private static void StripKnownTableRepairArtifacts(XElement body)
    {
        foreach (var tblPr in body.Descendants(W + "tblPr"))
        {
            foreach (var child in tblPr.Elements().ToList())
            {
                if ((child.Name == W + "tblStyle" && IsNormalTableStyle(child)) ||
                    (child.Name == W + "tblInd" && IsRepairTableIndent(child)) ||
                    (child.Name == W + "tblCellMar" && IsRepairCellMargins(child)))
                {
                    child.Remove();
                }
            }
        }

        foreach (var tblPrEx in body.Descendants(W + "tblPrEx").ToList())
        {
            var children = tblPrEx.Elements().ToList();
            if (!tblPrEx.HasAttributes && children.Count == 1 && children[0].Name == W + "tblCellMar" &&
                IsRepairCellMargins(children[0]))
            {
                tblPrEx.Remove();
            }
        }
    }

    private static bool IsNormalTableStyle(XElement element) =>
        !element.HasElements && element.Attributes().Count() == 1 &&
        string.Equals((string?)element.Attribute(W + "val"), "Normal", StringComparison.Ordinal);

    private static bool IsRepairTableIndent(XElement element)
    {
        if (element.HasElements || element.Attributes().Any(a => a.Name != W + "w" && a.Name != W + "type"))
            return false;
        if (!string.Equals((string?)element.Attribute(W + "type"), "dxa", StringComparison.Ordinal))
            return false;
        return (string?)element.Attribute(W + "w") is "-4" or "5";
    }

    private static bool IsRepairCellMargins(XElement margins)
    {
        if (margins.HasAttributes)
            return false;
        var children = margins.Elements().ToList();
        if (children.Count != 2 || children.Any(c => c.HasElements ||
            c.Attributes().Any(a => a.Name != W + "w" && a.Name != W + "type") ||
            !string.Equals((string?)c.Attribute(W + "type"), "dxa", StringComparison.Ordinal)))
        {
            return false;
        }

        bool horizontal = children.All(c => c.Name is var n && (n == W + "left" || n == W + "right") &&
            string.Equals((string?)c.Attribute(W + "w"), "10", StringComparison.Ordinal));
        bool vertical = children.All(c => c.Name is var n && (n == W + "top" || n == W + "bottom") &&
            string.Equals((string?)c.Attribute(W + "w"), "0", StringComparison.Ordinal));
        return horizontal || vertical;
    }
}
