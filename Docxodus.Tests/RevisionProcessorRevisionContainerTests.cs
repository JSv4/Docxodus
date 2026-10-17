// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Rejecting revisions reverses a <c>w:ins</c> or <c>w:del</c> in every container the schema lets hold one, not
/// only in the paragraph-level containers reject used to list (issue #1176): math (<c>m:oMath</c>, the <c>m:e</c>
/// of a math object), ruby (<c>w:rubyBase</c>), and the run-level markup a table holds between its rows or a row
/// between its cells.
/// </summary>
public class RevisionProcessorRevisionContainerTests
{
    private const string Stamp = "w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"";
    private const string MathNs = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace M = MathNs;

    private static string MathRun(string text) => $"<m:r><m:t>{text}</m:t></m:r>";

    private static string Ins(string content) => $"<w:ins w:id=\"1\" {Stamp}>{content}</w:ins>";

    private static string Del(string content) => $"<w:del w:id=\"2\" {Stamp}>{content}</w:del>";

    private const string Table =
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>";

    /// <summary>Containers holding one insertion and one deletion, with the text each review gives.</summary>
    private static readonly Dictionary<string, (string Body, string Rejected, string Accepted)> Containers = new()
    {
        ["math"] = (
            $"<w:p><m:oMath xmlns:m=\"{MathNs}\">{MathRun("x=")}{Ins(MathRun("added"))}{Del(MathRun("original"))}</m:oMath></w:p>",
            "x=original", "x=added"),
        ["the base of a math radical"] = (
            $"<w:p><m:oMath xmlns:m=\"{MathNs}\"><m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/>" +
            $"<m:e>{Ins(MathRun("added"))}{Del(MathRun("original"))}</m:e></m:rad></m:oMath></w:p>",
            "original", "added"),
        ["ruby base text"] = (
            "<w:p><w:r><w:ruby><w:rubyPr><w:rubyAlign w:val=\"center\"/><w:hps w:val=\"10\"/><w:hpsRaise w:val=\"18\"/>" +
            "<w:hpsBaseText w:val=\"20\"/><w:lid w:val=\"ja-JP\"/></w:rubyPr><w:rt><w:r><w:t>guide</w:t></w:r></w:rt><w:rubyBase>" +
            Ins("<w:r><w:t>added</w:t></w:r>") + Del("<w:r><w:delText>original</w:delText></w:r>") +
            "</w:rubyBase></w:ruby></w:r></w:p>",
            "guideoriginal", "guideadded"),
    };

    public static TheoryData<string> ContainerNames() => new(Containers.Keys);

    [Theory]
    [MemberData(nameof(ContainerNames))]
    public void Reject_RevisionsInTheContainer_RestoreTheOriginalText(string container)
    {
        var body = Reject(Containers[container].Body);

        Assert.Equal(Containers[container].Rejected, TextOf(body));
        AssertNoRevisionMarkup(body);
    }

    [Theory]
    [MemberData(nameof(ContainerNames))]
    public void Accept_RevisionsInTheContainer_KeepTheRevisedText(string container)
    {
        var body = Accept(Containers[container].Body);

        Assert.Equal(Containers[container].Accepted, TextOf(body));
        AssertNoRevisionMarkup(body);
    }

    [Fact]
    public void Reject_BookmarkInsertedBetweenTableRows_IsRemoved()
    {
        var bodyXml = Table + Ins("<w:bookmarkStart w:id=\"5\" w:name=\"added\"/>") +
            "<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>";

        Assert.Empty(Reject(bodyXml).Descendants(W.bookmarkStart));
        Assert.Single(Accept(bodyXml).Descendants(W.bookmarkStart));
    }

    [Fact]
    public void Reject_BookmarkDeletedBetweenRowCells_IsRestored()
    {
        var bodyXml = Table + "<w:tr>" + Del("<w:bookmarkStart w:id=\"5\" w:name=\"original\"/>") +
            "<w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>";

        Assert.Equal("original", (string?)Assert.Single(Reject(bodyXml).Descendants(W.bookmarkStart)).Attribute(W.name));
        Assert.Empty(Accept(bodyXml).Descendants(W.bookmarkStart));
    }

    [Fact]
    public void Reject_InsertedAndDeletedTableRows_StillFollowTheirRowMarks()
    {
        // A row's w:ins/w:del in w:trPr marks the row, not content; it is reversed as a mark.
        var bodyXml = Table +
            $"<w:tr><w:trPr><w:ins w:id=\"1\" {Stamp}/></w:trPr><w:tc><w:p><w:ins w:id=\"3\" {Stamp}><w:r><w:t>added</w:t></w:r></w:ins></w:p></w:tc></w:tr>" +
            $"<w:tr><w:trPr><w:del w:id=\"2\" {Stamp}/></w:trPr><w:tc><w:p><w:del w:id=\"4\" {Stamp}><w:r><w:delText>original</w:delText></w:r></w:del></w:p></w:tc></w:tr>" +
            "</w:tbl><w:p/>";

        Assert.Equal("original", TextOf(Reject(bodyXml)));
        Assert.Single(Reject(bodyXml).Descendants(W.tr));
        Assert.Equal("added", TextOf(Accept(bodyXml)));
    }

    private static string TextOf(XElement body) =>
        string.Concat(body.Descendants()
            .Where(e => e.Name == W.t || e.Name == W.delText || e.Name == M + "t")
            .Select(e => e.Value));

    private static void AssertNoRevisionMarkup(XElement body)
    {
        Assert.Empty(body.Descendants(W.ins));
        Assert.Empty(body.Descendants(W.del));
        Assert.Empty(body.Descendants(W.delText));
    }

    private static XElement Reject(string bodyXml) => BodyOf(RevisionProcessor.RejectRevisions(Build(bodyXml)));

    private static XElement Accept(string bodyXml) => BodyOf(RevisionProcessor.AcceptRevisions(Build(bodyXml)));

    private static WmlDocument Build(string bodyXml) => IrTestDocuments.FromBodyXml(bodyXml + "<w:sectPr/>");

    private static XElement BodyOf(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }
}
