using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Internal;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class CommentDateRenderingTests
{
    private const string Legacy = "2026-09-09T23:45:00Z";
    private const string Modern = "2026-09-10T04:45:00Z";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace Cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
    private static readonly XNamespace Cex = "http://schemas.microsoft.com/office/word/2018/wordml/cex";

    // Original valid comment graph. Only the last comment paragraph identifies the durable record.
    private static byte[] Fixture(string? utc = Modern, bool matched = true,
        bool leadingParagraph = false, bool validate = true)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(new Wp.Paragraph(
                new Wp.CommentRangeStart { Id = "6" },
                new Wp.Run(new Wp.Text("Cedar date anchor.")),
                new Wp.CommentRangeEnd { Id = "6" },
                new Wp.Run(new Wp.CommentReference { Id = "6" })), new Wp.SectionProperties()));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            var comment = new XElement(W + "comment", new XAttribute(W + "id", "6"),
                new XAttribute(W + "author", "Cedar Date Reviewer"), new XAttribute(W + "date", Legacy));
            if (leadingParagraph)
                comment.Add(new XElement(W + "p", new XAttribute(W14 + "paraId", "00CC00CC"),
                    new XElement(W + "r", new XElement(W + "t", "Cedar date preface."))));
            comment.Add(new XElement(W + "p", new XAttribute(W14 + "paraId", "00AA00AA"),
                new XElement(W + "r", new XElement(W + "t", "DateNote Check the timestamp."))));
            main.AddNewPart<WordprocessingCommentsPart>().PutXDocument(new XDocument(
                new XElement(W + "comments", new XAttribute(XNamespace.Xmlns + "w", W),
                    new XAttribute(XNamespace.Xmlns + "w14", W14),
                    new XAttribute(XNamespace.Xmlns + "mc", MC.mc), new XAttribute(MC.Ignorable, "w14"), comment)));
            if (utc != null)
            {
                main.AddNewPart<WordprocessingCommentsIdsPart>().PutXDocument(new XDocument(
                    new XElement(Cid + "commentsIds", new XAttribute(XNamespace.Xmlns + "w16cid", Cid),
                        new XElement(Cid + "commentId",
                            new XAttribute(Cid + "paraId", matched ? "00AA00AA" : "00000099"),
                            new XAttribute(Cid + "durableId", "02BB02BB")))));
                main.AddNewPart<WordCommentsExtensiblePart>().PutXDocument(new XDocument(
                    new XElement(Cex + "commentsExtensible", new XAttribute(XNamespace.Xmlns + "w16cex", Cex),
                        new XElement(Cex + "commentExtensible", new XAttribute(Cex + "durableId", "02BB02BB"),
                            new XAttribute(Cex + "dateUtc", utc)))));
            }
            doc.Save();
            if (validate) Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(doc));
        }
        return stream.ToArray();
    }

    private static XElement Render(byte[] bytes, CommentRenderMode mode) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-dates.docx", bytes),
            new WmlToHtmlConverterSettings { RenderComments = true, CommentRenderMode = mode,
                IncludeCommentMetadata = true, FabricateCssClasses = false });

    private static void DateIs(XElement html, string expected) =>
        Assert.Equal(new[] { expected }, html.DescendantsAndSelf().Attributes("data-date")
            .Select(attribute => attribute.Value).Distinct().ToArray());

    [Theory]
    [InlineData(CommentRenderMode.EndnoteStyle)]
    [InlineData(CommentRenderMode.Inline)]
    [InlineData(CommentRenderMode.Margin)]
    public void MatchingDurableMetadata_ProvidesTheTimestampInEveryMode(CommentRenderMode mode) =>
        DateIs(Render(Fixture(), mode), Modern);

    [Theory]
    [InlineData(CommentRenderMode.EndnoteStyle, "comment-date")]
    [InlineData(CommentRenderMode.Margin, "comment-margin-date")]
    public void VisibleCaption_UsesTheUtcDayAndIncludesTheYear(CommentRenderMode mode, string dateClass)
    {
        var html = Render(Fixture(), mode);
        var caption = html.Descendants().Single(element =>
            ((string?)element.Attribute("class") ?? "").Split(' ').Contains(dateClass));
        Assert.Equal("Sep 10, 2026", caption.Value);
    }

    [Theory]
    [InlineData("absent")]
    [InlineData("unmatched")]
    [InlineData("invalid")]
    public void UnavailableOrInvalidUtcMetadata_RetainsTheLegacyTimestamp(string kind)
    {
        var bytes = kind switch
        {
            "absent" => Fixture(utc: null),
            "unmatched" => Fixture(matched: false, validate: false),
            _ => Fixture(utc: "not-an-iso-date", validate: false),
        };
        DateIs(Render(bytes, CommentRenderMode.Margin), Legacy);
    }

    [Fact]
    public void DurableMapping_UsesTheLastCommentParagraph() =>
        DateIs(Render(Fixture(leadingParagraph: true), CommentRenderMode.Margin), Modern);

    [Fact]
    public void CommentListing_UsesTheSameEffectiveTimestampAsRendering()
    {
        using var session = new DocxSession(Fixture());
        Assert.Equal(Modern, Assert.Single(session.ListComments()).Date);
    }

    [Theory]
    [InlineData("2026-09-10T06:45:00+02:00", Modern)]
    [InlineData("2026-09-10T04:45:00", Modern)]
    [InlineData("2026-09-10T04:45:00.125Z", "2026-09-10T04:45:00.125Z")]
    [InlineData("2026-09-10", Legacy)]
    public void UtcDates_RespectOffsetsAndPrecisionButRejectDateOnlyStrings(string date, string expected) =>
        DateIs(Render(Fixture(utc: date, validate: false), CommentRenderMode.Margin), expected);

    [Theory]
    [InlineData("duplicate-map")]
    [InlineData("duplicate-date")]
    [InlineData("other-durable-id")]
    public void AmbiguousOrUnmatchedRecords_CannotOverrideTheLegacyDate(string kind)
    {
        var bytes = Fixture();
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var main = doc.MainDocumentPart!;
            var part = kind == "duplicate-map" ? (OpenXmlPart)main.WordprocessingCommentsIdsPart!
                : main.WordCommentsExtensiblePart!;
            var xml = part.GetXDocument();
            if (kind == "other-durable-id")
                xml.Root!.Elements().Single().SetAttributeValue(Cex + "durableId", "00000033");
            else
                xml.Root!.Add(new XElement(xml.Root.Elements().Single()));
            part.PutXDocument(xml);
        }
        DateIs(Render(stream.ToArray(), CommentRenderMode.Margin), Legacy);
    }

    [Fact]
    public void HexadecimalKeys_AreCaseInsensitive()
    {
        var bytes = Fixture();
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var part = doc.MainDocumentPart!.WordprocessingCommentsIdsPart!;
            var xml = part.GetXDocument();
            var entry = xml.Root!.Elements().Single();
            entry.SetAttributeValue(Cid + "paraId", "00aa00aa");
            entry.SetAttributeValue(Cid + "durableId", "02bb02bb");
            part.PutXDocument(xml);
        }
        DateIs(Render(stream.ToArray(), CommentRenderMode.Margin), Modern);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BlockRendering_RetainsTheModernMetadataInStatelessAndCachedPaths(bool cached)
    {
        var bytes = Fixture();
        using var session = new DocxSession(bytes);
        var anchor = session.ListBlocks().Body[0].Id;
        var options = new HtmlConversionOptions { CommentRenderMode = 1, FabricateCssClasses = false };
        for (var pass = 0; pass < 2; pass++)
        {
            var fragment = cached ? HtmlConversionOps.RenderBlockHtml(session, anchor, options)
                : HtmlConversionOps.RenderBlockHtml(bytes, anchor, options);
            DateIs(XElement.Parse(fragment), Modern);
        }
    }
}
