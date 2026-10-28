using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Security;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// An external relationship target that <see cref="System.Uri"/> cannot parse — a multi-recipient
/// <c>mailto:a@x;b@y</c> is the realistic shape — must open in every pipeline, read back as
/// spelled, be accepted by the writers that take what the readers return, and leave every output
/// exactly as the input spelled it (issue #1388). The target is relationship data, not a URI the
/// library resolves.
/// </summary>
public class UnparsedExternalTargetTests
{
    private const string MultiRecipient = "mailto:first@example.com;second@example.com";
    private const string OtherMultiRecipient = "mailto:third@example.com;fourth@example.com";
    private const string MultiRecipientWithQuery = "mailto:first@example.com;second@example.com?subject=Q&body=R";
    private const string QuotedNames = "mailto:\"Team A\" <first@example.com>;second@example.com";
    private const string SingleRecipient = "mailto:solo@example.com";
    // A parseable target that starts with the scheme the library encodes under.
    private const string Prefixed = "x-docxodus-verbatim:mailto:solo@example.com";
    private const string Scheme = "x-docxodus-verbatim:";

    private const string RelsEntry = "word/_rels/document.xml.rels";
    private const string HyperlinkType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    private const string ImageType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    private const string Png1x1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

    private static readonly XNamespace PackageRels = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>A package whose main part is <paramref name="bodyXml"/> plus the given external
    /// relationships, and optionally a 1x1 PNG at relationship <c>rIdPng</c>. The SDK cannot write
    /// a target <see cref="System.Uri"/> rejects, so a parseable stand-in is written first and the
    /// rels entry is then patched to the real target.</summary>
    private static WmlDocument Package(string bodyXml, bool withImage, params (string Id, string Type, string Target)[] externals)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            foreach (var (id, type, _) in externals)
            {
                var standIn = new Uri($"https://stand-in.example.com/{id}");
                if (type == HyperlinkType)
                    main.AddHyperlinkRelationship(standIn, true, id);
                else
                    main.AddExternalRelationship(type, standIn, id);
            }
            if (withImage)
            {
                using var png = new MemoryStream(Convert.FromBase64String(Png1x1));
                main.AddImagePart(ImagePartType.Png, "rIdPng").FeedData(png);
            }
            using var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write));
            writer.Write(
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" " +
                "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">" +
                $"<w:body>{bodyXml}</w:body></w:document>");
        }

        using var patched = new MemoryStream();
        patched.Write(ms.ToArray());
        using (var zip = new ZipArchive(patched, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(RelsEntry)!;
            string xml;
            using (var reader = new StreamReader(entry.Open()))
                xml = reader.ReadToEnd();
            entry.Delete();
            foreach (var (id, _, target) in externals)
                xml = xml.Replace($"https://stand-in.example.com/{id}", SecurityElement.Escape(target));
            using var writer = new StreamWriter(zip.CreateEntry(RelsEntry).Open());
            writer.Write(xml);
        }
        return new WmlDocument("unparsed-target.docx", patched.ToArray());
    }

    /// <summary>Plain paragraphs and one hyperlinked paragraph per non-null target.</summary>
    private static WmlDocument Doc(params (string Text, string? Target)[] paragraphs) => Package(
        string.Concat(paragraphs.Select((p, i) => p.Target is null
            ? $"<w:p><w:r><w:t>{p.Text}</w:t></w:r></w:p>"
            : $"<w:p><w:hyperlink r:id=\"rIdLink{i}\"><w:r><w:t>{p.Text}</w:t></w:r></w:hyperlink></w:p>")),
        withImage: false,
        paragraphs.Select((p, i) => (p.Target, i)).Where(x => x.Target is not null)
            .Select(x => ($"rIdLink{x.i}", HyperlinkType, x.Target!)).ToArray());

    private static string[] ExternalTargets(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var stream = zip.GetEntry(RelsEntry)!.Open();
        return XDocument.Load(stream).Root!.Elements(PackageRels + "Relationship")
            .Where(r => (string?)r.Attribute("TargetMode") == "External")
            .Select(r => (string)r.Attribute("Target")!)
            .ToArray();
    }

    /// <summary>Every external target in the output is one the inputs spelled, and the linked
    /// target survives at least once.</summary>
    private static void AssertTargetsPreserved(byte[] output, string target)
    {
        var targets = ExternalTargets(output);
        Assert.Contains(target, targets);
        Assert.All(targets, t => Assert.Equal(target, t));
    }

    private static string[] Hrefs(string html) => Regex.Matches(html, "href=\"([^\"]*)\"")
        .Select(m => WebUtility.HtmlDecode(m.Groups[1].Value)).ToArray();

    private static string LinkParagraph(DocxSession session) =>
        Regex.Match(session.Project().Markdown, @"\{#([^}]+)\} \[Mail the team\]").Groups[1].Value;

    public static TheoryData<string> Targets => new() { MultiRecipient, MultiRecipientWithQuery, QuotedNames, Prefixed, SingleRecipient };

    [Theory]
    [MemberData(nameof(Targets))]
    public void Compare_keeps_a_link_both_sides_share(string target)
    {
        var left = Doc(("Intro", null), ("Mail the team", target));
        var right = Doc(("Intro", null), ("Mail the team", target), ("An added paragraph.", null));

        AssertTargetsPreserved(DocxCompare.Compare(left, right).DocumentByteArray, target);
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Compare_keeps_the_target_of_an_inserted_link(string target)
    {
        var left = Doc(("Intro", null));
        var right = Doc(("Intro", null), ("Mail the team", target));

        AssertTargetsPreserved(DocxCompare.Compare(left, right).DocumentByteArray, target);
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Compare_keeps_the_target_of_a_deleted_link(string target)
    {
        var left = Doc(("Intro", null), ("Mail the team", target));
        var right = Doc(("Intro", null));

        AssertTargetsPreserved(DocxCompare.Compare(left, right).DocumentByteArray, target);
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Consolidate_keeps_the_target_of_an_inserted_link(string target)
    {
        var original = Doc(("Intro", null));
        var revised = Doc(("Intro", null), ("Mail the team", target));

        var output = DocxCompare.Consolidate(original,
            new[] { new DocxDiffReviewer { Document = revised, Author = "Reviewer A" } });

        AssertTargetsPreserved(output.DocumentByteArray, target);
    }

    [Theory]
    [InlineData(MultiRecipient, false)]
    [InlineData(MultiRecipient, true)]
    [InlineData(MultiRecipientWithQuery, false)]
    [InlineData(QuotedNames, false)]
    [InlineData(Prefixed, false)]
    [InlineData(SingleRecipient, false)]
    [InlineData(SingleRecipient, true)]
    public void Session_save_keeps_the_target(string target, bool persistAnchorIds)
    {
        using var session = new DocxSession(Doc(("Intro", null), ("Mail the team", target)).DocumentByteArray);
        session.Project();

        AssertTargetsPreserved(session.Save(persistAnchorIds), target);
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Undoing_a_link_removal_restores_the_target(string target)
    {
        using var session = new DocxSession(Doc(("Intro", null), ("Mail the team", target)).DocumentByteArray);
        Assert.True(session.RemoveHyperlink(Assert.Single(session.ListHyperlinks()).Id).Success);
        Assert.True(session.Undo());

        Assert.Equal(target, Assert.Single(session.ListHyperlinks()).Target);
        AssertTargetsPreserved(session.Save(false), target);
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void An_unedited_session_reports_no_semantic_changes(string target)
    {
        using var session = new DocxSession(Doc(("Intro", null), ("Mail the team", target)).DocumentByteArray,
            new DocxSessionSettings { CaptureInitialProjection = true });

        Assert.Empty(session.GetSemanticChanges().Changes);
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Readers_report_the_target_as_written(string target)
    {
        var doc = Doc(("Intro", null), ("Mail the team", target));

        using var session = new DocxSession(doc.DocumentByteArray);
        Assert.Equal(target, Assert.Single(session.ListHyperlinks()).Target);
        Assert.Contains($"[Mail the team]({target})", session.Project().Markdown);

        var html = WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings());
        Assert.Contains(target, html.Descendants().Select(e => (string?)e.Attribute("href")));
    }

    [Fact]
    public void Accept_revisions_keeps_the_target()
    {
        var output = RevisionProcessor.AcceptRevisions(Doc(("Intro", null), ("Mail the team", MultiRecipient)));

        AssertTargetsPreserved(output.DocumentByteArray, MultiRecipient);
    }

    [Fact]
    public void A_package_with_only_parseable_targets_is_read_as_the_same_bytes()
    {
        var bytes = Doc(("Intro", null), ("Mail the team", SingleRecipient)).DocumentByteArray;

        Assert.Same(bytes, StrictOoxmlNormalizer.NormalizeForReading(bytes));
    }

    [Fact]
    public void A_relationships_part_laid_out_one_attribute_per_line_is_repaired_where_it_needs_it()
    {
        var doc = Doc(("Intro", null), ("Mail the team", MultiRecipient));
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(RelsEntry)!;
            string xml;
            using (var reader = new StreamReader(entry.Open()))
                xml = reader.ReadToEnd();
            entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry(RelsEntry).Open());
            writer.Write(xml.Replace("<Relationship ", "<Relationship\n    "));
        }

        using var session = new DocxSession(ms.ToArray());
        var saved = session.Save(false);

        AssertTargetsPreserved(saved, MultiRecipient);
        using var reopened = WordprocessingDocument.Open(new MemoryStream(saved), false);
        Assert.NotNull(reopened.MainDocumentPart!.StyleDefinitionsPart);
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Reading_funnel_output_again_changes_nothing(string target)
    {
        var read = StrictOoxmlNormalizer.NormalizeForReading(Doc(("Intro", null), ("Mail the team", target)).DocumentByteArray);

        Assert.Equal(ExternalTargets(read), ExternalTargets(StrictOoxmlNormalizer.NormalizeForReading(read)));
    }

    [Fact]
    public void An_unparseable_target_is_read_parseable_and_written_back_as_spelled()
    {
        var bytes = Doc(("Intro", null), ("Mail the team", MultiRecipientWithQuery)).DocumentByteArray;

        var read = StrictOoxmlNormalizer.NormalizeForReading(bytes);
        var readTarget = Assert.Single(ExternalTargets(read));
        Assert.True(Uri.TryCreate(readTarget, UriKind.RelativeOrAbsolute, out _), readTarget);
        Assert.Equal(new[] { MultiRecipientWithQuery }, ExternalTargets(ZipPackageOutputNormalizer.Normalize(read)));
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Rolling_back_a_transaction_keeps_the_target(string target)
    {
        using var session = new DocxSession(Doc(("Intro", null), ("Mail the team", target)).DocumentByteArray);
        using (var transaction = session.BeginTransaction())
        {
            Assert.True(session.RemoveHyperlink(Assert.Single(session.ListHyperlinks()).Id).Success);
            transaction.Rollback();
        }

        Assert.Equal(target, Assert.Single(session.ListHyperlinks()).Target);
        AssertTargetsPreserved(session.Save(false), target);
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Undoing_a_committed_transaction_keeps_the_target(string target)
    {
        using var session = new DocxSession(Doc(("Intro", null), ("Mail the team", target)).DocumentByteArray);
        using (var transaction = session.BeginTransaction())
        {
            Assert.True(session.RemoveHyperlink(Assert.Single(session.ListHyperlinks()).Id).Success);
            transaction.Commit();
        }
        Assert.True(session.Undo());

        AssertTargetsPreserved(session.Save(false), target);
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Undoing_a_batch_keeps_the_target(string target)
    {
        using var session = new DocxSession(Doc(("Intro", null), ("Mail the team", target)).DocumentByteArray);
        var link = Assert.Single(session.ListHyperlinks()).Id;
        Assert.True(session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_links", "remove_hyperlink", s => s.RemoveHyperlink(link)),
        }).Success);
        Assert.True(session.Undo());

        AssertTargetsPreserved(session.Save(false), target);
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Writers_accept_the_target_readers_return(string target)
    {
        using var session = new DocxSession(Doc(("Intro", null), ("Mail the team", target)).DocumentByteArray);
        var link = Assert.Single(session.ListHyperlinks());

        var update = session.UpdateHyperlink(link.Id, HyperlinkTarget.External(link.Target!));
        Assert.True(update.Success, update.Error?.Message);
        var echoed = Regex.Match(session.Project().Markdown, @"\[Mail the team\]\([^\n]*\)").Value;
        var replace = session.ReplaceText(LinkParagraph(session), echoed);
        Assert.True(replace.Success, replace.Error?.Message);
        var intro = Regex.Match(session.Project().Markdown, @"\{#([^}]+)\} Intro").Groups[1].Value;
        var add = session.AddHyperlink(intro, new CharSpan(0, 5), HyperlinkTarget.External(target));
        Assert.True(add.Success, add.Error?.Message);

        Assert.All(session.ListHyperlinks(), l => Assert.Equal(target, l.Target));
        AssertTargetsPreserved(session.Save(false), target);
    }

    [Fact]
    public void Semantic_changes_report_targets_as_spelled()
    {
        var before = Doc(("Intro", null), ("Mail the team", MultiRecipient)).DocumentByteArray;
        var after = Doc(("Intro", null), ("Mail the team", OtherMultiRecipient)).DocumentByteArray;

        var json = Docxodus.Verification.SemanticDiff.Compare(before, after).ToCanonicalJson();

        Assert.DoesNotContain(Scheme, json);
        Assert.Contains("second@example.com", json);
        Assert.Contains("fourth@example.com", json);
    }

    [Fact]
    public void A_session_reports_an_edited_target_as_spelled()
    {
        using var session = new DocxSession(Doc(("Intro", null), ("Mail the team", MultiRecipient)).DocumentByteArray,
            new DocxSessionSettings { CaptureInitialProjection = true });
        Assert.True(session.UpdateHyperlink(Assert.Single(session.ListHyperlinks()).Id,
            HyperlinkTarget.External(OtherMultiRecipient)).Success);

        var changes = session.GetSemanticChangesJson();
        Assert.DoesNotContain(Scheme, changes);
        Assert.Contains("fourth@example.com", changes);
        Assert.DoesNotContain(Scheme, session.VerifyDeliverableJson());
    }

    [Fact]
    public void A_delivery_receipt_reports_an_edited_target_as_spelled()
    {
        using var session = new DocxSession(Doc(("Intro", null), ("Mail the team", MultiRecipient)).DocumentByteArray,
            new DocxSessionSettings { CaptureInitialProjection = true, CaptureDeliveryEvidence = true });
        Assert.True(session.UpdateHyperlink(Assert.Single(session.ListHyperlinks()).Id,
            HyperlinkTarget.External(OtherMultiRecipient)).Success);

        var bundle = session.BuildDeliveryReceipt();

        var text = bundle.ArtifactBytes.Where(a => !a.Key.EndsWith(".docx", StringComparison.Ordinal))
            .Select(a => System.Text.Encoding.UTF8.GetString(a.Value)).Append(System.Text.Encoding.UTF8.GetString(bundle.ManifestBytes))
            .ToList();
        Assert.All(text, t => Assert.DoesNotContain(Scheme, t));
        Assert.Contains(text, t => t.Contains("fourth@example.com", StringComparison.Ordinal));
        Assert.All(bundle.ArtifactBytes.Where(a => a.Key.EndsWith(".docx", StringComparison.Ordinal)),
            a => Assert.DoesNotContain(Scheme, string.Join("|", ExternalTargets(a.Value))));
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Html_from_bytes_links_the_target_as_spelled(string target)
    {
        var bytes = Doc(("Intro", null), ("Mail the team", target)).DocumentByteArray;
        using var session = new DocxSession(bytes);
        var options = new Docxodus.Internal.HtmlConversionOptions();

        Assert.Contains(target, Hrefs(Docxodus.Internal.HtmlConversionOps.ConvertToHtml(bytes, options)));
        Assert.Contains(target, Hrefs(Docxodus.Internal.HtmlConversionOps.RenderBlockHtml(bytes, LinkParagraph(session), options)));
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Ir_and_grep_report_the_target_as_spelled(string target)
    {
        var doc = Doc(("Intro", null), ("Mail the team", target));

        using var ir = JsonDocument.Parse(IrDiagnosticJson.Write(IrReader.Read(doc)));
        Assert.Contains(target, Strings(ir.RootElement, "target"));
        using var session = new DocxSession(doc.DocumentByteArray);
        Assert.Equal(target, Assert.Single(session.Grep("Mail")).Fragments.Single().Formatting.HyperlinkUrl);
    }

    private static IEnumerable<string> Strings(JsonElement element, string property) => element.ValueKind switch
    {
        JsonValueKind.Object => element.EnumerateObject().SelectMany(p =>
            p.Name == property && p.Value.ValueKind == JsonValueKind.String
                ? new[] { p.Value.GetString()! }
                : Strings(p.Value, property)),
        JsonValueKind.Array => element.EnumerateArray().SelectMany(e => Strings(e, property)),
        _ => Enumerable.Empty<string>(),
    };

    private const string LinkedPicture =
        "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"914400\"/>" +
        "<wp:docPr id=\"1\" name=\"Picture 1\">{0}</wp:docPr>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic>" +
        "<pic:nvPicPr><pic:cNvPr id=\"1\" name=\"Picture 1\"/><pic:cNvPicPr/></pic:nvPicPr>" +
        "<pic:blipFill><a:blip {1}/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
        "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\"rect\"/></pic:spPr>" +
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    [Theory]
    [InlineData("http://exa mple.com/picture.png")]
    [InlineData("http://example.com/picture.png")]
    public void Linked_images_report_the_target_as_spelled(string target)
    {
        var drawing = string.Format(LinkedPicture, "", "r:link=\"rIdLinked\"");
        var vml = "<w:p><w:r><w:pict><v:shape id=\"s1\" style=\"width:10pt;height:10pt\">" +
            "<v:imagedata r:id=\"rIdVml\" o:title=\"\"/></v:shape></w:pict></w:r></w:p>";
        var doc = Package(drawing + vml, withImage: false, ("rIdLinked", ImageType, target), ("rIdVml", ImageType, target));

        using var session = new DocxSession(doc.DocumentByteArray);

        Assert.Equal(new[] { target, target }, session.ListImages().Select(i => i.LinkedTarget));
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void Html_links_a_picture_to_the_target_as_spelled(string target)
    {
        var drawing = string.Format(LinkedPicture, "<a:hlinkClick r:id=\"rIdPicLink\"/>", "r:embed=\"rIdPng\"");
        var doc = Package(drawing, withImage: true, ("rIdPicLink", HyperlinkType, target));

        var html = WmlToHtmlConverter.ConvertToHtml(doc,
            new WmlToHtmlConverterSettings { ImageHandler = _ => new XElement("img") });

        Assert.Contains(target, html.Descendants().Select(e => (string?)e.Attribute("href")));
    }

    [Fact]
    public void Breaking_the_template_link_keeps_the_target()
    {
        var output = WordprocessingMLUtil.BreakLinkToTemplate(Doc(("Intro", null), ("Mail the team", MultiRecipient)));

        AssertTargetsPreserved(output.DocumentByteArray, MultiRecipient);
    }

    /// <summary>Replaces one entry's bytes.</summary>
    private static byte[] WithEntry(byte[] package, string name, Func<string, byte[]> rewrite)
    {
        using var ms = new MemoryStream();
        ms.Write(package);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(name)!;
            string xml;
            using (var reader = new StreamReader(entry.Open()))
                xml = reader.ReadToEnd();
            entry.Delete();
            using var output = zip.CreateEntry(name).Open();
            output.Write(rewrite(xml));
        }
        return ms.ToArray();
    }

    // Built in code rather than as theory data, which would replace the lone surrogate.
    public static TheoryData<int> CharactersXmlCannotCarry => new() { 0x01, 0x00, 0xFFFE, 0xD800 };

    [Theory]
    [MemberData(nameof(CharactersXmlCannotCarry))]
    public void Writers_refuse_a_target_xml_cannot_carry(int character)
    {
        var target = "mailto:first@example.com;second@example.com" + (char)character;
        using var session = new DocxSession(Doc(("Intro", null), ("Mail the team", MultiRecipient)).DocumentByteArray);
        var intro = Regex.Match(session.Project().Markdown, @"\{#([^}]+)\} Intro").Groups[1].Value;

        var add = session.AddHyperlink(intro, new CharSpan(0, 5), HyperlinkTarget.External(target));
        var update = session.UpdateHyperlink(Assert.Single(session.ListHyperlinks()).Id, HyperlinkTarget.External(target));

        var replace = session.ReplaceText(intro, $"[Intro]({target})");

        Assert.Equal(EditErrorCode.InvalidHyperlinkTarget, add.Error?.Code);
        Assert.Equal(EditErrorCode.InvalidHyperlinkTarget, update.Error?.Code);
        Assert.False(replace.Success); // the markdown reader already refuses the character
        using var reopened = new DocxSession(session.Save(false));
        Assert.Equal(MultiRecipient, Assert.Single(reopened.ListHyperlinks()).Target);
    }

    [Theory]
    [InlineData("mailto:first@example.com;second@example.com\n")]
    [InlineData("mailto:first@example.com;second@example.com\t")]
    [InlineData("mailto:first@example.com;\rsecond@example.com")]
    public void A_target_with_a_line_break_or_tab_round_trips(string target)
    {
        using var written = new DocxSession(Doc(("Intro", null), ("Mail the team", MultiRecipient)).DocumentByteArray);
        Assert.True(written.UpdateHyperlink(Assert.Single(written.ListHyperlinks()).Id, HyperlinkTarget.External(target)).Success);
        var read = WithEntry(Doc(("Intro", null), ("Mail the team", "STAND-IN")).DocumentByteArray, RelsEntry,
            xml => System.Text.Encoding.UTF8.GetBytes(xml.Replace("STAND-IN", new XAttribute("x", target).ToString()[3..^1])));
        using var reading = new DocxSession(read);

        Assert.Equal(new[] { target }, ExternalTargets(written.Save(false)));
        Assert.Equal(new[] { target }, ExternalTargets(reading.Save(false)));
    }

    [Fact]
    public void A_utf16_relationships_part_is_written_back_as_spelled()
    {
        var package = WithEntry(Doc(("Intro", null), ("Mail the team", MultiRecipient)).DocumentByteArray, RelsEntry, xml =>
        {
            xml = "<?xml version=\"1.0\" encoding=\"utf-16\"?>" + Regex.Replace(xml, "^<\\?xml[^>]*\\?>", "");
            return System.Text.Encoding.Unicode.GetPreamble().Concat(System.Text.Encoding.Unicode.GetBytes(xml)).ToArray();
        });
        using var session = new DocxSession(package);

        AssertTargetsPreserved(session.Save(false), MultiRecipient);
    }

    [Fact]
    public void A_target_of_just_the_scheme_round_trips()
    {
        using var session = new DocxSession(Doc(("Intro", null), ("Mail the team", Scheme)).DocumentByteArray);

        Assert.Equal(Scheme, Assert.Single(session.ListHyperlinks()).Target);
        AssertTargetsPreserved(session.Save(false), Scheme);
    }

    [Fact]
    public void Extracting_the_glossary_keeps_a_linked_image_target()
    {
        const string target = "http://exa mple.com/picture.png";
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new DocumentFormat.OpenXml.Wordprocessing.Body(
                new DocumentFormat.OpenXml.Wordprocessing.Paragraph()));
            var glossary = main.AddNewPart<GlossaryDocumentPart>();
            glossary.AddExternalRelationship(ImageType, new Uri("https://stand-in.example.com/glossary"), "rIdGlossaryImage");
            using var writer = new StreamWriter(glossary.GetStream(FileMode.Create, FileAccess.Write));
            writer.Write("<w:glossaryDocument xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:docParts><w:docPart><w:docPartPr>" +
                "<w:name w:val=\"Entry\"/><w:category><w:name w:val=\"General\"/><w:gallery w:val=\"placeholder\"/></w:category>" +
                "<w:guid w:val=\"{00000000-0000-0000-0000-000000000002}\"/></w:docPartPr><w:docPartBody><w:p><w:r><w:pict>" +
                "<v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:imagedata r:id=\"rIdGlossaryImage\"/></v:shape>" +
                "</w:pict></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>");
        }
        var package = WithEntry(ms.ToArray(), "word/glossary/_rels/document.xml.rels", xml =>
            System.Text.Encoding.UTF8.GetBytes(xml.Replace("https://stand-in.example.com/glossary", target)));

        var extracted = DocumentBuilder.ExtractGlossaryDocument(new WmlDocument("glossary.docx", package))!;

        Assert.Equal(new[] { target }, ExternalTargets(extracted.DocumentByteArray));
    }

    [Fact]
    public void A_package_level_unparseable_target_opens_as_a_document()
    {
        var package = WithEntry(Doc(("Intro", null)).DocumentByteArray, "_rels/.rels", xml =>
            System.Text.Encoding.UTF8.GetBytes(xml.Replace("</Relationships>",
                "<Relationship Id=\"rIdPackageLink\" Type=\"http://example.com/relationships/contact\" " +
                $"Target=\"{MultiRecipient}\" TargetMode=\"External\"/></Relationships>")));

        var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("package-link.docx", package));

        using var zip = new ZipArchive(new MemoryStream(accepted.DocumentByteArray), ZipArchiveMode.Read);
        using var stream = zip.GetEntry("_rels/.rels")!.Open();
        Assert.Contains(MultiRecipient, XDocument.Load(stream).Root!.Elements(PackageRels + "Relationship")
            .Select(r => (string?)r.Attribute("Target")));
    }
}
