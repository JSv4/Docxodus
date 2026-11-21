using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using static Docxodus.Tests.Ir.Diff.HeaderFooterFixtures;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// A header or footer reference written without <c>w:type</c>, which the schema requires, is read as the section's
/// default story. When that story changes, the comparison keeps the reference as written and binds it to the
/// redlined part, rather than adding a second, typed default reference beside it (issue #1803).
/// </summary>
public class UntypedStoryReferenceTests
{
    private static readonly XNamespace W = Wn;
    private static readonly XNamespace R = Rns;

    /// <summary>The document with <paramref name="edit"/> applied to its main part's body.</summary>
    private static WmlDocument EditBody(WmlDocument doc, System.Action<XElement> edit)
    {
        using var stream = new MemoryStream();
        stream.Write(doc.DocumentByteArray);
        using (var package = WordprocessingDocument.Open(stream, true))
        {
            var main = package.MainDocumentPart!;
            edit(main.GetXDocument().Root!.Element(W + "body")!);
            main.PutXDocument();
        }
        return new WmlDocument("edited.docx", stream.ToArray());
    }

    private static IEnumerable<XElement> StoryReferences(XElement container) => container.Descendants()
        .Where(e => e.Name == W + "headerReference" || e.Name == W + "footerReference");

    /// <summary>The document with <c>w:type="default"</c> removed from every header and footer reference.</summary>
    private static WmlDocument Untyped(WmlDocument doc) => EditBody(doc, body =>
        StoryReferences(body).Attributes(W + "type").Where(a => a.Value == "default").Remove());

    private static WmlDocument OneSection(bool header, string story) => Build(
        new[] { new Section(new[] { "Body" },
            Headers: header ? new[] { ("default", "rIdH1") } : null,
            Footers: header ? null : new[] { ("default", "rIdF1") }) },
        headerParts: header ? new Dictionary<string, string[]> { ["rIdH1"] = new[] { story } } : null,
        footerParts: header ? null : new Dictionary<string, string[]> { ["rIdF1"] = new[] { story } });

    private static XElement Body(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!;
    }

    /// <summary>The final section's header (or footer) references, as written.</summary>
    private static List<XElement> References(WmlDocument doc, bool header) =>
        Body(doc).Element(W + "sectPr")!.Elements(W + (header ? "headerReference" : "footerReference")).ToList();

    /// <summary>Each section's references by element, the story they select (an absent type is the default) and part.</summary>
    private static string[] Sections(WmlDocument doc) => Body(doc).Descendants(W + "sectPr")
        .Where(s => s.Parent!.Name != W + "sectPrChange")
        .Select(s => string.Join(",", s.Elements().Where(e => e.Name == W + "headerReference" || e.Name == W + "footerReference")
            .Select(e => $"{e.Name.LocalName}:{(string?)e.Attribute(W + "type") ?? "default"}:{(string?)e.Attribute(R + "id")}")))
        .ToArray();

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void A_changed_story_keeps_one_reference_and_gives_it_its_type(bool header)
    {
        // The reference the comparison binds to the redlined part is written with its type, so a reader that skips
        // an untyped reference, as LibreOffice does, still finds the redlined story.
        var left = Untyped(OneSection(header, "Old story"));
        var right = Untyped(OneSection(header, "New story"));

        var redline = DocxCompare.Compare(left, right);
        var reference = Assert.Single(References(redline, header));
        Assert.Equal("default", (string?)reference.Attribute(W + "type"));
        foreach (var (resolved, text) in new[]
                 {
                     (RevisionProcessor.RejectRevisions(redline), "Old story"),
                     (RevisionProcessor.AcceptRevisions(redline), "New story"),
                 })
        {
            Assert.Single(References(resolved, header));
            Assert.Equal(text, ReferencedStoryText(resolved, header, 0, "default"));
        }
    }

    [Theory]
    [InlineData(true, "Old story")]
    [InlineData(false, "Old story")]
    [InlineData(true, "New story")]
    public void A_typed_reference_or_an_unchanged_story_keeps_its_reference_as_written(bool typed, string revisedStory)
    {
        // Controls: a typed source reference stays one typed reference; an unchanged untyped story passes through.
        var left = OneSection(true, "Old story");
        var right = OneSection(true, revisedStory);
        if (!typed) (left, right) = (Untyped(left), Untyped(right));

        var reference = Assert.Single(References(DocxCompare.Compare(left, right), true));
        Assert.Equal(typed ? "default" : null, (string?)reference.Attribute(W + "type"));
    }

    /// <summary>
    /// The same comparison written with typed and with untyped references gives each section the same references,
    /// in the redline and after accepting or rejecting it. <c>story</c> changes a one-section document's story;
    /// <c>boundary</c> adds a section break after retained text, so each section's story is bound anew;
    /// <c>collapsed</c> drops the empty paragraph whose section break held the original's only footer reference,
    /// so that reference is attached again to the surviving section.
    /// </summary>
    [Theory]
    [InlineData("story", true)]
    [InlineData("story", false)]
    [InlineData("boundary", true)]
    [InlineData("boundary", false)]
    [InlineData("collapsed", false)]
    public void An_untyped_document_gets_the_references_its_typed_twin_gets(string scenario, bool header)
    {
        Section Story(string[] paras, string id) => new(paras,
            Headers: header ? new[] { ("default", id) } : null, Footers: header ? null : new[] { ("default", id) });
        WmlDocument Document(Section[] sections, Dictionary<string, string[]> stories) =>
            Build(sections, header ? stories : null, header ? null : stories);
        const string retained = "The maintenance summary stays in this section.";
        const string tail = "Arrange the welcome badges by arrival time.";
        var (left, right) = scenario switch
        {
            "story" => (OneSection(header, "Old story"), OneSection(header, "New story")),
            "boundary" => (Document(new[] { Story(new[] { retained, tail }, "a") }, new() { ["a"] = new[] { "MAINTENANCE COPY" } }),
                Document(new[] { Story(new[] { retained.Replace("maintenance", "inspection") }, "a"), Story(new[] { tail }, "empty") },
                    new() { ["a"] = new[] { "MAINTENANCE COPY" }, ["empty"] = new[] { "" } })),
            _ => (EditBody(Document(new[] { Story(new[] { "" }, "a"), new Section(new[] { "Ancient body words here." }) },
                        new() { ["a"] = new[] { "Old left footer text" } }),
                    body => body.Elements(W + "p").First().Elements(W + "r").Remove()),
                Build(new[] { new Section(new[] { "", "Completely fresh replacement body." }) })),
        };

        var typed = DocxCompare.Compare(left, right);
        var untyped = DocxCompare.Compare(Untyped(left), Untyped(right));
        Assert.Equal(Sections(typed), Sections(untyped));
        Assert.Equal(Sections(RevisionProcessor.RejectRevisions(typed)), Sections(RevisionProcessor.RejectRevisions(untyped)));
        Assert.Equal(Sections(RevisionProcessor.AcceptRevisions(typed)), Sections(RevisionProcessor.AcceptRevisions(untyped)));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void A_script_without_reference_bindings_keeps_the_single_untyped_reference(bool header)
    {
        // An edit script written before story diffs carried reference bindings re-attaches each changed story where
        // no reference of its kind is reachable; an untyped reference is one.
        var left = Untyped(OneSection(header, "Old story"));
        var right = Untyped(OneSection(header, "New story"));
        var script = Docxodus.Ir.Diff.IrEditScriptJson.Read(DocxDiff.GetEditScriptJson(left, right));
        script = script with { HeaderFooterOps = Docxodus.Ir.IrNodeList.From(
            script.HeaderFooterOps!.Select(story => story with { ReferenceBindings = null })) };

        var redline = Docxodus.Ir.Diff.IrMarkupRenderer.Render(script, left, right, new Docxodus.Ir.Diff.IrDiffSettings());
        Assert.Null(Assert.Single(References(redline, header)).Attribute(W + "type"));
        Assert.Equal("Old story", ReferencedStoryText(RevisionProcessor.RejectRevisions(redline), header, 0, "default"));
    }

    [Fact]
    public void An_untyped_reference_to_another_part_beside_the_typed_one_is_kept()
    {
        // The section names two default headers, the untyped one first; the typed one's story changes. Binding it
        // leaves the other part's reference alone, as before, rather than orphaning that part.
        static WmlDocument Pair(string story) => EditBody(Build(
            new[] { new Section(new[] { "Body" }, Headers: new[] { ("default", "rIdB") }) },
            headerParts: new Dictionary<string, string[]> { ["rIdB"] = new[] { story }, ["rIdA"] = new[] { "Alpha story" } }),
            body => body.Element(W + "sectPr")!.AddFirst(new XElement(W + "headerReference", new XAttribute(R + "id", "rIdA"))));
        var redline = DocxCompare.Compare(Pair("Old story"), Pair("New story"));

        var references = References(redline, true);
        Assert.Equal(2, references.Count);
        Assert.Contains(references, reference => (string?)reference.Attribute(R + "id") == "rIdA" && reference.Attribute(W + "type") is null);
    }

    [Fact]
    public void A_changed_story_referenced_with_and_without_its_type_comes_out_with_one_reference()
    {
        // The redline earlier versions wrote for an untyped reference: a typed copy beside it, naming the same part.
        // Comparing two such documents binds the changed story to one reference.
        static WmlDocument BothSpellings(string story) => EditBody(OneSection(true, story), body =>
        {
            foreach (var reference in StoryReferences(body).ToList())
                reference.AddAfterSelf(new XElement(reference.Name, reference.Attribute(R + "id")));
        });
        var redline = DocxCompare.Compare(BothSpellings("Old story"), BothSpellings("New story"));

        Assert.Equal("default", (string?)Assert.Single(References(redline, true)).Attribute(W + "type"));
        Assert.Equal("Old story", ReferencedStoryText(RevisionProcessor.RejectRevisions(redline), true, 0, "default"));
    }
}
