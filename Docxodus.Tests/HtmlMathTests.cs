using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlMathTests
{
    private const string RunX = "<m:r><m:t>x</m:t></m:r>";
    private const string RunTwo = "<m:r><m:t>2</m:t></m:r>";
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    [Theory]
    [InlineData("fraction", false)]
    [InlineData("fraction", true)]
    [InlineData("root", false)]
    [InlineData("root", true)]
    [InlineData("indexed-root", false)]
    [InlineData("indexed-root", true)]
    public void Equations_preserve_operands_and_structure_as_selectable_math(string shape, bool display)
    {
        var formula = shape == "fraction"
            ? "<m:f><m:num>" + RunX + "</m:num><m:den>" + RunTwo + "</m:den></m:f>"
            : "<m:rad><m:radPr><m:degHide m:val=\"" + (shape == "root" ? "1" : "0")
                + "\"/></m:radPr><m:deg>" + (shape == "root" ? "" : RunTwo) + "</m:deg><m:e>" + RunX + "</m:e></m:rad>";
        var html = Convert(formula, display);
        var math = Assert.Single(html.Descendants(MathMl + "math"));
        Assert.Equal(display ? "block" : "inline", (string?)math.Attribute("display"));
        var expression = Assert.Single(math.Descendants(MathMl + (shape == "fraction" ? "mfrac" : shape == "root" ? "msqrt" : "mroot")));
        Assert.Equal(shape == "root" ? "x" : "x2", expression.Value);
        Assert.DoesNotContain(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
        Assert.DoesNotContain("\u00a0", math.Ancestors().First(e => e.Name.LocalName == "p").Value);
    }

    [Theory]
    [InlineData("x\u0301")]
    [InlineData("x y")]
    [InlineData("x\u200dy")]
    public void Token_forms_without_a_faithful_projection_keep_the_diagnostic(string text)
    {
        var html = Convert("<m:r><m:t xml:space=\"preserve\">" + text + "</m:t></m:r>", false);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
        Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
    }

    [Theory]
    [InlineData("<m:smallFrac m:val=\"1\"/>")]
    [InlineData("<m:defJc m:val=\"right\"/>")]
    public void Unsupported_document_math_properties_keep_the_diagnostic(string property)
    {
        var html = Convert(RunX, true, property);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
        Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
    }

    [Fact]
    public void The_document_math_font_is_preserved_for_the_shared_font_resolver()
    {
        var math = Assert.Single(Convert(RunX, false, "<m:mathFont m:val=\"Noto Sans Math\"/>").Descendants(MathMl + "math"));
        Assert.Contains("Noto Sans Math", (string?)math.Attribute("style"));
    }

    [Theory]
    [InlineData(24, false, false, "12pt")]
    [InlineData(48, false, false, "24pt")]
    [InlineData(24, true, false, "12pt")]
    [InlineData(24, true, true, "18pt")]
    public void Equation_size_comes_from_run_styles_not_large_neighbors_or_the_paragraph_mark(
        int defaultSize, bool mark, bool style, string expected)
    {
        var paragraphStart = "<w:pPr>" + (style ? "<w:pStyle w:val=\"MathBody\"/>" : "")
            + (mark ? "<w:rPr><w:sz w:val=\"60\"/></w:rPr>" : "") + "</w:pPr>"
            + "<w:r><w:rPr><w:sz w:val=\"48\"/></w:rPr><w:t>Large neighbor</w:t></w:r>";
        var styles = "<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"" + defaultSize
            + "\"/></w:rPr></w:rPrDefault></w:docDefaults>"
            + (style ? "<w:style w:type=\"paragraph\" w:styleId=\"MathBody\"><w:name w:val=\"Math Body\"/><w:rPr><w:sz w:val=\"36\"/></w:rPr></w:style>" : "");
        var math = Assert.Single(Convert(RunX, false, paragraphStart: paragraphStart, styles: styles).Descendants(MathMl + "math"));
        Assert.Contains("font-size: " + expected + ";", (string?)math.Attribute("style"));
    }

    [Fact]
    public void Literal_operators_do_not_become_scalable_delimiters_or_large_operators()
    {
        var math = Assert.Single(Convert("<m:r><m:t>(x)+∑</m:t></m:r>", true).Descendants(MathMl + "math"));
        Assert.All(math.Descendants(MathMl + "mo"), token =>
        {
            Assert.Equal("false", (string?)token.Attribute("stretchy"));
            Assert.Equal("false", (string?)token.Attribute("largeop"));
        });
    }

    [Fact]
    public void An_empty_equation_keeps_the_empty_paragraph_line()
    {
        var html = Convert("", true);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
        Assert.Contains("\u00a0", html.Value);
    }

    [Fact]
    public void A_decimal_denominator_is_one_number_without_operator_spacing()
    {
        var formula = "<m:f><m:num>" + RunX + "</m:num><m:den><m:r><m:t>31.4</m:t></m:r></m:den></m:f>";
        var math = Assert.Single(Convert(formula, false).Descendants(MathMl + "math"));
        Assert.Equal("31.4", Assert.Single(math.Descendants(MathMl + "mn")).Value);
        Assert.Empty(math.Descendants(MathMl + "mo"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Paginated_source_identity_does_not_turn_supported_math_into_a_placeholder(bool root)
    {
        var formula = root ? "<m:rad><m:radPr><m:degHide/></m:radPr><m:deg/><m:e>" + RunX + "</m:e></m:rad>"
            : "<m:f><m:num>" + RunX + "</m:num><m:den>" + RunTwo + "</m:den></m:f>";
        Assert.Single(Convert(formula, false, paginated: true).Descendants(MathMl + (root ? "msqrt" : "mfrac")));
    }

    [Fact]
    public void Paginated_source_identity_keeps_Word_formatted_structures_projectable()
    {
        const string fonts = "<w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr>";
        var formula = "<m:sSup><m:sSupPr><m:ctrlPr>" + fonts + "</m:ctrlPr></m:sSupPr><m:e><m:r>" + fonts + "<m:t>x</m:t></m:r></m:e>"
            + "<m:sup><m:r>" + fonts + "<m:t>2</m:t></m:r></m:sup></m:sSup>";
        Assert.Single(Convert(formula, false, "<m:mathFont m:val=\"Cambria Math\"/>", paginated: true).Descendants(MathMl + "msup"));
    }

    [Fact]
    public void Supplementary_math_letters_keep_their_code_point()
    {
        var math = Assert.Single(Convert("<m:r><m:t>𝑥</m:t></m:r>", false).Descendants(MathMl + "math"));
        Assert.Equal("𝑥", Assert.Single(math.Descendants(MathMl + "mi")).Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Nested_fractions_keep_a_diagnostic_for_the_unmodeled_display_size_rule(bool display)
    {
        var fraction = "<m:f><m:num>" + RunX + "</m:num><m:den>" + RunTwo + "</m:den></m:f>";
        var html = Convert("<m:f><m:num>" + RunX + "</m:num><m:den>" + fraction + "</m:den></m:f>", display);
        if (display)
        {
            Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
            Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
        }
        else
            Assert.Equal(2, html.Descendants(MathMl + "mfrac").Count());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Unsupported_equation_structure_retains_its_diagnostic(bool display)
    {
        var html = Convert("<m:phant><m:e>" + RunX + "</m:e></m:phant>", display);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
        Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
    }

    // Formulas below write a math run as {text}; Shape() prints MathML as name(child,...) with
    // plain mrow grouping flattened to space-separated tokens.
    [Theory]
    [InlineData("<m:sSup><m:e>{x}</m:e><m:sup>{2}</m:sup></m:sSup>", "msup(x,2)")]
    [InlineData("<m:sSub><m:e>{x}</m:e><m:sub>{i}</m:sub></m:sSub>", "msub(x,i)")]
    [InlineData("<m:sSubSup><m:e>{x}</m:e><m:sub>{i}</m:sub><m:sup>{2}</m:sup></m:sSubSup>", "msubsup(x,i,2)")]
    [InlineData("<m:sPre><m:sub>{i}</m:sub><m:sup>{2}</m:sup><m:e>{x}</m:e></m:sPre>", "mmultiscripts(x,mprescripts(),i,2)")]
    [InlineData("<m:nary><m:naryPr><m:chr m:val=\"∑\"/></m:naryPr><m:sub>{i=1}</m:sub><m:sup>{n}</m:sup><m:e>{x}</m:e></m:nary>", "munderover(∑,i = 1,n) x")]
    [InlineData("<m:nary><m:sub>{0}</m:sub><m:sup>{1}</m:sup><m:e>{x}</m:e></m:nary>", "msubsup(∫,0,1) x")]
    [InlineData("<m:nary><m:naryPr><m:chr m:val=\"∮\"/><m:subHide/><m:supHide/></m:naryPr><m:sub/><m:sup/><m:e>{x}</m:e></m:nary>", "∮ x")]
    [InlineData("<m:nary><m:naryPr><m:chr m:val=\"∏\"/><m:limLoc m:val=\"subSup\"/><m:supHide m:val=\"1\"/></m:naryPr><m:sub>{k}</m:sub><m:sup/><m:e>{x}</m:e></m:nary>", "msub(∏,k) x")]
    [InlineData("<m:d><m:e>{x}</m:e><m:e>{y}</m:e></m:d>", "( x | y )")]
    [InlineData("<m:d><m:dPr><m:begChr m:val=\"[\"/><m:sepChr m:val=\",\"/><m:endChr m:val=\"\"/></m:dPr><m:e>{x}</m:e><m:e>{y}</m:e></m:d>", "[ x , y")]
    [InlineData("<m:m><m:mr><m:e>{a}</m:e><m:e>{b}</m:e></m:mr><m:mr><m:e>{c}</m:e><m:e>{d}</m:e></m:mr></m:m>", "mtable(mtr(mtd(a),mtd(b)),mtr(mtd(c),mtd(d)))")]
    [InlineData("<m:func><m:fName><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>sin</m:t></m:r></m:fName><m:e>{x}</m:e></m:func>", "sin mspace() x")]
    [InlineData("<m:acc><m:e>{x}</m:e></m:acc>", "mover(x,ˆ)")]
    [InlineData("<m:acc><m:accPr><m:chr m:val=\"\u20D7\"/></m:accPr><m:e>{v}</m:e></m:acc>", "mover(v,→)")]
    [InlineData("<m:limLow><m:e><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>lim</m:t></m:r></m:e><m:lim>{n→∞}</m:lim></m:limLow>", "munder(lim,n → ∞)")]
    [InlineData("<m:limUpp><m:e>{x}</m:e><m:lim>{y}</m:lim></m:limUpp>", "mover(x,y)")]
    [InlineData("<m:groupChr><m:e>{x}</m:e></m:groupChr>", "munder(x,⏟)")]
    [InlineData("<m:groupChr><m:groupChrPr><m:chr m:val=\"⏞\"/><m:pos m:val=\"top\"/><m:vertJc m:val=\"bot\"/></m:groupChrPr><m:e>{x}</m:e></m:groupChr>", "mover(x,⏞)")]
    [InlineData("<m:box><m:boxPr><m:noBreak m:val=\"0\"/></m:boxPr><m:e>{x}</m:e></m:box>", "x")]
    [InlineData("<m:borderBox><m:e>{x}</m:e></m:borderBox>", "mrow(x)")]
    [InlineData("<m:f><m:fPr><m:type m:val=\"noBar\"/></m:fPr><m:num>{x}</m:num><m:den>{2}</m:den></m:f>", "mfrac(x,2)")]
    [InlineData("<m:f><m:fPr><m:type m:val=\"lin\"/></m:fPr><m:num>{x}</m:num><m:den>{2}</m:den></m:f>", "x / 2")]
    [InlineData("<m:eqArr><m:e>{x&amp;=1}</m:e><m:e>{y&amp;=2}</m:e></m:eqArr>", "mtable(mtr(mtd(x),mtd(= 1)),mtr(mtd(y),mtd(= 2)))")]
    [InlineData("<m:eqArr><m:e>{x}</m:e><m:e>{y}</m:e></m:eqArr>", "mtable(mtr(mtd(x)),mtr(mtd(y)))")]
    public void Equation_structures_project_to_their_MathML_shape(string formula, string expected)
    {
        Assert.Equal(expected, Shape(ProjectExpression(formula)));
    }

    [Fact]
    public void Nary_operators_are_large_and_grow_only_when_asked()
    {
        var expression = ProjectExpression("<m:nary><m:naryPr><m:chr m:val=\"∑\"/><m:grow/></m:naryPr><m:sub/><m:sup/><m:e>{x}</m:e></m:nary>");
        var op = expression.Descendants(MathMl + "mo").First();
        Assert.Equal(("true", "true", "true"), ((string?)op.Attribute("largeop"), (string?)op.Attribute("movablelimits"), (string?)op.Attribute("stretchy")));
        // An integral's limits default beside it (intLim), and a document setting can move them under and over.
        var integral = ProjectExpression("<m:nary><m:sub>{0}</m:sub><m:sup>{1}</m:sup><m:e>{x}</m:e></m:nary>", mathProperties: "<m:intLim m:val=\"undOvr\"/>");
        Assert.Equal("munderover(∫,0,1) x", Shape(integral));
        Assert.Equal("false", (string?)integral.Descendants(MathMl + "mo").First().Attribute("stretchy"));
    }

    [Theory]
    [InlineData("", "true", "true")]
    [InlineData("<m:grow m:val=\"0\"/><m:shp m:val=\"match\"/>", "false", "false")]
    public void Delimiters_are_fences_that_grow_by_default(string properties, string stretchy, string symmetric)
    {
        var fences = ProjectExpression("<m:d><m:dPr>" + properties + "</m:dPr><m:e>{x}</m:e></m:d>").Descendants(MathMl + "mo").ToArray();
        Assert.Equal(new[] { "prefix", "postfix" }, fences.Select(f => (string?)f.Attribute("form")));
        Assert.All(fences, f =>
        {
            Assert.Equal("true", (string?)f.Attribute("fence"));
            Assert.Equal(stretchy, (string?)f.Attribute("stretchy"));
            Assert.Equal(symmetric, (string?)f.Attribute("symmetric"));
        });
    }

    [Theory]
    [InlineData("<m:sty m:val=\"b\"/>", "x", "𝐱")]
    [InlineData("<m:sty m:val=\"bi\"/>", "α2", "𝜶𝟐")]
    [InlineData("<m:scr m:val=\"double-struck\"/>", "R", "ℝ")]
    [InlineData("<m:scr m:val=\"script\"/><m:sty m:val=\"p\"/>", "L", "ℒ")]
    [InlineData("<m:scr m:val=\"fraktur\"/>", "g", "𝔤")]
    [InlineData("<m:scr m:val=\"sans-serif\"/><m:sty m:val=\"b\"/>", "A1", "𝗔𝟭")]
    [InlineData("<m:scr m:val=\"monospace\"/>", "z", "𝚣")]
    [InlineData("<m:lit/>", "x", "x")]
    public void Run_styles_select_the_Unicode_mathematical_alphabet(string properties, string text, string expected)
    {
        var expression = ProjectExpression("<m:r><m:rPr>" + properties + "</m:rPr><m:t>" + text + "</m:t></m:r>");
        Assert.Equal(expected, expression.Value);
        Assert.DoesNotContain(expression.Descendants(), e => e.Attribute("mathvariant") != null);
    }

    [Fact]
    public void Plain_style_letters_are_upright()
    {
        var single = Assert.Single(ProjectExpression("<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>d</m:t></m:r>").Descendants(MathMl + "mi"));
        Assert.Equal("normal", (string?)single.Attribute("mathvariant"));
        var word = Assert.Single(ProjectExpression("<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>max</m:t></m:r>").Descendants(MathMl + "mi"));
        Assert.Equal("max", word.Value);
        Assert.Null(word.Attribute("mathvariant"));
    }

    [Fact]
    public void Word_run_and_control_formatting_in_the_math_font_and_run_colour_are_drawn()
    {
        const string font = "<m:mathFont m:val=\"Cambria Math\"/>";
        const string fonts = "<w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/>";
        var formula = "<m:sSup><m:sSupPr><m:ctrlPr><w:rPr>" + fonts + "<w:i/></w:rPr></m:ctrlPr></m:sSupPr>"
            + "<m:e><m:r><w:rPr>" + fonts + "<w:color w:val=\"c00000\"/></w:rPr><m:t>x</m:t></m:r></m:e>"
            + "<m:sup>" + RunTwo + "<m:ctrlPr><w:rPr>" + fonts + "</w:rPr></m:ctrlPr></m:sup></m:sSup>";
        var expression = ProjectExpression(formula, mathProperties: font);
        Assert.Equal("msup(x,2)", Shape(expression));
        Assert.Equal("color: #C00000;", (string?)expression.Descendants(MathMl + "mi").Single().Attribute("style"));
    }

    [Fact]
    public void Matrix_columns_keep_their_justification()
    {
        var cells = ProjectExpression("<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val=\"1\"/><m:mcJc m:val=\"left\"/></m:mcPr></m:mc>"
            + "<m:mc><m:mcPr><m:count m:val=\"1\"/><m:mcJc m:val=\"center\"/></m:mcPr></m:mc></m:mcs></m:mPr>"
            + "<m:mr><m:e>{a}</m:e><m:e>{b}</m:e></m:mr></m:m>").Descendants(MathMl + "mtd").ToArray();
        // The -webkit- keyword is the one Chromium applies to a cell's math (checked against Word's render).
        Assert.Equal(new[] { "text-align: left;text-align: -webkit-left;", null }, cells.Select(c => (string?)c.Attribute("style")));
    }

    [Fact]
    public void Display_equation_arrays_keep_display_style_and_align_rows_at_their_points()
    {
        var table = Assert.Single(Convert(Expand("<m:eqArr><m:e>{x&amp;=1}</m:e><m:e>{y&amp;=2}</m:e></m:eqArr>"), true).Descendants(MathMl + "mtable"));
        Assert.Equal("true", (string?)table.Attribute("displaystyle"));
        Assert.Equal(new[] { "right", "left" }, table.Elements().First().Elements()
            .Select(c => ((string?)c.Attribute("style"))!.Split(';')[0].Split(' ')[1]));
        Assert.All(table.Descendants(MathMl + "mtd"), c => Assert.Matches("text-align: -webkit-(right|left);", (string?)c.Attribute("style")));
    }

    [Theory]
    [InlineData("", "bottom")]
    [InlineData("<m:barPr><m:pos m:val=\"top\"/></m:barPr>", "top")]
    public void A_bar_is_a_rule_as_wide_as_its_operand(string properties, string side)
    {
        // Word draws the bar across the whole operand. A stretchy bar character stays one character
        // wide in Cambria Math, so the bar is the operand box's own edge.
        var bar = ProjectExpression("<m:bar>" + properties + "<m:e>{x+y}</m:e></m:bar>").Descendants(MathMl + "mrow")
            .Single(e => e.Attribute("style") != null);
        Assert.Contains("border-" + side + ": 0.06em solid currentColor;", (string?)bar.Attribute("style"));
        Assert.Equal("x+y", bar.Value);
    }

    [Fact]
    public void A_linear_fraction_slash_is_set_tight()
    {
        var slash = Assert.Single(ProjectExpression("<m:f><m:fPr><m:type m:val=\"lin\"/></m:fPr><m:num>{a}</m:num><m:den>{b}</m:den></m:f>")
            .Descendants(MathMl + "mo"));
        Assert.Equal(("0", "0"), ((string?)slash.Attribute("lspace"), (string?)slash.Attribute("rspace")));
    }

    [Theory]
    [InlineData(64, true)]
    [InlineData(65, false)]
    public void An_equation_array_row_holds_at_most_64_aligned_cells(int cells, bool projected)
    {
        // Every row is padded to the widest one, so an unbounded row would multiply the whole array.
        var wide = string.Concat(Enumerable.Repeat("&amp;", cells - 1));
        var html = Convert(Expand("<m:eqArr><m:e>{x" + wide + "}</m:e><m:e>{y}</m:e><m:e>{z}</m:e></m:eqArr>"), false);
        Assert.Equal(projected ? cells * 3 : 0, html.Descendants(MathMl + "mtd").Count());
        Assert.Equal(projected ? 0 : 1, html.Descendants().Count(e => (string?)e.Attribute("data-content-type") == "MathEquation"));
    }

    [Fact]
    public void A_long_plain_text_run_is_one_identifier_built_in_linear_time()
    {
        // Appending to the identifier letter by letter copied the whole string each time: this run
        // did not finish in twenty minutes. Built once, it converts in a few seconds.
        var name = new string('a', 400_000);
        var identifier = Assert.Single(ProjectExpression("<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>" + name + "</m:t></m:r>")
            .Descendants(MathMl + "mi"));
        Assert.Equal(name.Length, identifier.Value.Length);
    }

    [Fact]
    public void A_run_colour_must_be_exactly_six_hex_digits()
    {
        // The schema rejects this value, so the document is built without validation.
        var html = Convert("<m:r><w:rPr><w:color w:val=\"FF0000&#10;\"/></w:rPr><m:t>x</m:t></m:r>", false, validate: false);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
        Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
    }

    [Fact]
    public void Border_boxes_draw_only_their_visible_sides()
    {
        var box = ProjectExpression("<m:borderBox><m:borderBoxPr><m:hideTop/><m:hideLeft m:val=\"1\"/></m:borderBoxPr><m:e>{x}</m:e></m:borderBox>")
            .Descendants(MathMl + "mrow").First(e => e.Attribute("style") != null);
        Assert.Contains("border-top-style: none;", (string?)box.Attribute("style"));
        Assert.Contains("border-left-style: none;", (string?)box.Attribute("style"));
        Assert.DoesNotContain("border-right-style", (string?)box.Attribute("style"));
    }

    // OMML keeps these nested operands full size in display math, where MathML would shrink them.
    [Theory]
    [InlineData("<m:m><m:mr><m:e><m:f><m:num>{x}</m:num><m:den>{2}</m:den></m:f></m:e></m:mr></m:m>", false)]
    [InlineData("<m:m><m:mr><m:e><m:f><m:num>{x}</m:num><m:den>{2}</m:den></m:f></m:e></m:mr></m:m>", true)]
    [InlineData("<m:f><m:num><m:nary><m:sub/><m:sup/><m:e>{x}</m:e></m:nary></m:num><m:den>{2}</m:den></m:f>", false)]
    [InlineData("<m:f><m:num><m:nary><m:sub/><m:sup/><m:e>{x}</m:e></m:nary></m:num><m:den>{2}</m:den></m:f>", true)]
    public void Nested_fractions_and_large_operators_keep_a_diagnostic_only_in_display_math(string formula, bool display)
    {
        var html = Convert(Expand(formula), display);
        Assert.Equal(display ? 0 : 1, html.Descendants(MathMl + "math").Count());
        Assert.Equal(display ? 1 : 0, html.Descendants().Count(e => (string?)e.Attribute("data-content-type") == "MathEquation"));
    }

    [Theory]
    [InlineData("<m:f><m:fPr><m:type m:val=\"skw\"/></m:fPr><m:num>{x}</m:num><m:den>{2}</m:den></m:f>")]
    [InlineData("<m:r><m:rPr><m:nor/></m:rPr><m:t>x</m:t></m:r>")]
    [InlineData("<m:r><m:rPr><m:sty m:val=\"b\"/></m:rPr><m:t>+</m:t></m:r>")]
    // A bold alphabet has digits only for 0-9; any other digit would be drawn unbolded.
    [InlineData("<m:r><m:rPr><m:sty m:val=\"b\"/></m:rPr><m:t>٣</m:t></m:r>")]
    [InlineData("<m:r><m:rPr><m:scr m:val=\"script\"/></m:rPr><m:t>α</m:t></m:r>")]
    [InlineData("<m:r><m:rPr><m:scr m:val=\"double-struck\"/><m:sty m:val=\"b\"/></m:rPr><m:t>R</m:t></m:r>")]
    [InlineData("<m:r><w:rPr><w:b/></w:rPr><m:t>x</m:t></m:r>")]
    [InlineData("<m:r><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr><m:t>x</m:t></m:r>")]
    [InlineData("<m:sSup><m:sSupPr><m:ctrlPr><w:rPr><w:color w:val=\"FF0000\"/></w:rPr></m:ctrlPr></m:sSupPr><m:e>{x}</m:e><m:sup>{2}</m:sup></m:sSup>")]
    [InlineData("<m:sSubSup><m:sSubSupPr><m:alnScr/></m:sSubSupPr><m:e>{x}</m:e><m:sub>{i}</m:sub><m:sup>{2}</m:sup></m:sSubSup>")]
    [InlineData("<m:borderBox><m:borderBoxPr><m:strikeH/></m:borderBoxPr><m:e>{x}</m:e></m:borderBox>")]
    [InlineData("<m:groupChr><m:groupChrPr><m:pos m:val=\"top\"/></m:groupChrPr><m:e>{x}</m:e></m:groupChr>")]
    [InlineData("<m:acc><m:accPr><m:chr m:val=\"\u20DB\"/></m:accPr><m:e>{x}</m:e></m:acc>")]
    [InlineData("<m:box><m:boxPr><m:opEmu/></m:boxPr><m:e>{x}</m:e></m:box>")]
    [InlineData("<m:eqArr><m:eqArrPr><m:rSp m:val=\"4\"/></m:eqArrPr><m:e>{x}</m:e></m:eqArr>")]
    [InlineData("<m:m><m:mPr><m:cSp m:val=\"200\"/></m:mPr><m:mr><m:e>{a}</m:e></m:mr></m:m>")]
    [InlineData("<m:d><m:e><m:argPr><m:argSz m:val=\"-1\"/></m:argPr>{x}</m:e></m:d>")]
    [InlineData("<m:nary><m:naryPr><m:chr m:val=\"∑\"/></m:naryPr><m:sub><m:phant><m:e>{i}</m:e></m:phant></m:sub><m:sup>{n}</m:sup><m:e>{x}</m:e></m:nary>")]
    public void Unsupported_properties_and_descendants_keep_the_whole_equation_diagnostic(string formula)
    {
        var html = Convert(Expand(formula), false);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
        Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
    }

    [Fact]
    public void A_Word_authored_equation_projects_completely_through_the_public_converter()
    {
        // Shaped like Word's own output: every run and properties element names the math font.
        const string fonts = "<w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr>";
        const string control = "<m:ctrlPr><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/><w:i/></w:rPr></m:ctrlPr>";
        static string R(string text, string style = "") => "<m:r>" + style + fonts + "<m:t>" + text + "</m:t></m:r>";
        var formula = R("f") + "<m:d><m:dPr>" + control + "</m:dPr><m:e>" + R("x") + "</m:e></m:d>" + R("=")
            + "<m:nary><m:naryPr><m:chr m:val=\"∑\"/>" + control + "</m:naryPr><m:sub>" + R("k=0") + "</m:sub><m:sup>" + R("∞") + "</m:sup><m:e>"
            + "<m:f><m:fPr>" + control + "</m:fPr><m:num><m:sSup><m:sSupPr>" + control + "</m:sSupPr><m:e>" + R("x") + "</m:e><m:sup>" + R("k") + "</m:sup></m:sSup></m:num>"
            + "<m:den>" + R("k!") + "</m:den></m:f>" + R("+")
            + "<m:func><m:funcPr>" + control + "</m:funcPr><m:fName>" + R("cos", "<m:rPr><m:sty m:val=\"p\"/></m:rPr>") + "</m:fName><m:e>"
            + "<m:acc><m:accPr><m:chr m:val=\"\u0303\"/>" + control + "</m:accPr><m:e>" + R("θ") + "</m:e></m:acc></m:e></m:func></m:e></m:nary>";
        var document = new WmlDocument("equation.docx", BuildDocument(formula, display: true,
            "<m:mathFont m:val=\"Cambria Math\"/><m:intLim m:val=\"subSup\"/><m:naryLim m:val=\"undOvr\"/>", "", ""));
        var html = WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings { RenderUnsupportedContentPlaceholders = true });
        Assert.DoesNotContain(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
        var math = Assert.Single(html.Descendants(MathMl + "math"));
        Assert.Equal("f ( x ) = munderover(∑,k = 0,∞) mfrac(msup(x,k),k !) + cos mspace() mover(θ,˜)", Shape(math.Elements().Single()));
    }

    private static XElement ProjectExpression(string formula, string mathProperties = "")
    {
        var math = Assert.Single(Convert(Expand(formula), false, mathProperties).Descendants(MathMl + "math"));
        return math.Elements().Single();
    }

    private static string Expand(string formula) =>
        System.Text.RegularExpressions.Regex.Replace(formula, "\\{([^}]*)\\}", m => "<m:r><m:t>" + m.Groups[1].Value + "</m:t></m:r>");

    private static string Shape(XElement element)
    {
        var children = element.Elements().Select(Shape).Where(s => s.Length > 0);
        return element.Name.LocalName switch
        {
            "mi" or "mn" or "mo" => element.Value,
            "mrow" when !element.HasAttributes => string.Join(" ", children),
            var name => name + "(" + string.Join(",", children) + ")",
        };
    }

    private static XElement Convert(string formula, bool display, string mathProperties = "", bool paginated = false,
        string paragraphStart = "", string styles = "", bool validate = true)
    {
        return XElement.Parse(HtmlConversionOps.ConvertToHtml(BuildDocument(formula, display, mathProperties, paragraphStart, styles, validate), new HtmlConversionOptions
        {
            RenderUnsupportedContentPlaceholders = true,
            FabricateCssClasses = false,
            PaginationMode = paginated ? (int)PaginationMode.Paginated : (int)PaginationMode.None,
        }));
    }

    private static byte[] BuildDocument(string formula, bool display, string mathProperties, string paragraphStart, string styles,
        bool validate = true)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var math = "<m:oMath>" + formula + "</m:oMath>";
            if (display) math = "<m:oMathPara>" + math + "</m:oMathPara>";
            main.PutXDocument(XDocument.Parse("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><w:body><w:p>" + paragraphStart + math + "</w:p><w:sectPr/></w:body></w:document>"));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse("<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + styles + "</w:styles>"));
            var settings = main.AddNewPart<DocumentSettingsPart>();
            settings.PutXDocument(XDocument.Parse("<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
                + (mathProperties == "" ? "" : "<m:mathPr>" + mathProperties + "</m:mathPr>") + "</w:settings>"));
            document.Save();
            if (validate) Assert.Empty(new OpenXmlValidator().Validate(document).Select(e => e.Description));
        }
        return stream.ToArray();
    }
}
