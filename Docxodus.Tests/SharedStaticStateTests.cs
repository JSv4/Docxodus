using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Process-wide state the library keeps between documents must survive concurrent callers
/// (issue #974): the MCP server, the stdio host and NuGet consumers all run conversions on
/// more than one thread.
/// </summary>
public class SharedStaticStateTests
{
    [Fact]
    public async Task SymToChar_FromManyThreads_MapsEverySymbolOnceAndConsistently()
    {
        // Fonts unique to this run, so the symbols are new to the process-wide map even when other
        // tests have already populated it. Every font shares the same character codes, so all but
        // the first font of each code land in the private-use area: the path that read and wrote
        // the shared counter. The first font claims each code's own character for the whole
        // process, so the codes sit above the U+F020-U+F0FF range real symbol fonts use; claiming
        // U+F028 here would move Wingdings' U+F028 in other tests into the private-use area.
        var fonts = Enumerable.Range(0, 8).Select(_ => "Font-" + Guid.NewGuid().ToString("N")).ToArray();
        var codes = Enumerable.Range(0, 32).Select(i => 0xF700 + i).ToArray();
        var symbols = fonts.SelectMany(font => codes.Select(code => (font, code))).ToArray();

        var results = new char[8][];
        using var start = new Barrier(results.Length);
        var threads = Enumerable.Range(0, results.Length).Select(t => Task.Factory.StartNew(() =>
        {
            start.SignalAndWait();
            // Each thread walks the symbols in its own order, so threads race on the same keys.
            var order = t % 2 == 0 ? symbols : symbols.Reverse().ToArray();
            var mapped = new Dictionary<(string, int), char>();
            foreach (var (font, code) in order)
                mapped[(font, code)] = UnicodeMapper.SymToChar(font, code);
            results[t] = symbols.Select(s => mapped[s]).ToArray();
        }, TaskCreationOptions.LongRunning)).ToArray();

        var all = Task.WhenAll(threads);
        Assert.Same(all, await Task.WhenAny(all, Task.Delay(TimeSpan.FromMinutes(1))));
        await all;

        // Every thread saw the same character for the same symbol...
        for (var t = 1; t < results.Length; t++)
            Assert.Equal(results[0], results[t]);
        // ...distinct symbols got distinct characters...
        Assert.Equal(symbols.Length, results[0].Distinct().Count());
        // ...and each character maps back to the symbol it was issued for.
        for (var i = 0; i < symbols.Length; i++)
        {
            var sym = UnicodeMapper.CharToRunChild(results[0][i])!;
            Assert.Equal(W.sym, sym.Name);
            Assert.Equal(symbols[i].font, (string?)sym.Attribute(W.font));
            Assert.Equal(symbols[i].code, Convert.ToInt32((string)sym.Attribute(W._char)!, 16));
        }
    }

    [Fact]
    public void CharToRunChild_ForASymbol_ReturnsAnElementTheCallerOwns()
    {
        var font = "Font-" + Guid.NewGuid().ToString("N");
        var c = UnicodeMapper.SymToChar(font, 0xF741);

        var first = UnicodeMapper.CharToRunChild(c)!;
        first.SetAttributeValue(W.font, "Changed");
        new XElement(W.r, first);

        var second = UnicodeMapper.CharToRunChild(c)!;
        Assert.NotSame(first, second);
        Assert.Equal(font, (string?)second.Attribute(W.font));
        Assert.Null(second.Parent);
    }

    [Fact]
    public void ListItemTextDefaults_CannotBeChangedThroughASettingsInstance()
    {
        var settings = new ListItemRetrieverSettings();
        settings.ListItemTextImplementations["xx-XX"] = (_, _, _) => "x";

        Assert.False(new ListItemRetrieverSettings().ListItemTextImplementations.ContainsKey("xx-XX"));
        Assert.False(ListItemRetrieverSettings.DefaultListItemTextImplementations.ContainsKey("xx-XX"));
        Assert.False(new WmlToHtmlConverterSettings().ListItemImplementations.ContainsKey("xx-XX"));
    }

    [Fact]
    public void ListItemTextDefaults_CannotBeChangedThroughTheDefaultsProperty()
    {
        ListItemRetrieverSettings.DefaultListItemTextImplementations.Remove("fr-FR");

        Assert.True(ListItemRetrieverSettings.DefaultListItemTextImplementations.ContainsKey("fr-FR"));
        Assert.True(new ListItemRetrieverSettings().ListItemTextImplementations.ContainsKey("fr-FR"));
    }
}
