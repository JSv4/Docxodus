#nullable enable

using System.IO;
using System.Text.Json;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

public class IrDiagnosticJsonTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    private static WmlDocument Fixture(string name) =>
        new(Path.Combine(TestFilesDir.FullName, name));

    [Fact]
    public void DiagnosticJson_TwoReads_ByteIdentical()
    {
        var doc = Fixture("HC031-Complicated-Document.docx");

        var json1 = IrDiagnosticJson.Write(IrReader.Read(doc));
        var json2 = IrDiagnosticJson.Write(IrReader.Read(doc));

        Assert.Equal(json1, json2);
    }

    [Fact]
    public void DiagnosticJson_SimpleDocument_ContainsExpectedStructure()
    {
        var doc = IrTestDocuments.Create("Hello world", "Second line");

        var json = IrDiagnosticJson.Write(IrReader.Read(doc));

        // Valid JSON.
        using var parsed = JsonDocument.Parse(json);
        var root = parsed.RootElement;

        Assert.Equal("body", root.GetProperty("scope").GetString());
        var blocks = root.GetProperty("blocks");
        Assert.Equal(JsonValueKind.Array, blocks.ValueKind);

        // Two paragraphs, each carrying an anchor, content hash, and the right text.
        Assert.Equal(2, blocks.GetArrayLength());
        var first = blocks[0];
        Assert.Equal("paragraph", first.GetProperty("type").GetString());
        Assert.Matches("^p:body:[0-9a-f]{32}$", first.GetProperty("anchor").GetString());
        Assert.Matches("^[0-9a-f]{64}$", first.GetProperty("contentHash").GetString());
        Assert.Matches("^[0-9a-f]{64}$", first.GetProperty("formatFingerprint").GetString());

        var firstInlines = first.GetProperty("inlines");
        Assert.Equal("text", firstInlines[0].GetProperty("kind").GetString());
        Assert.Equal("Hello world", firstInlines[0].GetProperty("text").GetString());

        var second = blocks[1];
        Assert.Equal("Second line", second.GetProperty("inlines")[0].GetProperty("text").GetString());
    }

    [Fact]
    public void DiagnosticJson_IsValidJson_ForComplexFixture()
    {
        var doc = Fixture("HC001-5DayTourPlanTemplate.docx");

        var json = IrDiagnosticJson.Write(IrReader.Read(doc));

        // Parsing throws if the output is not well-formed JSON.
        using var parsed = JsonDocument.Parse(json);
        Assert.Equal("body", parsed.RootElement.GetProperty("scope").GetString());
        Assert.True(parsed.RootElement.GetProperty("blocks").GetArrayLength() > 0);
    }
}
