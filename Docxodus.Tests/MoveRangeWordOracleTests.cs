using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class MoveRangeWordOracleTests
{
    private static string FixtureRoot => Path.GetFullPath(Path.Combine(
        AppContext.BaseDirectory, "../../../../TestFiles/RP/MoveRangeMembership"));

    public static IEnumerable<object[]> Endpoints()
    {
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureRoot, "word-endpoints.json")));
        foreach (var fixture in json.RootElement.GetProperty("cases").EnumerateArray())
        {
            foreach (bool accept in new[] { true, false })
            {
                var endpoint = fixture.GetProperty("word").GetProperty(accept ? "accept-all" : "reject-all");
                yield return new object[]
                {
                    fixture.GetProperty("id").GetString()!,
                    fixture.GetProperty("inputSha256").GetString()!,
                    accept,
                    endpoint.GetProperty("paragraphs").EnumerateArray().Select(p => p.GetString()!).ToArray(),
                };
            }
        }
    }

    public static IEnumerable<object[]> SessionEndpoints() => Endpoints().SelectMany(endpoint =>
        new[] { false, true }.Select(selective => endpoint.Concat(new object[] { selective }).ToArray()));

    [Theory]
    [MemberData(nameof(Endpoints))]
    public void Processor_MatchesWordEndpoint(string id, string sourceHash, bool accept, string[] expected)
    {
        var bytes = Input(id, sourceHash);
        var source = new WmlDocument(id + ".docx", bytes);
        var result = accept ? RevisionProcessor.AcceptRevisions(source) : RevisionProcessor.RejectRevisions(source);
        AssertEndpoint(result.DocumentByteArray, expected);
    }

    [Theory]
    [MemberData(nameof(SessionEndpoints))]
    public void Session_MatchesWordEndpoint(
        string id, string sourceHash, bool accept, string[] expected, bool selective)
    {
        using var session = new DocxSession(Input(id, sourceHash));
        var move = Assert.Single(session.ListRevisions());
        Assert.Equal("move", move.Type);
        Assert.True(move.ResolutionStatus == RevisionResolutionStatus.Supported, move.Diagnostic?.Code);
        var result = selective
            ? accept ? session.AcceptRevision(move.Id) : session.RejectRevision(move.Id)
            : accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
        Assert.True(result.Success, result.Error?.Message);
        Assert.Empty(session.ListRevisions());
        AssertEndpoint(session.Save(), expected);
    }

    private static byte[] Input(string id, string sourceHash)
    {
        var bytes = File.ReadAllBytes(Path.Combine(FixtureRoot, id + ".docx"));
        Assert.Equal(sourceHash, Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        PackageValidation.AssertValid(bytes, FileFormatVersions.Microsoft365);
        return bytes;
    }

    private static void AssertEndpoint(byte[] bytes, string[] expected)
    {
        PackageValidation.AssertValid(bytes, FileFormatVersions.Microsoft365);
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var xml = document.MainDocumentPart!.GetXDocument();
        var paragraphs = xml.Root!.Element(W.body)!.Elements(W.p)
            .Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value)))
            .ToArray();
        Assert.Equal(expected, paragraphs);
        var markers = new[] { W.moveFrom, W.moveTo, W.moveFromRangeStart, W.moveFromRangeEnd,
            W.moveToRangeStart, W.moveToRangeEnd };
        Assert.Empty(xml.Descendants().Where(e => markers.Contains(e.Name)));
    }
}
