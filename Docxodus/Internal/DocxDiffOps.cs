#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.Json;

namespace Docxodus.Internal;

/// <summary>
/// Single owner of the <see cref="DocxDiff"/> wire contract. Both the WASM
/// bridge (<c>DocxDiffBridge</c>) and the stdio Python host
/// (<c>tools/python-host</c> dispatcher) route the three diff entry points —
/// Compare, GetRevisions, GetEditScriptJson — through here, so the JSON shapes
/// for settings (in) and revisions (out) live in exactly one place. This
/// mirrors the role <see cref="HtmlConversionOps"/> plays for HTML conversion.
///
/// <para>Settings arrive as a JSON object (the transport mirror of
/// <see cref="DocxDiffSettings"/>); every field is optional and an omitted
/// field uses the .NET default. Revisions are serialized by hand (no reflection
/// <c>JsonSerializer</c>) to stay trim/AOT-safe, consistent with the rest of the
/// core bridge layer (<see cref="DocxSessionJson"/>).</para>
/// </summary>
internal static class DocxDiffOps
{
    /// <summary>Compare two DOCX byte arrays; return the redlined DOCX bytes.</summary>
    public static byte[] Compare(byte[] leftBytes, byte[] rightBytes, string? settingsJson)
    {
        var (left, right, settings) = Prepare(leftBytes, rightBytes, settingsJson);
        return DocxDiff.Compare(left, right, settings).DocumentByteArray;
    }

    /// <summary>Compare two DOCX byte arrays; return the revision list as a JSON string.</summary>
    public static string GetRevisionsJson(byte[] leftBytes, byte[] rightBytes, string? settingsJson)
    {
        var (left, right, settings) = Prepare(leftBytes, rightBytes, settingsJson);
        var revisions = DocxDiff.GetRevisions(left, right, settings);
        return SerializeRevisions(revisions);
    }

    /// <summary>Compare two DOCX byte arrays; return the edit script as a JSON string.</summary>
    public static string GetEditScriptJson(byte[] leftBytes, byte[] rightBytes, string? settingsJson)
    {
        var (left, right, settings) = Prepare(leftBytes, rightBytes, settingsJson);
        return DocxDiff.GetEditScriptJson(left, right, settings);
    }

    private static (WmlDocument left, WmlDocument right, DocxDiffSettings settings) Prepare(
        byte[] leftBytes, byte[] rightBytes, string? settingsJson)
    {
        if (leftBytes == null || leftBytes.Length == 0)
            throw new ArgumentException("No left document data provided", nameof(leftBytes));
        if (rightBytes == null || rightBytes.Length == 0)
            throw new ArgumentException("No right document data provided", nameof(rightBytes));

        var left = new WmlDocument("left.docx", leftBytes);
        var right = new WmlDocument("right.docx", rightBytes);
        return (left, right, ParseSettings(settingsJson));
    }

    /// <summary>
    /// Parse the transport JSON object into <see cref="DocxDiffSettings"/>. A
    /// null/empty/whitespace string or a non-object yields the defaults; each
    /// field falls back to its default when absent. Enum fields are integer-coded
    /// to match the TypeScript enum positions.
    /// </summary>
    public static DocxDiffSettings ParseSettings(string? settingsJson)
    {
        var settings = new DocxDiffSettings();
        if (string.IsNullOrWhiteSpace(settingsJson))
            return settings;

        using var doc = JsonDocument.Parse(settingsJson);
        var root = doc.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            return settings;

        if (root.TryGetProperty("authorForRevisions", out var author) && author.ValueKind == JsonValueKind.String)
            settings.AuthorForRevisions = author.GetString()!;
        if (TryGetBool(root, "deterministic", out var deterministic))
            settings.Deterministic = deterministic;
        if (root.TryGetProperty("dateTimeForRevisions", out var date) && date.ValueKind == JsonValueKind.String)
            settings.DateTimeForRevisions = date.GetString();
        if (TryGetBool(root, "caseInsensitive", out var ci))
            settings.CaseInsensitive = ci;
        if (TryGetBool(root, "conflateBreakingAndNonbreakingSpaces", out var conflate))
            settings.ConflateBreakingAndNonbreakingSpaces = conflate;
        if (root.TryGetProperty("wordSeparators", out var seps) && seps.ValueKind == JsonValueKind.String)
        {
            var s = seps.GetString();
            if (!string.IsNullOrEmpty(s))
                settings.WordSeparators = s!.ToCharArray();
        }
        if (TryGetBool(root, "detectMoves", out var detectMoves))
            settings.DetectMoves = detectMoves;
        if (root.TryGetProperty("moveSimilarityThreshold", out var sim) && sim.ValueKind == JsonValueKind.Number)
            settings.MoveSimilarityThreshold = sim.GetDouble();
        if (root.TryGetProperty("moveMinimumWordCount", out var minWords) && minWords.ValueKind == JsonValueKind.Number)
            settings.MoveMinimumWordCount = minWords.GetInt32();
        if (root.TryGetProperty("revisionGranularity", out var gran) && gran.ValueKind == JsonValueKind.Number)
            settings.RevisionGranularity = gran.GetInt32() == 1
                ? DocxDiffRevisionGranularity.WmlComparerCompatible
                : DocxDiffRevisionGranularity.Fine;
        if (root.TryGetProperty("formatComparison", out var fmt) && fmt.ValueKind == JsonValueKind.Number)
            settings.FormatComparison = fmt.GetInt32() == 1
                ? DocxDiffFormatComparison.Full
                : DocxDiffFormatComparison.ModeledOnly;

        return settings;
    }

    private static bool TryGetBool(JsonElement root, string name, out bool value)
    {
        if (root.TryGetProperty(name, out var v) &&
            (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False))
        {
            value = v.GetBoolean();
            return true;
        }
        value = false;
        return false;
    }

    /// <summary>
    /// Serialize a revision list to the wire JSON shape
    /// <c>{"revisions":[{revisionType,text,author,date,moveGroupId,isMoveSource,formatChange,leftAnchor,rightAnchor}]}</c>.
    /// Built by hand (no reflection serializer) to stay trim/AOT-safe.
    /// </summary>
    public static string SerializeRevisions(IReadOnlyList<DocxDiffRevision> revisions)
    {
        var sb = new StringBuilder(256);
        sb.Append("{\"revisions\":[");
        for (var i = 0; i < revisions.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendRevision(sb, revisions[i]);
        }
        sb.Append("]}");
        return sb.ToString();
    }

    private static void AppendRevision(StringBuilder sb, DocxDiffRevision r)
    {
        sb.Append("{\"revisionType\":").Append(DocxSessionJson.JsonString(r.Type.ToString()));
        sb.Append(",\"text\":").Append(DocxSessionJson.JsonString(r.Text));
        sb.Append(",\"author\":").Append(DocxSessionJson.JsonString(r.Author));
        sb.Append(",\"date\":").Append(DocxSessionJson.JsonString(r.Date));

        sb.Append(",\"moveGroupId\":");
        sb.Append(r.MoveGroupId is { } mg ? mg.ToString(CultureInfo.InvariantCulture) : "null");

        sb.Append(",\"isMoveSource\":");
        sb.Append(r.IsMoveSource is { } ms ? (ms ? "true" : "false") : "null");

        sb.Append(",\"formatChange\":");
        if (r.FormatChange is { } fc)
            AppendFormatChange(sb, fc);
        else
            sb.Append("null");

        sb.Append(",\"leftAnchor\":");
        sb.Append(r.LeftAnchor is { } la ? DocxSessionJson.JsonString(la) : "null");
        sb.Append(",\"rightAnchor\":");
        sb.Append(r.RightAnchor is { } ra ? DocxSessionJson.JsonString(ra) : "null");

        sb.Append('}');
    }

    private static void AppendFormatChange(StringBuilder sb, DocxDiffFormatChange fc)
    {
        sb.Append("{\"oldProperties\":");
        AppendStringMap(sb, fc.OldProperties);
        sb.Append(",\"newProperties\":");
        AppendStringMap(sb, fc.NewProperties);
        sb.Append(",\"changedPropertyNames\":[");
        for (var i = 0; i < fc.ChangedPropertyNames.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(DocxSessionJson.JsonString(fc.ChangedPropertyNames[i]));
        }
        sb.Append("]}");
    }

    private static void AppendStringMap(StringBuilder sb, IReadOnlyDictionary<string, string> map)
    {
        sb.Append('{');
        var first = true;
        foreach (var kvp in map)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(DocxSessionJson.JsonString(kvp.Key)).Append(':').Append(DocxSessionJson.JsonString(kvp.Value));
        }
        sb.Append('}');
    }
}
