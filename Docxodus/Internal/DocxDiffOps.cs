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

    /// <summary>
    /// Accept every tracked revision in a redlined DOCX — materialize the "right"/revised side. The byte-in,
    /// byte-out counterpart of <see cref="Compare"/>: <c>Accept(Compare(left, right))</c> ≡ <c>right</c> (per-block
    /// text level). Exposing accept/reject lets clients verify the round-trip contract, not just the shape of the
    /// redline. Wraps <see cref="RevisionProcessor.AcceptRevisions(WmlDocument)"/>.
    /// </summary>
    public static byte[] AcceptRevisions(byte[] bytes)
    {
        if (bytes == null || bytes.Length == 0)
            throw new ArgumentException("No document data provided", nameof(bytes));
        return RevisionProcessor.AcceptRevisions(new WmlDocument("redline.docx", bytes)).DocumentByteArray;
    }

    /// <summary>
    /// Reject every tracked revision in a redlined DOCX — materialize the "left"/original side:
    /// <c>Reject(Compare(left, right))</c> ≡ <c>left</c> (per-block text level). Wraps
    /// <see cref="RevisionProcessor.RejectRevisions(WmlDocument)"/>.
    /// </summary>
    public static byte[] RejectRevisions(byte[] bytes)
    {
        if (bytes == null || bytes.Length == 0)
            throw new ArgumentException("No document data provided", nameof(bytes));
        return RevisionProcessor.RejectRevisions(new WmlDocument("redline.docx", bytes)).DocumentByteArray;
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
        if (root.TryGetProperty("culture", out var culture) && culture.ValueKind == JsonValueKind.String)
        {
            var name = culture.GetString();
            if (!string.IsNullOrEmpty(name))
                settings.Culture = System.Globalization.CultureInfo.GetCultureInfo(name!);
        }
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
        if (TryGetBool(root, "compareHeadersFooters", out var compareHf))
            settings.CompareHeadersFooters = compareHf;
        if (TryGetBool(root, "trackBlockFormatChanges", out var trackBlockFmt))
            settings.TrackBlockFormatChanges = trackBlockFmt;

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
        sb.Append(']');
        // Additive (block-format-change family, 2026-07-03): which property container the change describes.
        // Always emitted; "run" (the default and the historical behavior) for every pre-campaign revision.
        sb.Append(",\"scope\":").Append(DocxSessionJson.JsonString(FormatChangeScopeWire(fc.Scope)));
        sb.Append('}');
    }

    private static string FormatChangeScopeWire(DocxDiffFormatChangeScope scope) => scope switch
    {
        DocxDiffFormatChangeScope.Run => "run",
        DocxDiffFormatChangeScope.Paragraph => "paragraph",
        DocxDiffFormatChangeScope.TableCell => "tableCell",
        DocxDiffFormatChangeScope.TableRow => "tableRow",
        DocxDiffFormatChangeScope.Table => "table",
        DocxDiffFormatChangeScope.Section => "section",
        _ => "run",
    };

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

    // ---- Consolidate (composite N-way) wire surface --------------------------
    //
    // Reviewers arrive as a JSON array, each element carrying the reviewer's
    // author name and their full revised DOCX as base64:
    //   [{"author":"Bob","docB64":"<base64 docx>"}, ...]
    // Settings reuse the diff settings object (same camelCase fields parsed by
    // ParseSettings) and additionally carry an optional integer
    // "conflictResolution" (0=BaseWins,1=FirstReviewerWins,2=StackAll).

    /// <summary>Consolidate N reviewer documents against a base; return the merged DOCX bytes.</summary>
    public static byte[] Consolidate(byte[] baseBytes, string reviewersJson, string? settingsJson)
    {
        if (baseBytes == null || baseBytes.Length == 0)
            throw new ArgumentException("No base document data provided", nameof(baseBytes));

        var baseDoc = new WmlDocument("base.docx", baseBytes);
        var reviewers = ParseReviewers(reviewersJson);
        var settings = ParseConsolidateSettings(settingsJson);
        return DocxDiff.Consolidate(baseDoc, reviewers, settings).DocumentByteArray;
    }

    /// <summary>Consolidate; return the merged revision list as JSON.</summary>
    public static string GetConsolidatedRevisionsJson(byte[] baseBytes, string reviewersJson, string? settingsJson)
    {
        if (baseBytes == null || baseBytes.Length == 0)
            throw new ArgumentException("No base document data provided", nameof(baseBytes));

        var baseDoc = new WmlDocument("base.docx", baseBytes);
        var revs = DocxDiff.GetConsolidatedRevisions(
            baseDoc, ParseReviewers(reviewersJson), ParseConsolidateSettings(settingsJson));
        return SerializeConsolidatedRevisions(revs);
    }

    /// <summary>Consolidate; return the merged edit script as JSON.</summary>
    public static string GetConsolidatedEditScriptJson(byte[] baseBytes, string reviewersJson, string? settingsJson)
    {
        if (baseBytes == null || baseBytes.Length == 0)
            throw new ArgumentException("No base document data provided", nameof(baseBytes));

        var baseDoc = new WmlDocument("base.docx", baseBytes);
        return DocxDiff.GetConsolidatedEditScriptJson(
            baseDoc, ParseReviewers(reviewersJson), ParseConsolidateSettings(settingsJson));
    }

    /// <summary>Consolidate; return the per-token conflict report as JSON.</summary>
    public static string GetConflictsJson(byte[] baseBytes, string reviewersJson, string? settingsJson)
    {
        if (baseBytes == null || baseBytes.Length == 0)
            throw new ArgumentException("No base document data provided", nameof(baseBytes));

        var baseDoc = new WmlDocument("base.docx", baseBytes);
        var conflicts = DocxDiff.GetConflicts(
            baseDoc, ParseReviewers(reviewersJson), ParseConsolidateSettings(settingsJson));
        return SerializeConflicts(conflicts);
    }

    /// <summary>
    /// Parse the reviewer transport array
    /// <c>[{"author":"Bob","docB64":"&lt;base64 docx&gt;"}, ...]</c> into
    /// <see cref="DocxDiffReviewer"/> objects. A null/empty/whitespace string or
    /// a non-array yields an empty list; elements missing <c>docB64</c> are skipped.
    /// </summary>
    public static List<DocxDiffReviewer> ParseReviewers(string reviewersJson)
    {
        var reviewers = new List<DocxDiffReviewer>();
        if (string.IsNullOrWhiteSpace(reviewersJson))
            return reviewers;

        using var doc = JsonDocument.Parse(reviewersJson);
        var root = doc.RootElement;
        if (root.ValueKind != JsonValueKind.Array)
            return reviewers;

        foreach (var el in root.EnumerateArray())
        {
            if (el.ValueKind != JsonValueKind.Object)
                continue;
            if (!el.TryGetProperty("docB64", out var docB64) || docB64.ValueKind != JsonValueKind.String)
                continue;
            var b64 = docB64.GetString();
            if (string.IsNullOrEmpty(b64))
                continue;

            var author = el.TryGetProperty("author", out var a) && a.ValueKind == JsonValueKind.String
                ? a.GetString()!
                : string.Empty;

            reviewers.Add(new DocxDiffReviewer
            {
                Author = author,
                Document = new WmlDocument("reviewer.docx", Convert.FromBase64String(b64!)),
            });
        }
        return reviewers;
    }

    /// <summary>
    /// Parse consolidate settings: the same JSON object carries the diff fields
    /// (parsed via <see cref="ParseSettings"/>) plus an optional integer
    /// <c>conflictResolution</c> (0=BaseWins,1=FirstReviewerWins,2=StackAll).
    /// </summary>
    public static DocxDiffConsolidateSettings ParseConsolidateSettings(string? settingsJson)
    {
        var settings = new DocxDiffConsolidateSettings { Diff = ParseSettings(settingsJson) };
        if (string.IsNullOrWhiteSpace(settingsJson))
            return settings;

        using var doc = JsonDocument.Parse(settingsJson);
        var root = doc.RootElement;
        if (root.ValueKind == JsonValueKind.Object &&
            root.TryGetProperty("conflictResolution", out var cr) && cr.ValueKind == JsonValueKind.Number)
        {
            settings.ConflictResolution = cr.GetInt32() switch
            {
                1 => ConflictResolution.FirstReviewerWins,
                2 => ConflictResolution.StackAll,
                _ => ConflictResolution.BaseWins,
            };
        }
        return settings;
    }

    /// <summary>
    /// Serialize a consolidated revision list to the wire JSON shape — mirrors
    /// <see cref="SerializeRevisions"/> with the added <c>conflictId</c> field
    /// (<c>author</c> is already present on every revision).
    /// </summary>
    public static string SerializeConsolidatedRevisions(IReadOnlyList<DocxDiffConsolidatedRevision> revisions)
    {
        var sb = new StringBuilder(256);
        sb.Append("{\"revisions\":[");
        for (var i = 0; i < revisions.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendConsolidatedRevision(sb, revisions[i]);
        }
        sb.Append("]}");
        return sb.ToString();
    }

    private static void AppendConsolidatedRevision(StringBuilder sb, DocxDiffConsolidatedRevision r)
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

        sb.Append(",\"conflictId\":");
        sb.Append(r.ConflictId is { } cid ? cid.ToString(CultureInfo.InvariantCulture) : "null");

        sb.Append('}');
    }

    /// <summary>
    /// Serialize the conflict report to the wire JSON shape
    /// <c>{"conflicts":[{"id","baseAnchor","tokenStart","tokenEnd","policy","competitors":[{"author","resultText"}]}]}</c>.
    /// <c>policy</c> is the integer-coded <see cref="ConflictResolution"/> that was applied.
    /// </summary>
    public static string SerializeConflicts(IReadOnlyList<DocxDiffConflict> conflicts)
    {
        var sb = new StringBuilder(256);
        sb.Append("{\"conflicts\":[");
        for (var i = 0; i < conflicts.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendConflict(sb, conflicts[i]);
        }
        sb.Append("]}");
        return sb.ToString();
    }

    private static void AppendConflict(StringBuilder sb, DocxDiffConflict c)
    {
        sb.Append("{\"id\":").Append(c.Id.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"baseAnchor\":").Append(DocxSessionJson.JsonString(c.BaseAnchor));
        sb.Append(",\"tokenStart\":").Append(c.TokenStart.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"tokenEnd\":").Append(c.TokenEnd.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"policy\":").Append(((int)c.AppliedPolicy).ToString(CultureInfo.InvariantCulture));

        sb.Append(",\"competitors\":[");
        for (var i = 0; i < c.Competitors.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var comp = c.Competitors[i];
            sb.Append("{\"author\":").Append(DocxSessionJson.JsonString(comp.Author));
            sb.Append(",\"resultText\":").Append(DocxSessionJson.JsonString(comp.ResultText));
            sb.Append('}');
        }
        sb.Append("]}");
    }
}
