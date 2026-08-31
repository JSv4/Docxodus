#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.Json;
using Docxodus.Verification;

namespace Docxodus.Internal;

/// <summary>
/// Single owner of the <see cref="DocxDiff"/> wire contract. Both the WASM
/// bridge (<c>DocxDiffBridge</c>) and the stdio Python host
/// (<c>tools/python-host</c> dispatcher) route the shared diff entry points —
/// Compare, GetRevisions, GetEditScriptJson, and GetSemanticChangesJson — through
/// here, so the JSON shapes for settings (in) and revisions (out) live in exactly
/// one place. This mirrors the role <see cref="HtmlConversionOps"/> plays for HTML conversion.
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

    /// <summary>
    /// The requested products of ONE memoized comparison pass (issue #594). Unrequested
    /// products are null. Each non-null product is byte/string-identical to the corresponding
    /// single-product op on the same inputs and settings.
    /// </summary>
    public sealed record DocxDiffProducts(
        byte[]? RedlineBytes,
        string? RevisionsJson,
        string? EditScriptJson,
        string? SemanticChangesJson);

    /// <summary>
    /// Compare two DOCX byte arrays ONCE (via <see cref="DocxDiff.CreateComparison"/>) and return
    /// every requested product from that single pass — the facade counterpart of calling
    /// <see cref="Compare"/>, <see cref="GetRevisionsJson"/>, <see cref="GetEditScriptJson"/>, and
    /// <see cref="GetSemanticChangesJson"/> separately, each of which recomputes the diff.
    /// The semantic product runs its own pipeline pass (its reader differs), but shares the input
    /// snapshot; note that unlike the standalone <see cref="GetSemanticChangesJson"/>, the bytes
    /// here are opened as packages for the other products regardless.
    /// </summary>
    public static DocxDiffProducts CompareProducts(
        byte[] leftBytes,
        byte[] rightBytes,
        string? settingsJson,
        bool redline,
        bool revisions,
        bool editScript,
        bool semanticChanges)
    {
        var (left, right, settings) = Prepare(leftBytes, rightBytes, settingsJson);
        var comparison = DocxDiff.CreateComparison(left, right, settings);
        return new DocxDiffProducts(
            redline ? comparison.ToRedline().DocumentByteArray : null,
            revisions ? SerializeRevisions(comparison.GetRevisions()) : null,
            editScript ? comparison.GetEditScriptJson() : null,
            semanticChanges ? comparison.GetSemanticChangesJson(indented: false) : null);
    }

    /// <summary>
    /// Wire form of <see cref="CompareProducts"/> shared by the WASM bridge and the stdio host.
    /// <paramref name="productsJson"/> is a JSON array drawn from <c>"redline"</c>,
    /// <c>"revisions"</c>, <c>"editScript"</c>, <c>"semanticChanges"</c>; null/empty selects all
    /// four. Returns <c>{"redlineB64":…, "revisions":[…], "editScript":…, "semanticChanges":…}</c>
    /// with unrequested keys omitted — the nested values carry exactly the standalone wire shapes
    /// (the revisions array elements, the edit-script object, the canonical compact semantic
    /// object).
    /// </summary>
    public static string CompareProductsJson(
        byte[] leftBytes, byte[] rightBytes, string? settingsJson, string? productsJson)
    {
        var (redline, revisions, editScript, semanticChanges) = ParseProductSelection(productsJson);

        var products = CompareProducts(
            leftBytes, rightBytes, settingsJson, redline, revisions, editScript, semanticChanges);

        var sb = new StringBuilder(256);
        sb.Append('{');
        var first = true;
        void Key(string name)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append('"').Append(name).Append("\":");
        }

        if (products.RedlineBytes is { } bytes)
        {
            Key("redlineB64");
            sb.Append(DocxSessionJson.JsonString(Convert.ToBase64String(bytes)));
        }
        if (products.RevisionsJson is { } revisionsJson)
        {
            // Re-emit just the array so the envelope reads {"revisions":[…]} like the standalone op.
            using var parsed = JsonDocument.Parse(revisionsJson);
            Key("revisions");
            sb.Append(parsed.RootElement.GetProperty("revisions").GetRawText());
        }
        if (products.EditScriptJson is { } scriptJson)
        {
            // The standalone op returns the script indented; the envelope must stay a single
            // line (the stdio host frames responses as NDJSON), so re-emit it compact.
            Key("editScript");
            sb.Append(CompactJson(scriptJson));
        }
        if (products.SemanticChangesJson is { } semanticJson)
        {
            Key("semanticChanges");
            sb.Append(semanticJson);
        }
        sb.Append('}');
        return sb.ToString();
    }

    /// <summary>
    /// Compare ONE baseline against MANY candidates, reading the baseline once (issue #617).
    /// <paramref name="candidatesJson"/> is a JSON array of <c>{"name":…,"docB64":…}</c> — the same
    /// shape the consolidate ops take their reviewers in. Returns
    /// <c>{"results":[{"name":…, …products…}]}</c>, one entry per candidate in the order given, each
    /// carrying exactly the keys <see cref="CompareProductsJson"/> emits.
    /// </summary>
    /// <remarks>
    /// <para><b>Why a batch rather than a snapshot handle.</b> The .NET API for this is
    /// <see cref="DocxDiff.CreateSnapshot"/> — an object holding the parsed document. That object
    /// cannot cross a process or language boundary, and exposing it as a handle would put a
    /// memory-pinning lifetime on transports that have no good way to bound it. What every transport
    /// actually wants is the workload the snapshot exists for, so that is what they get: the baseline
    /// is snapshotted once inside this call and every candidate is compared against it.</para>
    /// <para>A candidate that throws does not fail the batch — its entry carries an <c>error</c>
    /// string instead of products, so one malformed counterparty markup cannot cost the caller the
    /// other ninety-nine comparisons.</para>
    /// </remarks>
    public static string CompareBatchJson(
        byte[] baselineBytes, string candidatesJson, string? settingsJson, string? productsJson)
    {
        if (baselineBytes == null || baselineBytes.Length == 0)
            throw new ArgumentException("No baseline document data provided", nameof(baselineBytes));

        var selection = ParseProductSelection(productsJson);
        var settings = ParseSettings(settingsJson);
        var candidates = ParseNamedDocuments(candidatesJson, nameof(candidatesJson));

        // ONE read of the baseline for the whole batch — the point of the op.
        var baseline = DocxDiff.CreateSnapshot(new WmlDocument("baseline.docx", baselineBytes), settings);

        var sb = new StringBuilder(1024);
        sb.Append("{\"results\":[");
        for (var i = 0; i < candidates.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var (name, bytes) = candidates[i];
            sb.Append("{\"name\":").Append(DocxSessionJson.JsonString(name));
            try
            {
                // The document is constructed INSIDE the try: a candidate that is not a readable
                // Wordprocessing package is that candidate's error, not the batch's.
                var comparison = DocxDiff.CreateComparison(
                    baseline,
                    DocxDiff.CreateSnapshot(new WmlDocument($"{name}.docx", bytes), settings),
                    settings);
                AppendProducts(sb, comparison, selection);
            }
            catch (Exception ex)
            {
                sb.Append(",\"error\":")
                  .Append(DocxSessionJson.JsonString($"{ex.GetType().Name}: {ex.Message}"));
            }
            sb.Append('}');
        }

        sb.Append("]}");
        return sb.ToString();
    }

    /// <summary>The four product flags a <c>products</c> array selects; null/empty selects all.</summary>
    private static (bool Redline, bool Revisions, bool EditScript, bool SemanticChanges)
        ParseProductSelection(string? productsJson)
    {
        if (string.IsNullOrWhiteSpace(productsJson)) return (true, true, true, true);

        bool redline = false, revisions = false, editScript = false, semanticChanges = false;
        using var doc = JsonDocument.Parse(productsJson);
        if (doc.RootElement.ValueKind != JsonValueKind.Array)
            throw new ArgumentException(
                "products must be a JSON array of \"redline\"/\"revisions\"/\"editScript\"/\"semanticChanges\"",
                nameof(productsJson));
        foreach (var entry in doc.RootElement.EnumerateArray())
        {
            switch (entry.ValueKind == JsonValueKind.String ? entry.GetString() : null)
            {
                case "redline": redline = true; break;
                case "revisions": revisions = true; break;
                case "editScript": editScript = true; break;
                case "semanticChanges": semanticChanges = true; break;
                default:
                    throw new ArgumentException(
                        $"unknown product {entry}; expected \"redline\", \"revisions\", \"editScript\", or \"semanticChanges\"",
                        nameof(productsJson));
            }
        }

        if (!(redline || revisions || editScript || semanticChanges))
            throw new ArgumentException("products selected nothing", nameof(productsJson));
        return (redline, revisions, editScript, semanticChanges);
    }

    /// <summary>Parse a <c>[{"name":…,"docB64":…}]</c> array. An entry without usable bytes is an
    /// error rather than a silent skip: a batch that quietly drops a candidate returns the wrong
    /// number of results.</summary>
    private static List<(string Name, byte[] Bytes)> ParseNamedDocuments(
        string json, string parameterName)
    {
        var documents = new List<(string, byte[])>();
        if (string.IsNullOrWhiteSpace(json)) return documents;

        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array)
            throw new ArgumentException("candidates must be a JSON array", parameterName);

        var index = 0;
        foreach (var element in doc.RootElement.EnumerateArray())
        {
            if (element.ValueKind != JsonValueKind.Object
                || !element.TryGetProperty("docB64", out var docB64)
                || docB64.ValueKind != JsonValueKind.String
                || docB64.GetString() is not { Length: > 0 } b64)
                throw new ArgumentException($"candidates[{index}] has no docB64", parameterName);

            var name = element.TryGetProperty("name", out var n) && n.ValueKind == JsonValueKind.String
                ? n.GetString()!
                : index.ToString(System.Globalization.CultureInfo.InvariantCulture);
            documents.Add((name, Convert.FromBase64String(b64)));
            index++;
        }

        return documents;
    }

    /// <summary>Append the selected products of one comparison as object members (leading comma),
    /// in the same wire shapes <see cref="CompareProductsJson"/> emits.</summary>
    private static void AppendProducts(
        StringBuilder sb, DocxDiffComparison comparison,
        (bool Redline, bool Revisions, bool EditScript, bool SemanticChanges) selection)
    {
        if (selection.Redline)
            sb.Append(",\"redlineB64\":").Append(DocxSessionJson.JsonString(
                Convert.ToBase64String(comparison.ToRedline().DocumentByteArray)));
        if (selection.Revisions)
        {
            using var parsed = JsonDocument.Parse(SerializeRevisions(comparison.GetRevisions()));
            sb.Append(",\"revisions\":").Append(parsed.RootElement.GetProperty("revisions").GetRawText());
        }

        // The envelope must stay a single line — the stdio host frames responses as NDJSON — so the
        // indented standalone script is re-emitted compact, exactly as CompareProductsJson does.
        if (selection.EditScript)
            sb.Append(",\"editScript\":").Append(CompactJson(comparison.GetEditScriptJson()));
        if (selection.SemanticChanges)
            sb.Append(",\"semanticChanges\":").Append(comparison.GetSemanticChangesJson(indented: false));
    }

    /// <summary>Compare two DOCX byte arrays; return the edit script as a JSON string.</summary>
    public static string GetEditScriptJson(byte[] leftBytes, byte[] rightBytes, string? settingsJson)
    {
        var (left, right, settings) = Prepare(leftBytes, rightBytes, settingsJson);
        return DocxDiff.GetEditScriptJson(left, right, settings);
    }

    /// <summary>
    /// Compare two DOCX byte arrays and return the public semantic-change schema as compact JSON.
    /// Compact output is the canonical wire form shared by WASM, npm, Python, and MCP.
    /// </summary>
    public static string GetSemanticChangesJson(
        byte[] leftBytes, byte[] rightBytes, string? settingsJson)
    {
        var settings = ParseSettings(settingsJson);
        return SemanticDiff.CompareJson(
            leftBytes,
            rightBytes,
            new SemanticDiffOptions { DiffSettings = settings },
            indented: false);
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
        if (TryGetBool(root, "preAcceptInputRevisions", out var preAcceptInputRevisions))
            settings.PreAcceptInputRevisions = preAcceptInputRevisions;
        if (TryGetBool(root, "preserveInputRevisions", out var preserveInputRevisions))
            settings.PreserveInputRevisions = preserveInputRevisions;
        if (TryGetBool(root, "normalizeRevisionAuthors", out var normalizeRevisionAuthors))
            settings.NormalizeRevisionAuthors = normalizeRevisionAuthors;
        if (TryGetBool(root, "crossParagraphTokenDiff", out var crossParagraphTokenDiff))
            settings.CrossParagraphTokenDiff = crossParagraphTokenDiff;

        return settings;
    }

    private static string CompactJson(string json)
    {
        using var doc = JsonDocument.Parse(json);
        using var buffer = new System.IO.MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer))
            doc.RootElement.WriteTo(writer);
        return Encoding.UTF8.GetString(buffer.ToArray());
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
    /// <summary>The requested products of ONE memoized consolidation pass.</summary>
    public sealed record DocxDiffConsolidatedProducts(
        byte[]? RedlineBytes,
        string? RevisionsJson,
        string? EditScriptJson,
        string? ConflictsJson);

    /// <summary>
    /// Consolidate ONCE (via <see cref="DocxDiff.CreateConsolidation"/>) and return every requested
    /// product from that single pass — the N-way counterpart of <see cref="CompareProducts"/>
    /// (issue #617). Calling the standalone ops separately reads the whole reviewer set once each:
    /// an <c>N</c>-reviewer set is <c>N+1</c> documents, so a caller that wants the redline and its
    /// attributed revisions was reading <c>2(N+1)</c> packages to answer a question needing
    /// <c>N+1</c>.
    /// </summary>
    public static DocxDiffConsolidatedProducts ConsolidateProducts(
        byte[] baseBytes, string reviewersJson, string? settingsJson,
        bool redline, bool revisions, bool editScript, bool conflicts)
    {
        var baseDoc = new WmlDocument("base.docx", baseBytes);
        var reviewers = ParseReviewers(reviewersJson);
        var settings = ParseConsolidateSettings(settingsJson);
        var consolidation = DocxDiff.CreateConsolidation(baseDoc, reviewers, settings);
        return new DocxDiffConsolidatedProducts(
            redline ? consolidation.Consolidate().DocumentByteArray : null,
            revisions ? SerializeConsolidatedRevisions(consolidation.GetConsolidatedRevisions()) : null,
            editScript ? consolidation.GetConsolidatedEditScriptJson() : null,
            conflicts ? SerializeConflicts(consolidation.GetConflicts()) : null);
    }

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
