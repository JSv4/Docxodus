#nullable enable

using System.Collections.Generic;
using System.Text;
using System.Text.Json;

namespace Docxodus.Internal;

/// <summary>
/// Shared JSON serialization + parsing for the DocxSession bridge wire format.
/// Both the WASM JSExport bridge and the stdio NDJSON host emit and consume the
/// shapes defined here, so the TypeScript and Python clients see identical JSON.
/// All output is camelCase; clients that prefer snake_case (e.g. Python dataclasses)
/// convert during deserialization.
/// </summary>
internal static class DocxSessionJson
{
    // ─── Parsers ────────────────────────────────────────────────────────

    public static Position ParsePos(string s) =>
        string.Equals(s, "before", System.StringComparison.OrdinalIgnoreCase) ? Position.Before : Position.After;

    public static PageMap ParsePageMap(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return ParsePageMap(doc.RootElement);
    }

    public static PageMap ParsePageMap(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object)
            throw new FormatException("PageMap must be a JSON object");
        EnsureOnlyProperties(root, "PageMap", "schemaVersion", "mode", "availability",
            "documentVersion", "rendererFingerprint", "pages", "fragments");

        var pages = new List<PageMapPage>();
        var pagesElement = RequiredProperty(root, "pages", JsonValueKind.Array, "PageMap");
        foreach (var page in pagesElement.EnumerateArray())
        {
            if (page.ValueKind != JsonValueKind.Object)
                throw new FormatException("PageMap pages must be objects");
            EnsureOnlyProperties(page, "PageMap page", "pageNumber", "pageInSection", "width",
                "height", "sectionIndex", "pageName");

            int? sectionIndex = null;
            if (page.TryGetProperty("sectionIndex", out var sectionIndexElement))
            {
                if (sectionIndexElement.ValueKind != JsonValueKind.Number
                    || !sectionIndexElement.TryGetInt32(out var parsedSectionIndex))
                    throw new FormatException("PageMap page sectionIndex must be an integer when supplied");
                sectionIndex = parsedSectionIndex;
            }

            pages.Add(new PageMapPage
            {
                PageNumber = RequiredInt32(page, "pageNumber", "PageMap page"),
                PageInSection = RequiredInt32(page, "pageInSection", "PageMap page"),
                Width = RequiredDouble(page, "width", "PageMap page"),
                Height = RequiredDouble(page, "height", "PageMap page"),
                SectionIndex = sectionIndex,
                PageName = RequiredString(page, "pageName", "PageMap page"),
            });
        }

        var fragments = new List<PageMapFragment>();
        var fragmentsElement = RequiredProperty(root, "fragments", JsonValueKind.Array, "PageMap");
        foreach (var fragment in fragmentsElement.EnumerateArray())
        {
            if (fragment.ValueKind != JsonValueKind.Object)
                throw new FormatException("PageMap fragments must be objects");
            EnsureOnlyProperties(fragment, "PageMap fragment", "fragmentId", "anchorId",
                "fragmentIndex", "pageNumber", "geometry", "story", "inTableCell");
            var geometry = RequiredProperty(fragment, "geometry", JsonValueKind.Object, "PageMap fragment");
            EnsureOnlyProperties(geometry, "PageMap geometry", "x", "y", "width", "height");
            fragments.Add(new PageMapFragment
            {
                FragmentId = RequiredString(fragment, "fragmentId", "PageMap fragment"),
                AnchorId = RequiredString(fragment, "anchorId", "PageMap fragment"),
                FragmentIndex = RequiredInt32(fragment, "fragmentIndex", "PageMap fragment"),
                PageNumber = RequiredInt32(fragment, "pageNumber", "PageMap fragment"),
                Geometry = new PageMapRect(
                    RequiredDouble(geometry, "x", "PageMap geometry"),
                    RequiredDouble(geometry, "y", "PageMap geometry"),
                    RequiredDouble(geometry, "width", "PageMap geometry"),
                    RequiredDouble(geometry, "height", "PageMap geometry")),
                Story = ParsePageMapStory(RequiredString(fragment, "story", "PageMap fragment")),
                InTableCell = RequiredBoolean(fragment, "inTableCell", "PageMap fragment"),
            });
        }

        return new PageMap
        {
            SchemaVersion = RequiredInt32(root, "schemaVersion", "PageMap"),
            Mode = ParsePageMapMode(RequiredString(root, "mode", "PageMap")),
            Availability = ParsePageMapAvailability(RequiredString(root, "availability", "PageMap")),
            DocumentVersion = RequiredInt64(root, "documentVersion", "PageMap"),
            RendererFingerprint = RequiredString(root, "rendererFingerprint", "PageMap"),
            Pages = pages,
            Fragments = fragments,
        };
    }

    public static PageCitationRequest? ParsePageCitationRequest(JsonElement root, string key = "citation")
    {
        if (!root.TryGetProperty(key, out var value))
            return null;
        if (value.ValueKind != JsonValueKind.Object)
            throw new FormatException($"{key} must be a JSON object");
        EnsureOnlyProperties(value, key, "documentVersion", "rendererFingerprint");
        return new PageCitationRequest(
            RequiredInt64(value, "documentVersion", key),
            RequiredString(value, "rendererFingerprint", key));
    }

    private static JsonElement RequiredProperty(
        JsonElement root, string name, JsonValueKind kind, string owner)
    {
        if (!root.TryGetProperty(name, out var value) || value.ValueKind != kind)
            throw new FormatException($"{owner} requires {name} with JSON type {kind}");
        return value;
    }

    private static void EnsureOnlyProperties(JsonElement root, string owner, params string[] allowed)
    {
        foreach (var property in root.EnumerateObject())
        {
            if (System.Array.IndexOf(allowed, property.Name) < 0)
                throw new FormatException($"{owner} contains unknown property {property.Name}");
        }
    }

    private static string RequiredString(JsonElement root, string name, string owner) =>
        RequiredProperty(root, name, JsonValueKind.String, owner).GetString()!;

    private static bool RequiredBoolean(JsonElement root, string name, string owner)
    {
        if (!root.TryGetProperty(name, out var value)
            || value.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            throw new FormatException($"{owner} requires {name} with JSON type Boolean");
        return value.GetBoolean();
    }

    private static int RequiredInt32(JsonElement root, string name, string owner)
    {
        if (!root.TryGetProperty(name, out var value)
            || value.ValueKind != JsonValueKind.Number
            || !value.TryGetInt32(out var result))
            throw new FormatException($"{owner} requires integer {name}");
        return result;
    }

    private static long RequiredInt64(JsonElement root, string name, string owner)
    {
        if (!root.TryGetProperty(name, out var value)
            || value.ValueKind != JsonValueKind.Number
            || !value.TryGetInt64(out var result))
            throw new FormatException($"{owner} requires integer {name}");
        return result;
    }

    private static double RequiredDouble(JsonElement root, string name, string owner)
    {
        if (!root.TryGetProperty(name, out var value)
            || value.ValueKind != JsonValueKind.Number
            || !value.TryGetDouble(out var result))
            throw new FormatException($"{owner} requires numeric {name}");
        return result;
    }

    private static PageMapMode ParsePageMapMode(string? value) => value switch
    {
        "paginated" => PageMapMode.Paginated,
        "continuous" => PageMapMode.Continuous,
        _ => throw new FormatException($"Unknown PageMap mode: {value}"),
    };

    private static PageMapAvailability ParsePageMapAvailability(string? value) => value switch
    {
        "available" => PageMapAvailability.Available,
        "unavailable" => PageMapAvailability.Unavailable,
        _ => throw new FormatException($"Unknown PageMap availability: {value}"),
    };

    private static PageMapStory ParsePageMapStory(string? value) => value switch
    {
        "body" => PageMapStory.Body,
        "header" => PageMapStory.Header,
        "footer" => PageMapStory.Footer,
        "footnote" => PageMapStory.Footnote,
        "endnote" => PageMapStory.Endnote,
        "comment" => PageMapStory.Comment,
        _ => throw new FormatException($"Unknown PageMap story: {value}"),
    };

    public static HeaderFooterKind ParseHeaderFooterKind(string? s) =>
        (s?.ToLowerInvariant()) switch
        {
            "first" => HeaderFooterKind.First,
            "even" => HeaderFooterKind.Even,
            _ => HeaderFooterKind.Default,
        };

    public static PageNumberField ParsePageNumberField(string? s) =>
        (s?.ToLowerInvariant()) switch
        {
            "totalpages" or "numpages" => PageNumberField.TotalPages,
            _ => PageNumberField.CurrentPage,
        };

    /// <summary>
    /// Wire → <see cref="NumberFormat"/>, where an EMPTY or absent token means "no format
    /// specified" rather than a default — the distinction the tri-state page-numbering surface is
    /// built on. Unlike <see cref="NumberFormats.ParseOoxml"/>, an unrecognized non-empty token also
    /// reads as null so a typo cannot silently become <see cref="NumberFormat.Decimal"/>; the typed
    /// clients (TypeScript union, Python enum) constrain the value long before it gets here.
    /// </summary>
    public static NumberFormat? ParseNumberFormatOrNull(string? s) => s switch
    {
        "decimal" => NumberFormat.Decimal,
        "upperLetter" => NumberFormat.UpperLetter,
        "lowerLetter" => NumberFormat.LowerLetter,
        "upperRoman" => NumberFormat.UpperRoman,
        "lowerRoman" => NumberFormat.LowerRoman,
        "bullet" => NumberFormat.Bullet,
        _ => null,
    };

    /// <summary>
    /// Parse a <see cref="PageNumberingOp"/> from <c>{ start?: int, format?: string }</c>. An
    /// omitted field stays null, which is what "leave this attribute unchanged" means on the op.
    /// </summary>
    public static PageNumberingOp ParsePageNumberingOp(string json)
    {
        if (string.IsNullOrEmpty(json)) return new PageNumberingOp();
        using var doc = JsonDocument.Parse(json);
        return ParsePageNumberingOp(doc.RootElement);
    }

    /// <summary>Element overload — the stdio host already holds a parsed request body.</summary>
    public static PageNumberingOp ParsePageNumberingOp(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object) return new PageNumberingOp();
        int? start = root.TryGetProperty("start", out var s) && s.ValueKind == JsonValueKind.Number
            ? s.GetInt32()
            : null;
        return new PageNumberingOp
        {
            Start = start,
            Format = ParseNumberFormatOrNull(TryGetString(root, "format", null)),
        };
    }

    /// <summary>Lenient wire-name → enum: unknown/absent falls back to Accept, mirroring
    /// ParseSettings' historical behavior. Strict callers (MCP set_mode) do their own switch.</summary>
    public static TrackedChangeMode ParseTrackedChangeMode(string? mode) => mode switch
    {
        "render_inline" => TrackedChangeMode.RenderInline,
        "strip_deletions" => TrackedChangeMode.StripDeletions,
        _ => TrackedChangeMode.Accept,
    };

    public static string TrackedChangeModeName(TrackedChangeMode mode) => mode switch
    {
        TrackedChangeMode.RenderInline => "render_inline",
        TrackedChangeMode.StripDeletions => "strip_deletions",
        _ => "accept",
    };

    public static DocxSessionSettings ParseSettings(string settingsJson)
    {
        if (string.IsNullOrEmpty(settingsJson)) return new DocxSessionSettings();
        using var doc = JsonDocument.Parse(settingsJson);
        var root = doc.RootElement;

        // Defaults are read from the settings object itself rather than repeated as literals here,
        // so the wire default can never drift from the .NET default (it had, for undoDepth).
        var defaults = new DocxSessionSettings();

        int undoDepth = TryGetInt(root, "undoDepth", defaults.UndoDepth);
        long undoMemoryBudgetBytes =
            TryGetLong(root, "undoMemoryBudgetBytes", defaults.UndoMemoryBudgetBytes);
        bool validateRawOps = TryGetBool(root, "validateRawOps", false);
        var tracked = ParseTrackedChangeMode(TryGetString(root, "trackedChanges", "accept"));
        var revisionAuthor = TryGetString(root, "revisionAuthor", null);
        bool persistAnchorIds = TryGetBool(root, "persistAnchorIds", false);
        bool smartQuotes = TryGetBool(root, "smartQuotes", false);
        bool emitMarkdownPatch = TryGetBool(root, "emitMarkdownPatch", true);
        bool captureInitialProjection = TryGetBool(root, "captureInitialProjection", true);
        var projectionSettings = root.TryGetProperty("projectionSettings", out var ps) && ps.ValueKind == JsonValueKind.Object
            ? ParseProjectionSettings(ps)
            : new WmlToMarkdownConverterSettings();
        return new DocxSessionSettings
        {
            UndoDepth = undoDepth,
            UndoMemoryBudgetBytes = undoMemoryBudgetBytes,
            ValidateRawOps = validateRawOps,
            TrackedChanges = tracked,
            RevisionAuthor = revisionAuthor,
            PersistAnchorIds = persistAnchorIds,
            SmartQuotes = smartQuotes,
            EmitMarkdownPatch = emitMarkdownPatch,
            CaptureInitialProjection = captureInitialProjection,
            ProjectionSettings = projectionSettings,
        };
    }

    /// <summary>
    /// Parse a JSON object into <see cref="WmlToMarkdownConverterSettings"/>. Mirrors
    /// the <c>MarkdownProjectionSettings</c> TS interface and the
    /// <c>MarkdownProjectionSettingsDto</c> WASM DTO — numeric enum fields use the
    /// same flag/value layout as the .NET enums. Unknown / missing fields fall back
    /// to <see cref="WmlToMarkdownConverterSettings"/> defaults.
    /// </summary>
    public static WmlToMarkdownConverterSettings ParseProjectionSettings(JsonElement root)
    {
        var settings = new WmlToMarkdownConverterSettings();
        if (root.ValueKind != JsonValueKind.Object) return settings;
        if (root.TryGetProperty("scopes", out var sc) && sc.ValueKind == JsonValueKind.Number)
            settings.Scopes = (ProjectionScopes)sc.GetInt32();
        if (root.TryGetProperty("headingLevelOffset", out var hl) && hl.ValueKind == JsonValueKind.Number)
            settings.HeadingLevelOffset = hl.GetInt32();
        if (root.TryGetProperty("anchorMode", out var am) && am.ValueKind == JsonValueKind.Number)
            settings.AnchorMode = (AnchorRenderMode)am.GetInt32();
        if (root.TryGetProperty("tableMode", out var tm) && tm.ValueKind == JsonValueKind.Number)
            settings.TableMode = (TableRenderMode)tm.GetInt32();
        if (root.TryGetProperty("tableInlineCellMax", out var tic) && tic.ValueKind == JsonValueKind.Number)
            settings.TableInlineCellMax = tic.GetInt32();
        if (root.TryGetProperty("trackedChanges", out var tc) && tc.ValueKind == JsonValueKind.Number)
            settings.TrackedChanges = (TrackedChangeMode)tc.GetInt32();
        if (root.TryGetProperty("resolveNumbering", out var rn) && (rn.ValueKind == JsonValueKind.True || rn.ValueKind == JsonValueKind.False))
            settings.ResolveNumbering = rn.GetBoolean();
        if (root.TryGetProperty("emptyParagraphs", out var ep) && ep.ValueKind == JsonValueKind.Number)
            settings.EmptyParagraphs = (EmptyParagraphMode)ep.GetInt32();
        if (root.TryGetProperty("anchorIdRendering", out var air) && air.ValueKind == JsonValueKind.Number)
            settings.AnchorIdRendering = (AnchorIdRendering)air.GetInt32();
        return settings;
    }

    public static FormatOp ParseFormatOp(string json)
    {
        if (string.IsNullOrEmpty(json)) return new FormatOp();
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        return new FormatOp
        {
            Bold = TryGetBoolNullable(root, "bold"),
            Italic = TryGetBoolNullable(root, "italic"),
            Underline = TryGetBoolNullable(root, "underline"),
            Strike = TryGetBoolNullable(root, "strike"),
            Code = TryGetBoolNullable(root, "code"),
            Color = TryGetString(root, "color", null),
            RunStyle = TryGetString(root, "runStyle", null),
            VertAlign = TryGetString(root, "vertAlign", null),
            FontSizePts = TryGetDoubleNullable(root, "fontSizePts"),
            FontFamily = TryGetString(root, "fontFamily", null),
        };
    }

    /// <summary>Parse the common optimistic-mutation guard object used by every transport.</summary>
    public static MutationPreconditions? ParseMutationPreconditions(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        using var doc = JsonDocument.Parse(json);
        return ParseMutationPreconditions(doc.RootElement);
    }

    public static MutationPreconditions? ParseMutationPreconditions(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object) return null;
        TextRangePrecondition? range = null;
        if (root.TryGetProperty("expectedTextRange", out var r) && r.ValueKind == JsonValueKind.Object
            && r.TryGetProperty("start", out var start) && start.ValueKind == JsonValueKind.Number
            && r.TryGetProperty("length", out var length) && length.ValueKind == JsonValueKind.Number
            && r.TryGetProperty("text", out var text) && text.ValueKind == JsonValueKind.String)
        {
            range = new TextRangePrecondition(start.GetInt32(), length.GetInt32(), text.GetString() ?? string.Empty);
        }
        return new MutationPreconditions
        {
            ExpectedVersion = root.TryGetProperty("expectedVersion", out var version)
                && version.ValueKind == JsonValueKind.Number ? version.GetInt64() : null,
            AnchorId = TryGetString(root, "anchorId", null),
            ExpectedContentHash = TryGetString(root, "expectedContentHash", null),
            ExpectedText = TryGetString(root, "expectedText", null),
            ExpectedTextRange = range,
            ExpectedKind = TryGetString(root, "expectedKind", null),
            ExpectedScope = TryGetString(root, "expectedScope", null),
            ExpectedMatchCount = TryGetIntNullable(root, "expectedMatchCount"),
        };
    }

    /// <summary>
    /// Parse a ParagraphFormatOp wire object: { alignment?: "left"|"center"|"right"|"justify",
    /// indentDelta?: int (twips), firstLineIndent?/hangingIndent?: int (twips, mutually exclusive),
    /// spacingBefore?/spacingAfter?: int (twips), lineSpacing?: int (240ths of a line under "auto",
    /// twips under "exact"/"atLeast"), lineSpacingRule?: "auto"|"exact"|"atLeast",
    /// pageBreakBefore?: bool, topBorder?/bottomBorder?: BorderEdge, clearBorders?: bool }.
    /// Missing fields leave that property unchanged.
    /// </summary>
    public static ParagraphFormatOp ParseParagraphFormatOp(string json)
    {
        if (string.IsNullOrEmpty(json)) return new ParagraphFormatOp();
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        ParagraphAlignment? align = TryGetString(root, "alignment", null)?.ToLowerInvariant() switch
        {
            "left" => ParagraphAlignment.Left,
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            "justify" or "both" => ParagraphAlignment.Justify,
            _ => null,
        };
        LineSpacingRule? lineRule = TryGetString(root, "lineSpacingRule", null)?.ToLowerInvariant() switch
        {
            "auto" => LineSpacingRule.Auto,
            "exact" or "exactly" => LineSpacingRule.Exact,
            "atleast" => LineSpacingRule.AtLeast,
            _ => null,
        };
        return new ParagraphFormatOp
        {
            Alignment = align,
            IndentDelta = TryGetIntNullable(root, "indentDelta"),
            FirstLineIndent = TryGetIntNullable(root, "firstLineIndent"),
            HangingIndent = TryGetIntNullable(root, "hangingIndent"),
            SpacingBefore = TryGetIntNullable(root, "spacingBefore"),
            SpacingAfter = TryGetIntNullable(root, "spacingAfter"),
            LineSpacing = TryGetIntNullable(root, "lineSpacing"),
            LineSpacingRule = lineRule,
            PageBreakBefore = TryGetBoolNullable(root, "pageBreakBefore"),
            TopBorder = ParseBorderEdge(root, "topBorder"),
            BottomBorder = ParseBorderEdge(root, "bottomBorder"),
            ClearBorders = TryGetBoolNullable(root, "clearBorders"),
        };
    }

    /// <summary>
    /// Parse a <see cref="ParagraphBorderEdge"/> from a named object property
    /// ({ style?, size?, color?, space? }), or null when the property is absent/not an object.
    /// </summary>
    public static ParagraphBorderEdge? ParseBorderEdge(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var e) || e.ValueKind != JsonValueKind.Object) return null;
        return new ParagraphBorderEdge
        {
            Style = TryGetString(e, "style", null),
            Size = e.TryGetProperty("size", out var s) && s.ValueKind == JsonValueKind.Number ? s.GetInt32() : null,
            Color = TryGetString(e, "color", null),
            Space = e.TryGetProperty("space", out var sp) && sp.ValueKind == JsonValueKind.Number ? sp.GetInt32() : null,
        };
    }

    /// <summary>
    /// Parse a <see cref="TableInsertOptions"/> wire object:
    /// { borderless?: bool, cellContents?: string[], cellAlignment?: "left"|"center"|"right"|"justify",
    ///   columnWidths?: number[] (twips, one per column) }.
    /// </summary>
    public static TableInsertOptions ParseTableInsertOptions(string json)
    {
        if (string.IsNullOrEmpty(json)) return new TableInsertOptions();
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        List<string>? cells = null;
        if (root.TryGetProperty("cellContents", out var cc) && cc.ValueKind == JsonValueKind.Array)
        {
            cells = new List<string>(cc.GetArrayLength());
            foreach (var item in cc.EnumerateArray())
                cells.Add(item.ValueKind == JsonValueKind.String ? (item.GetString() ?? string.Empty) : string.Empty);
        }
        ParagraphAlignment? align = TryGetString(root, "cellAlignment", null)?.ToLowerInvariant() switch
        {
            "left" => ParagraphAlignment.Left,
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            "justify" or "both" => ParagraphAlignment.Justify,
            _ => null,
        };
        List<int>? widths = null;
        if (root.TryGetProperty("columnWidths", out var cw) && cw.ValueKind == JsonValueKind.Array)
        {
            widths = new List<int>(cw.GetArrayLength());
            foreach (var item in cw.EnumerateArray())
                widths.Add(item.ValueKind == JsonValueKind.Number ? item.GetInt32() : 0);
        }
        return new TableInsertOptions
        {
            Borderless = TryGetBool(root, "borderless", false),
            CellContents = cells,
            CellAlignment = align,
            ColumnWidths = widths,
        };
    }

    /// <summary>Parse a JSON array of integers (e.g. column widths in twips). Non-number entries
    /// become 0 so the session-side validation rejects them; empty/absent json → empty list.</summary>
    public static List<int> ParseIntArray(string json)
    {
        var result = new List<int>();
        if (string.IsNullOrEmpty(json)) return result;
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind == JsonValueKind.Array)
            foreach (var item in doc.RootElement.EnumerateArray())
                result.Add(item.ValueKind == JsonValueKind.Number ? item.GetInt32() : 0);
        return result;
    }

    /// <summary>Parse a <see cref="TableBorderSpec"/> ({ scope?, style?, size?, color? });
    /// empty json → all defaults (thin single border on every edge).</summary>
    public static TableBorderSpec ParseTableBorderSpec(string json)
    {
        if (string.IsNullOrEmpty(json)) return new TableBorderSpec();
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        var scope = TryGetString(root, "scope", null)?.ToLowerInvariant() switch
        {
            "outside" => TableBorderScope.Outside,
            "inside" => TableBorderScope.Inside,
            _ => TableBorderScope.All,
        };
        TryGetIntNullable(root, "size", out var size);
        return new TableBorderSpec
        {
            Scope = scope,
            Style = TryGetString(root, "style", null),
            Size = size,
            Color = TryGetString(root, "color", null),
        };
    }

    /// <summary>"row" → <see cref="TableShadingScope.Row"/>; anything else → Cell.</summary>
    public static TableShadingScope ParseTableShadingScope(string? scope) =>
        string.Equals(scope, "row", System.StringComparison.OrdinalIgnoreCase)
            ? TableShadingScope.Row : TableShadingScope.Cell;

    /// <summary>Parse the absorbed-content policy token for <see cref="DocxSession.MergeCells"/>:
    /// "discard" | "reject"; anything else (including null) keeps the lossless default.</summary>
    public static TableMergeContent ParseTableMergeContent(string? content) =>
        content?.ToLowerInvariant() switch
        {
            "discard" => TableMergeContent.Discard,
            "reject" => TableMergeContent.Reject,
            _ => TableMergeContent.Append,
        };

    /// <summary>Parse the OOXML row-height rule token used by the bridge.</summary>
    public static TableRowHeightRule ParseTableRowHeightRule(string? rule) =>
        rule?.ToLowerInvariant() switch
        {
            "auto" => TableRowHeightRule.Auto,
            "exact" => TableRowHeightRule.Exact,
            _ => TableRowHeightRule.AtLeast,
        };

    /// <summary>
    /// Parse a list-format kind token (case-insensitive camelCase of the <see cref="ListFormat"/>
    /// member: "bullet", "decimal", "lowerLetter", "upperRoman", "decimalParenthesis", …; "none"
    /// or anything unrecognized maps to <see cref="ListFormat.None"/>, matching the historical
    /// leniency of this parser).
    /// </summary>
    public static ListFormat ParseListFormat(string? kind) => kind?.ToLowerInvariant() switch
    {
        "bullet" => ListFormat.Bullet,
        "decimal" or "number" or "numbered" => ListFormat.Decimal,
        "lowerletter" => ListFormat.LowerLetter,
        "upperletter" => ListFormat.UpperLetter,
        "lowerroman" => ListFormat.LowerRoman,
        "upperroman" => ListFormat.UpperRoman,
        "decimalparenthesis" => ListFormat.DecimalParenthesis,
        "lowerletterparenthesis" => ListFormat.LowerLetterParenthesis,
        "upperletterparenthesis" => ListFormat.UpperLetterParenthesis,
        "lowerromanparenthesis" => ListFormat.LowerRomanParenthesis,
        "upperromanparenthesis" => ListFormat.UpperRomanParenthesis,
        _ => ListFormat.None,
    };

    public static FindOptions? ParseFindOptions(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object) return null;
        var scopes = (ProjectionScopes)TryGetInt(root, "scopes", (int)ProjectionScopes.All);
        return new FindOptions
        {
            IgnoreCase = TryGetBool(root, "ignoreCase", false),
            IgnoreWhitespace = TryGetBool(root, "ignoreWhitespace", false),
            KindFilter = TryGetString(root, "kindFilter", null),
            Scopes = scopes,
            ScopeFilter = TryGetString(root, "scopeFilter", null),
            CitationRequest = ParsePageCitationRequest(root),
        };
    }

    public static int TryGetInt(JsonElement root, string name, int fallback) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : fallback;

    public static long TryGetLong(JsonElement root, string name, long fallback) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt64() : fallback;

    public static bool TryGetBool(JsonElement root, string name, bool fallback) =>
        root.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False) ? v.GetBoolean() : fallback;

    public static bool? TryGetBoolNullable(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False) ? v.GetBoolean() : (bool?)null;

    public static int? TryGetIntNullable(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : (int?)null;

    public static string? TryGetString(JsonElement root, string name, string? fallback) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : fallback;

    public static double? TryGetDoubleNullable(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetDouble() : (double?)null;

    public static ImageInsertOptions ParseImageInsertOptions(string json)
    {
        if (string.IsNullOrEmpty(json)) return new ImageInsertOptions();
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        RequireObject(root, "image options");
        return new ImageInsertOptions
        {
            Placement = ParseImagePlacement(StrictString(root, "placement", "inline")),
            WidthPoints = StrictDoubleNullable(root, "widthPoints"),
            HeightPoints = StrictDoubleNullable(root, "heightPoints"),
            PreserveAspect = StrictBool(root, "preserveAspect", true),
            AltText = StrictString(root, "altText", null),
            Title = StrictString(root, "title", null),
            FloatingLayout = ParseOptionalFloatingImageLayout(root),
        };
    }

    private static FloatingImageLayout? ParseOptionalFloatingImageLayout(JsonElement root)
    {
        if (!root.TryGetProperty("floatingLayout", out var layout)) return null;
        if (layout.ValueKind != JsonValueKind.Object)
            throw new System.ArgumentException("floatingLayout must be a JSON object");
        return ParseFloatingImageLayout(layout);
    }

    public static FloatingImageLayout ParseFloatingImageLayout(string json)
    {
        using var document = JsonDocument.Parse(json);
        return ParseFloatingImageLayout(document.RootElement);
    }

    public static FloatingImageLayout ParseFloatingImageLayout(JsonElement root)
    {
        RequireObject(root, "floating layout");
        long? horizontalOffset = StrictInt64Nullable(root, "horizontalOffsetEmu");
        long? verticalOffset = StrictInt64Nullable(root, "verticalOffsetEmu");
        var horizontalAlignment = ParseHorizontalAlignment(StrictString(root, "horizontalAlignment", null));
        var verticalAlignment = ParseVerticalAlignment(StrictString(root, "verticalAlignment", null));
        if (horizontalAlignment is not null && !root.TryGetProperty("horizontalOffsetEmu", out _))
            horizontalOffset = null;
        if (verticalAlignment is not null && !root.TryGetProperty("verticalOffsetEmu", out _))
            verticalOffset = null;
        return new FloatingImageLayout
        {
            HorizontalRelativeFrom = ParseHorizontalReference(
                StrictString(root, "horizontalRelativeFrom", "column")),
            HorizontalOffsetEmu = horizontalOffset ?? (horizontalAlignment is null ? 0 : null),
            HorizontalAlignment = horizontalAlignment,
            VerticalRelativeFrom = ParseVerticalReference(
                StrictString(root, "verticalRelativeFrom", "paragraph")),
            VerticalOffsetEmu = verticalOffset ?? (verticalAlignment is null ? 0 : null),
            VerticalAlignment = verticalAlignment,
            WrapMode = ParseWrapMode(StrictString(root, "wrapMode", "square")),
            WrapSide = ParseWrapSide(StrictString(root, "wrapSide", "both_sides")),
            DistanceTopEmu = StrictInt64(root, "distanceTopEmu", 0),
            DistanceBottomEmu = StrictInt64(root, "distanceBottomEmu", 0),
            DistanceLeftEmu = StrictInt64(root, "distanceLeftEmu", 0),
            DistanceRightEmu = StrictInt64(root, "distanceRightEmu", 0),
            RelativeHeight = StrictUInt32(root, "relativeHeight", 251658240),
            BehindDocument = StrictBool(root, "behindDocument", false),
            Locked = StrictBool(root, "locked", false),
            LayoutInCell = StrictBool(root, "layoutInCell", true),
            AllowOverlap = StrictBool(root, "allowOverlap", true),
        };
    }

    public static (double? Width, double? Height, bool PreserveAspect) ParseImageDimensions(string json)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        RequireObject(root, "image dimensions");
        return (StrictDoubleNullable(root, "widthPoints"),
            StrictDoubleNullable(root, "heightPoints"), StrictBool(root, "preserveAspect", true));
    }

    private static void RequireObject(JsonElement root, string description)
    { if (root.ValueKind != JsonValueKind.Object) throw new System.ArgumentException($"{description} must be a JSON object"); }
    private static string? StrictString(JsonElement root, string name, string? fallback)
    {
        if (!root.TryGetProperty(name, out var value)) return fallback;
        if (value.ValueKind == JsonValueKind.Null) return null;
        if (value.ValueKind != JsonValueKind.String) throw new System.ArgumentException($"{name} must be a string or null");
        return value.GetString();
    }
    private static double? StrictDoubleNullable(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null) return null;
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetDouble(out var parsed))
            throw new System.ArgumentException($"{name} must be a number or null");
        return parsed;
    }
    private static bool StrictBool(JsonElement root, string name, bool fallback)
    {
        if (!root.TryGetProperty(name, out var value)) return fallback;
        if (value.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            throw new System.ArgumentException($"{name} must be a boolean");
        return value.GetBoolean();
    }
    private static long StrictInt64(JsonElement root, string name, long fallback) =>
        StrictInt64Nullable(root, name) ?? fallback;
    private static long? StrictInt64Nullable(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null) return null;
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt64(out var parsed))
            throw new System.ArgumentException($"{name} must be a 64-bit integer or null");
        return parsed;
    }
    private static uint StrictUInt32(JsonElement root, string name, uint fallback)
    {
        if (!root.TryGetProperty(name, out var value)) return fallback;
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetUInt32(out var parsed))
            throw new System.ArgumentException($"{name} must be an integer from 0 through {uint.MaxValue}");
        return parsed;
    }

    private static ImagePlacement ParseImagePlacement(string? token) => token switch
    { "inline" => ImagePlacement.Inline, "floating" => ImagePlacement.Floating, _ => (ImagePlacement)(-1) };
    private static ImageHorizontalReference ParseHorizontalReference(string? token) => token switch
    { "page" => ImageHorizontalReference.Page, "margin" => ImageHorizontalReference.Margin,
      "column" => ImageHorizontalReference.Column, "character" => ImageHorizontalReference.Character,
      _ => ImageHorizontalReference.Unknown };
    private static ImageVerticalReference ParseVerticalReference(string? token) => token switch
    { "page" => ImageVerticalReference.Page, "margin" => ImageVerticalReference.Margin,
      "paragraph" => ImageVerticalReference.Paragraph, "line" => ImageVerticalReference.Line,
      _ => ImageVerticalReference.Unknown };
    private static ImageHorizontalAlignment? ParseHorizontalAlignment(string? token) => token switch
    { null => null, "left" => ImageHorizontalAlignment.Left, "center" => ImageHorizontalAlignment.Center,
      "right" => ImageHorizontalAlignment.Right, "inside" => ImageHorizontalAlignment.Inside,
      "outside" => ImageHorizontalAlignment.Outside, _ => ImageHorizontalAlignment.Unknown };
    private static ImageVerticalAlignment? ParseVerticalAlignment(string? token) => token switch
    { null => null, "top" => ImageVerticalAlignment.Top, "center" => ImageVerticalAlignment.Center,
      "bottom" => ImageVerticalAlignment.Bottom, "inside" => ImageVerticalAlignment.Inside,
      "outside" => ImageVerticalAlignment.Outside, _ => ImageVerticalAlignment.Unknown };
    private static ImageWrapMode ParseWrapMode(string? token) => token switch
    { "none" => ImageWrapMode.None, "square" => ImageWrapMode.Square, "tight" => ImageWrapMode.Tight,
      "through" => ImageWrapMode.Through, "top_and_bottom" => ImageWrapMode.TopAndBottom,
      _ => ImageWrapMode.Unknown };
    private static ImageWrapSide ParseWrapSide(string? token) => token switch
    { "both_sides" => ImageWrapSide.BothSides, "left" => ImageWrapSide.Left,
      "right" => ImageWrapSide.Right, "largest" => ImageWrapSide.Largest,
      _ => ImageWrapSide.Unknown };

    // ─── Serializers ────────────────────────────────────────────────────

    public static string Serialize(EditResult r)
    {
        var sb = new StringBuilder(256);
        sb.Append("{\"success\":").Append(r.Success ? "true" : "false");
        if (r.Error is not null)
        {
            sb.Append(",\"error\":{")
              .Append("\"code\":\"").Append(EnumToSnake(r.Error.Code)).Append('"')
              .Append(",\"message\":").Append(JsonString(r.Error.Message));
            if (r.Error.AnchorId is not null)
                sb.Append(",\"anchorId\":").Append(JsonString(r.Error.AnchorId));
            if (r.Error.Precondition is { } p)
            {
                sb.Append(",\"precondition\":{")
                  .Append("\"condition\":").Append(JsonString(p.Condition))
                  .Append(",\"expected\":");
                AppendJsonValue(sb, p.Expected);
                sb.Append(",\"actual\":");
                AppendJsonValue(sb, p.Actual);
                sb.Append(",\"currentVersion\":").Append(p.CurrentVersion);
                if (p.CurrentTarget is { } target)
                {
                    sb.Append(",\"currentTarget\":{")
                      .Append("\"exists\":").Append(target.Exists ? "true" : "false");
                    if (target.AnchorId is not null)
                        sb.Append(",\"anchorId\":").Append(JsonString(target.AnchorId));
                    if (target.Kind is not null)
                        sb.Append(",\"kind\":").Append(JsonString(target.Kind));
                    if (target.Scope is not null)
                        sb.Append(",\"scope\":").Append(JsonString(target.Scope));
                    if (target.ContentHash is not null)
                        sb.Append(",\"contentHash\":").Append(JsonString(target.ContentHash));
                    if (target.VisibleText is not null)
                        sb.Append(",\"visibleText\":").Append(JsonString(target.VisibleText));
                    sb.Append('}');
                }
                sb.Append('}');
            }
            sb.Append('}');
        }
        sb.Append(",\"created\":"); AppendAnchorArray(sb, r.Created);
        sb.Append(",\"removed\":"); AppendAnchorArray(sb, r.Removed);
        sb.Append(",\"modified\":"); AppendAnchorArray(sb, r.Modified);
        if (r.TableAnchors is not null)
        {
            sb.Append(",\"tableAnchors\":");
            AppendTableAnchorMapping(sb, r.TableAnchors);
        }
        if (r.AnnotationId is not null)
            sb.Append(",\"annotationId\":").Append(JsonString(r.AnnotationId));
        if (r.HyperlinkId is not null)
            sb.Append(",\"hyperlinkId\":").Append(JsonString(r.HyperlinkId));
        if (r.BookmarkName is not null)
            sb.Append(",\"bookmarkName\":").Append(JsonString(r.BookmarkName));
        if (r.ImageId is not null)
            sb.Append(",\"imageId\":").Append(JsonString(r.ImageId));
        if (r.Patch is not null)
        {
            sb.Append(",\"patch\":{")
              .Append("\"scopeAnchorId\":").Append(JsonString(r.Patch.ScopeAnchorId))
              .Append(",\"markdown\":").Append(JsonString(r.Patch.Markdown))
              .Append('}');
        }
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeTableMetadataResult(TableMetadataResult result)
    {
        var sb = new StringBuilder(1024);
        sb.Append("{\"success\":").Append(result.Success ? "true" : "false");
        if (result.Error is not null) { sb.Append(",\"error\":"); AppendEditError(sb, result.Error); }
        if (result.Metadata is not null) { sb.Append(",\"metadata\":"); AppendTableMetadata(sb, result.Metadata); }
        return sb.Append('}').ToString();
    }

    public static string SerializeTableCellResolutionResult(TableCellResolutionResult result)
    {
        var sb = new StringBuilder(512);
        sb.Append("{\"success\":").Append(result.Success ? "true" : "false");
        if (result.Error is not null) { sb.Append(",\"error\":"); AppendEditError(sb, result.Error); }
        if (result.Cell is not null) { sb.Append(",\"cell\":"); AppendTableCellMetadata(sb, result.Cell); }
        return sb.Append('}').ToString();
    }

    private static void AppendEditError(StringBuilder sb, EditError error)
    {
        sb.Append("{\"code\":\"").Append(EnumToSnake(error.Code)).Append('"')
          .Append(",\"message\":").Append(JsonString(error.Message));
        if (error.AnchorId is not null) sb.Append(",\"anchorId\":").Append(JsonString(error.AnchorId));
        sb.Append('}');
    }

    private static void AppendTableMetadata(StringBuilder sb, TableMetadata metadata)
    {
        sb.Append("{\"anchor\":"); AppendAnchorValue(sb, metadata.Anchor);
        sb.Append(",\"columns\":[");
        for (int i = 0; i < metadata.Columns.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var column = metadata.Columns[i];
            sb.Append("{\"anchor\":"); AppendAnchorValue(sb, column.Anchor);
            sb.Append(",\"tableAnchorId\":").Append(JsonString(column.TableAnchorId))
              .Append(",\"columnIndex\":").Append(column.ColumnIndex)
              .Append(",\"widthTwips\":").Append(column.WidthTwips)
              .Append(",\"isVirtual\":").Append(column.IsVirtual ? "true" : "false")
              .Append(",\"cellAnchorIds\":"); AppendStringArray(sb, column.CellAnchorIds);
            sb.Append('}');
        }
        sb.Append("],\"rows\":[");
        for (int i = 0; i < metadata.Rows.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var row = metadata.Rows[i];
            sb.Append("{\"anchor\":"); AppendAnchorValue(sb, row.Anchor);
            sb.Append(",\"tableAnchorId\":").Append(JsonString(row.TableAnchorId))
              .Append(",\"rowIndex\":").Append(row.RowIndex)
              .Append(",\"gridBefore\":").Append(row.GridBefore)
              .Append(",\"gridAfter\":").Append(row.GridAfter)
              .Append(",\"cells\":[");
            for (int c = 0; c < row.Cells.Count; c++)
            {
                if (c > 0) sb.Append(',');
                AppendTableCellMetadata(sb, row.Cells[c]);
            }
            sb.Append("]}");
        }
        sb.Append("]}");
    }

    private static void AppendTableCellMetadata(StringBuilder sb, TableCellMetadata cell)
    {
        sb.Append("{\"anchor\":"); AppendAnchorValue(sb, cell.Anchor);
        sb.Append(",\"tableAnchorId\":").Append(JsonString(cell.TableAnchorId))
          .Append(",\"rowAnchorId\":").Append(JsonString(cell.RowAnchorId))
          .Append(",\"rowIndex\":").Append(cell.RowIndex)
          .Append(",\"columnIndex\":").Append(cell.ColumnIndex)
          .Append(",\"rowSpan\":").Append(cell.RowSpan)
          .Append(",\"columnSpan\":").Append(cell.ColumnSpan)
          .Append(",\"verticalMerge\":").Append(JsonString(cell.VerticalMerge.ToString().ToLowerInvariant()))
          .Append(",\"paragraphAnchors\":"); AppendAnchorArray(sb, cell.ParagraphAnchors);
        sb.Append('}');
    }

    private static void AppendTableAnchorMapping(StringBuilder sb, TableAnchorMapping mapping)
    {
        sb.Append("{\"retained\":[");
        for (int i = 0; i < mapping.Retained.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append("{\"before\":"); AppendTableAnchorLocation(sb, mapping.Retained[i].Before);
            sb.Append(",\"after\":"); AppendTableAnchorLocation(sb, mapping.Retained[i].After);
            sb.Append('}');
        }
        sb.Append("],\"added\":[");
        for (int i = 0; i < mapping.Added.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendTableAnchorLocation(sb, mapping.Added[i]);
        }
        sb.Append("],\"invalidated\":[");
        for (int i = 0; i < mapping.Invalidated.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendTableAnchorLocation(sb, mapping.Invalidated[i]);
        }
        sb.Append("]}");
    }

    private static void AppendTableAnchorLocation(StringBuilder sb, TableAnchorLocation location)
    {
        sb.Append("{\"anchor\":"); AppendAnchorValue(sb, location.Anchor);
        sb.Append(",\"entityKind\":").Append(JsonString(location.EntityKind.ToString().ToLowerInvariant()));
        if (location.RowIndex is not null) sb.Append(",\"rowIndex\":").Append(location.RowIndex.Value);
        if (location.ColumnIndex is not null) sb.Append(",\"columnIndex\":").Append(location.ColumnIndex.Value);
        if (location.RowSpan is not null) sb.Append(",\"rowSpan\":").Append(location.RowSpan.Value);
        if (location.ColumnSpan is not null) sb.Append(",\"columnSpan\":").Append(location.ColumnSpan.Value);
        if (location.IsVirtual) sb.Append(",\"isVirtual\":true");
        sb.Append('}');
    }

    private static void AppendAnchorValue(StringBuilder sb, Anchor anchor) =>
        sb.Append("{\"id\":").Append(JsonString(anchor.Id))
          .Append(",\"kind\":").Append(JsonString(anchor.Kind))
          .Append(",\"scope\":").Append(JsonString(anchor.Scope))
          .Append(",\"unid\":").Append(JsonString(anchor.Unid))
          .Append('}');
    private static void AppendJsonValue(StringBuilder sb, object? value)
    {
        switch (value)
        {
            case null: sb.Append("null"); break;
            case string s: sb.Append(JsonString(s)); break;
            case bool b: sb.Append(b ? "true" : "false"); break;
            case int i: sb.Append(i); break;
            case long l: sb.Append(l); break;
            default: sb.Append(JsonSerializer.Serialize(value)); break;
        }
    }

    public static string SerializeVersion(long version) =>
        "{\"version\":" + version.ToString(System.Globalization.CultureInfo.InvariantCulture) + "}";

    public static string SerializePageMapRegistration(PageMapRegistrationResult result)
    {
        var sb = new StringBuilder("{\"success\":").Append(result.Success ? "true" : "false");
        if (result.Error is { } error)
            sb.Append(",\"error\":\"").Append(EnumToSnake(error)).Append('"');
        if (result.Message is { } message)
            sb.Append(",\"message\":").Append(JsonString(message));
        return sb.Append('}').ToString();
    }

    public static string SerializePageMapStatus(PageMapStatus status)
    {
        var sb = new StringBuilder("{\"availability\":")
            .Append(JsonString(PageMapAvailabilityString(status.Availability)))
            .Append(",\"documentVersion\":").Append(status.DocumentVersion);
        if (status.UnavailableReason is { } reason)
            sb.Append(",\"unavailableReason\":").Append(JsonString(EnumToSnake(reason)));
        if (status.RendererFingerprint is { } fingerprint)
            sb.Append(",\"rendererFingerprint\":").Append(JsonString(fingerprint));
        if (status.Mode is { } mode)
            sb.Append(",\"mode\":").Append(JsonString(mode == PageMapMode.Paginated ? "paginated" : "continuous"));
        return sb.Append('}').ToString();
    }

    public static string SerializePageCitation(PageCitation citation)
    {
        var sb = new StringBuilder(256);
        AppendPageCitation(sb, citation);
        return sb.ToString();
    }

    private static void AppendPageCitation(StringBuilder sb, PageCitation citation)
    {
        sb.Append("{\"anchorId\":").Append(JsonString(citation.AnchorId))
          .Append(",\"availability\":").Append(JsonString(PageMapAvailabilityString(citation.Availability)))
          .Append(",\"documentVersion\":").Append(citation.DocumentVersion)
          .Append(",\"rendererFingerprint\":").Append(JsonString(citation.RendererFingerprint));
        if (citation.UnavailableReason is { } reason)
            sb.Append(",\"unavailableReason\":").Append(JsonString(EnumToSnake(reason)));
        sb.Append(",\"pages\":[");
        for (int i = 0; i < citation.Pages.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendPageMapPage(sb, citation.Pages[i]);
        }
        sb.Append("],\"fragments\":[");
        for (int i = 0; i < citation.Fragments.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendPageMapFragment(sb, citation.Fragments[i]);
        }
        sb.Append("]}");
    }

    private static void AppendPageMapPage(StringBuilder sb, PageMapPage page)
    {
        sb.Append("{\"pageNumber\":").Append(page.PageNumber)
          .Append(",\"pageInSection\":").Append(page.PageInSection)
          .Append(",\"width\":").Append(Invariant(page.Width))
          .Append(",\"height\":").Append(Invariant(page.Height));
        if (page.SectionIndex is { } sectionIndex)
            sb.Append(",\"sectionIndex\":").Append(sectionIndex);
        sb.Append(",\"pageName\":").Append(JsonString(page.PageName)).Append('}');
    }

    private static void AppendPageMapFragment(StringBuilder sb, PageMapFragment fragment)
    {
        sb.Append("{\"fragmentId\":").Append(JsonString(fragment.FragmentId))
          .Append(",\"anchorId\":").Append(JsonString(fragment.AnchorId))
          .Append(",\"fragmentIndex\":").Append(fragment.FragmentIndex)
          .Append(",\"pageNumber\":").Append(fragment.PageNumber)
          .Append(",\"geometry\":{\"x\":").Append(Invariant(fragment.Geometry.X))
          .Append(",\"y\":").Append(Invariant(fragment.Geometry.Y))
          .Append(",\"width\":").Append(Invariant(fragment.Geometry.Width))
          .Append(",\"height\":").Append(Invariant(fragment.Geometry.Height)).Append('}')
          .Append(",\"story\":").Append(JsonString(PageMapStoryString(fragment.Story)))
          .Append(",\"inTableCell\":").Append(fragment.InTableCell ? "true" : "false")
          .Append('}');
    }

    private static string PageMapAvailabilityString(PageMapAvailability availability) =>
        availability == PageMapAvailability.Available ? "available" : "unavailable";

    private static string PageMapStoryString(PageMapStory story) => story switch
    {
        PageMapStory.Header => "header",
        PageMapStory.Footer => "footer",
        PageMapStory.Footnote => "footnote",
        PageMapStory.Endnote => "endnote",
        PageMapStory.Comment => "comment",
        _ => "body",
    };

    private static string Invariant(double value) =>
        value.ToString("R", System.Globalization.CultureInfo.InvariantCulture);

    public static string SerializeHyperlinks(IReadOnlyList<HyperlinkInfo> links)
    {
        var sb = new StringBuilder(links.Count * 220 + 2).Append('[');
        for (int i = 0; i < links.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var link = links[i];
            sb.Append("{\"id\":").Append(JsonString(link.Id))
              .Append(",\"kind\":").Append(JsonString(link.Kind == HyperlinkKind.Internal ? "internal" : "external"))
              .Append(",\"owningPartUri\":").Append(JsonString(link.OwningPartUri))
              .Append(",\"scope\":").Append(JsonString(link.Scope))
              .Append(",\"anchorId\":").Append(JsonString(link.AnchorId))
              .Append(",\"span\":{\"start\":").Append(link.Span.Start)
              .Append(",\"length\":").Append(link.Span.Length).Append('}')
              .Append(",\"text\":").Append(JsonString(link.Text));
            if (link.Target is not null) sb.Append(",\"target\":").Append(JsonString(link.Target));
            if (link.RelationshipId is not null) sb.Append(",\"relationshipId\":").Append(JsonString(link.RelationshipId));
            if (link.RelationshipIsExternal is not null)
                sb.Append(",\"relationshipIsExternal\":").Append(link.RelationshipIsExternal.Value ? "true" : "false");
            sb.Append(",\"isBroken\":").Append(link.IsBroken ? "true" : "false").Append('}');
        }
        return sb.Append(']').ToString();
    }

    public static string SerializeImages(IReadOnlyList<ImageOccurrence> images)
    {
        var sb = new StringBuilder(images.Count * 700 + 2).Append('[');
        for (int i = 0; i < images.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var image = images[i];
            sb.Append("{\"id\":").Append(JsonString(image.Id))
              .Append(",\"markupKind\":").Append(JsonString(ToSnake(image.MarkupKind.ToString())));
            AppendEnum(sb, "placement", image.Placement);
            sb.Append(",\"canMutate\":").Append(image.CanMutate ? "true" : "false");
            AppendString(sb, "unsupportedReason", image.UnsupportedReason);
            sb.Append(",\"owningPartUri\":").Append(JsonString(image.OwningPartUri))
              .Append(",\"scope\":").Append(JsonString(image.Scope))
              .Append(",\"anchorId\":").Append(JsonString(image.AnchorId))
              .Append(",\"span\":{\"start\":").Append(image.Span.Start)
              .Append(",\"length\":").Append(image.Span.Length).Append('}');
            AppendString(sb, "relationshipId", image.RelationshipId);
            AppendString(sb, "targetPartUri", image.TargetPartUri);
            AppendString(sb, "linkedRelationshipId", image.LinkedRelationshipId);
            AppendString(sb, "linkedTarget", image.LinkedTarget);
            sb.Append(",\"isEmbedded\":").Append(image.IsEmbedded ? "true" : "false")
              .Append(",\"isLinked\":").Append(image.IsLinked ? "true" : "false")
              .Append(",\"isBroken\":").Append(image.IsBroken ? "true" : "false");
            AppendString(sb, "mediaFileName", image.MediaFileName);
            AppendString(sb, "contentType", image.ContentType);
            sb.Append(",\"format\":").Append(JsonString(ToSnake(image.Format.ToString())));
            AppendNullableBool(sb, "contentTypeMatchesBytes", image.ContentTypeMatchesBytes);
            AppendNullableNumber(sb, "intrinsicWidthPixels", image.IntrinsicWidthPixels);
            AppendNullableNumber(sb, "intrinsicHeightPixels", image.IntrinsicHeightPixels);
            AppendNullableDouble(sb, "renderedWidthPoints", image.RenderedWidthPoints);
            AppendNullableDouble(sb, "renderedHeightPoints", image.RenderedHeightPoints);
            AppendString(sb, "altText", image.AltText);
            AppendString(sb, "title", image.Title);
            if (image.FloatingLayout is not null)
            {
                sb.Append(",\"floatingLayout\":");
                AppendFloatingLayout(sb, image.FloatingLayout);
            }
            sb.Append(",\"floatingLayoutSupported\":")
              .Append(image.FloatingLayoutSupported ? "true" : "false").Append('}');
        }
        return sb.Append(']').ToString();
    }

    public static string SerializeImageCapabilities(ImageCapabilities capabilities)
    {
        var sb = new StringBuilder(1200).Append("{\"schemaVersion\":")
            .Append(capabilities.SchemaVersion).Append(",\"runtime\":")
            .Append(JsonString(capabilities.Runtime)).Append(",\"formats\":[");
        for (int i = 0; i < capabilities.Formats.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var format = capabilities.Formats[i];
            sb.Append("{\"format\":").Append(JsonString(ToSnake(format.Format.ToString())))
              .Append(",\"contentType\":").Append(JsonString(format.ContentType))
              .Append(",\"canInspect\":").Append(format.CanInspect ? "true" : "false")
              .Append(",\"canInsert\":").Append(format.CanInsert ? "true" : "false")
              .Append(",\"canReplace\":").Append(format.CanReplace ? "true" : "false");
            AppendString(sb, "limitation", format.Limitation);
            sb.Append('}');
        }
        sb.Append("],\"operations\":"); AppendStringArray(sb, capabilities.Operations);
        sb.Append(",\"mutableWrapModes\":");
        AppendEnumArray(sb, capabilities.MutableWrapModes);
        sb.Append(",\"horizontalReferences\":");
        AppendEnumArray(sb, capabilities.HorizontalReferences.Where(value => value != ImageHorizontalReference.Unknown).ToArray());
        sb.Append(",\"verticalReferences\":");
        AppendEnumArray(sb, capabilities.VerticalReferences.Where(value => value != ImageVerticalReference.Unknown).ToArray());
        sb.Append(",\"maxInputBytes\":").Append(capabilities.MaxInputBytes)
          .Append(",\"maxRenderedPoints\":").Append(capabilities.MaxRenderedPoints.ToString(System.Globalization.CultureInfo.InvariantCulture))
          .Append(",\"defaultDpi\":").Append(capabilities.DefaultDpi.ToString(System.Globalization.CultureInfo.InvariantCulture))
          .Append(",\"usesHeaderParsingOnly\":").Append(capabilities.UsesHeaderParsingOnly ? "true" : "false")
          .Append(",\"acceptsBinaryBytes\":").Append(capabilities.AcceptsBinaryBytes ? "true" : "false")
          .Append(",\"supportsNetworkFetch\":").Append(capabilities.SupportsNetworkFetch ? "true" : "false")
          .Append(",\"supportsFileIo\":").Append(capabilities.SupportsFileIo ? "true" : "false")
          .Append('}');
        return sb.ToString();
    }

    private static void AppendFloatingLayout(StringBuilder sb, FloatingImageLayout layout)
    {
        sb.Append("{\"horizontalRelativeFrom\":").Append(JsonString(ToSnake(layout.HorizontalRelativeFrom.ToString())));
        AppendNullableNumber(sb, "horizontalOffsetEmu", layout.HorizontalOffsetEmu);
        AppendEnum(sb, "horizontalAlignment", layout.HorizontalAlignment);
        sb.Append(",\"verticalRelativeFrom\":").Append(JsonString(ToSnake(layout.VerticalRelativeFrom.ToString())));
        AppendNullableNumber(sb, "verticalOffsetEmu", layout.VerticalOffsetEmu);
        AppendEnum(sb, "verticalAlignment", layout.VerticalAlignment);
        sb.Append(",\"wrapMode\":").Append(JsonString(ToSnake(layout.WrapMode.ToString())))
          .Append(",\"wrapSide\":").Append(JsonString(ToSnake(layout.WrapSide.ToString())))
          .Append(",\"distanceTopEmu\":").Append(layout.DistanceTopEmu)
          .Append(",\"distanceBottomEmu\":").Append(layout.DistanceBottomEmu)
          .Append(",\"distanceLeftEmu\":").Append(layout.DistanceLeftEmu)
          .Append(",\"distanceRightEmu\":").Append(layout.DistanceRightEmu)
          .Append(",\"relativeHeight\":").Append(layout.RelativeHeight)
          .Append(",\"behindDocument\":").Append(layout.BehindDocument ? "true" : "false")
          .Append(",\"locked\":").Append(layout.Locked ? "true" : "false")
          .Append(",\"layoutInCell\":").Append(layout.LayoutInCell ? "true" : "false")
          .Append(",\"allowOverlap\":").Append(layout.AllowOverlap ? "true" : "false");
        AppendString(sb, "rawHorizontalReference", layout.RawHorizontalReference);
        AppendString(sb, "rawVerticalReference", layout.RawVerticalReference);
        AppendString(sb, "rawHorizontalPosition", layout.RawHorizontalPosition);
        AppendString(sb, "rawVerticalPosition", layout.RawVerticalPosition);
        AppendString(sb, "rawWrapMode", layout.RawWrapMode);
        AppendString(sb, "rawWrapSide", layout.RawWrapSide);
        AppendString(sb, "rawRelativeSizeHorizontal", layout.RawRelativeSizeHorizontal);
        AppendString(sb, "rawRelativeSizeVertical", layout.RawRelativeSizeVertical);
        if (layout.RawFlagTokens is not null)
        {
            sb.Append(",\"rawFlagTokens\":{");
            int i = 0;
            foreach (var pair in layout.RawFlagTokens)
            {
                if (i++ > 0) sb.Append(',');
                sb.Append(JsonString(pair.Key)).Append(':').Append(JsonString(pair.Value));
            }
            sb.Append('}');
        }
        sb.Append('}');
    }

    private static string ToSnake(string value)
    {
        var sb = new StringBuilder(value.Length + 4);
        for (int i = 0; i < value.Length; i++)
        {
            if (i > 0 && char.IsUpper(value[i])) sb.Append('_');
            sb.Append(char.ToLowerInvariant(value[i]));
        }
        return sb.ToString();
    }

    private static void AppendString(StringBuilder sb, string name, string? value)
    { if (value is not null) sb.Append(',').Append(JsonString(name)).Append(':').Append(JsonString(value)); }
    private static void AppendNullableNumber<T>(StringBuilder sb, string name, T? value) where T : struct
    { if (value is not null) sb.Append(',').Append(JsonString(name)).Append(':').Append(value.Value); }
    private static void AppendNullableDouble(StringBuilder sb, string name, double? value)
    { if (value is not null) sb.Append(',').Append(JsonString(name)).Append(':').Append(value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)); }
    private static void AppendEnum<T>(StringBuilder sb, string name, T? value) where T : struct, System.Enum
    { if (value is not null) sb.Append(',').Append(JsonString(name)).Append(':').Append(JsonString(ToSnake(value.Value.ToString()))); }
    private static void AppendEnumArray<T>(StringBuilder sb, IReadOnlyList<T> values) where T : struct, System.Enum
    { sb.Append('['); for (int i = 0; i < values.Count; i++) { if (i > 0) sb.Append(','); sb.Append(JsonString(ToSnake(values[i].ToString()))); } sb.Append(']'); }

    public static string SerializeBookmarks(IReadOnlyList<BookmarkInfo> bookmarks)
    {
        var sb = new StringBuilder(bookmarks.Count * 320 + 2).Append('[');
        for (int i = 0; i < bookmarks.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var bookmark = bookmarks[i];
            sb.Append("{\"name\":").Append(JsonString(bookmark.Name))
              .Append(",\"bookmarkId\":").Append(JsonString(bookmark.BookmarkId))
              .Append(",\"startPartUri\":").Append(JsonString(bookmark.StartPartUri))
              .Append(",\"startScope\":").Append(JsonString(bookmark.StartScope));
            if (bookmark.EndPartUri is not null) sb.Append(",\"endPartUri\":").Append(JsonString(bookmark.EndPartUri));
            if (bookmark.EndScope is not null) sb.Append(",\"endScope\":").Append(JsonString(bookmark.EndScope));
            if (bookmark.Range is { } range)
            {
                sb.Append(",\"range\":{\"startAnchorId\":").Append(JsonString(range.StartAnchorId))
                  .Append(",\"startOffset\":").Append(range.StartOffset)
                  .Append(",\"endAnchorId\":").Append(JsonString(range.EndAnchorId))
                  .Append(",\"endOffset\":").Append(range.EndOffset).Append('}');
            }
            sb.Append(",\"segments\":[");
            for (int s = 0; s < bookmark.Segments.Count; s++)
            {
                if (s > 0) sb.Append(',');
                var segment = bookmark.Segments[s];
                sb.Append("{\"owningPartUri\":").Append(JsonString(segment.OwningPartUri))
                  .Append(",\"scope\":").Append(JsonString(segment.Scope))
                  .Append(",\"anchorId\":").Append(JsonString(segment.AnchorId))
                  .Append(",\"span\":{\"start\":").Append(segment.Span.Start)
                  .Append(",\"length\":").Append(segment.Span.Length).Append('}')
                  .Append(",\"text\":").Append(JsonString(segment.Text)).Append('}');
            }
            sb.Append(']').Append(",\"text\":").Append(JsonString(bookmark.Text))
              .Append(",\"isPaired\":").Append(bookmark.IsPaired ? "true" : "false")
              .Append(",\"isManaged\":").Append(bookmark.IsManaged ? "true" : "false")
              .Append(",\"isValid\":").Append(bookmark.IsValid ? "true" : "false");
            if (bookmark.ValidationError is not null)
                sb.Append(",\"validationError\":").Append(JsonString(bookmark.ValidationError));
            sb.Append('}');
        }
        return sb.Append(']').ToString();
    }

    public static string SerializeEditResults(IReadOnlyList<EditResult> results)
    {
        var sb = new StringBuilder(256);
        sb.Append('[');
        for (int i = 0; i < results.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(Serialize(results[i]));
        }
        sb.Append(']');
        return sb.ToString();
    }

    /// <summary>Parse one standard EditResult envelope or an array of them for batch adapters.</summary>
    public static IReadOnlyList<EditResult> DeserializeEditResults(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.ValueKind == JsonValueKind.Array
            ? doc.RootElement.EnumerateArray().Select(ParseEditResult).ToArray()
            : new[] { ParseEditResult(doc.RootElement) };
    }

    private static EditResult ParseEditResult(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object
            || !root.TryGetProperty("success", out var success)
            || success.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            return EditResult.Fail(EditErrorCode.InternalError, "batch step returned a non-EditResult payload");

        EditError? error = null;
        if (root.TryGetProperty("error", out var e) && e.ValueKind == JsonValueKind.Object)
        {
            var codeText = TryGetString(e, "code", "internal_error") ?? "internal_error";
            var code = Enum.GetValues<EditErrorCode>()
                .Where(c => string.Equals(EnumToSnake(c), codeText, StringComparison.Ordinal))
                .Cast<EditErrorCode?>()
                .FirstOrDefault() ?? EditErrorCode.InternalError;
            error = new EditError(
                code,
                TryGetString(e, "message", "batch step failed") ?? "batch step failed",
                TryGetString(e, "anchorId", null));
            if (e.TryGetProperty("precondition", out var p) && p.ValueKind == JsonValueKind.Object)
            {
                PreconditionTarget? target = null;
                if (p.TryGetProperty("currentTarget", out var t) && t.ValueKind == JsonValueKind.Object)
                {
                    target = new PreconditionTarget
                    {
                        Exists = t.TryGetProperty("exists", out var exists) && exists.ValueKind == JsonValueKind.True,
                        AnchorId = TryGetString(t, "anchorId", null),
                        Kind = TryGetString(t, "kind", null),
                        Scope = TryGetString(t, "scope", null),
                        ContentHash = TryGetString(t, "contentHash", null),
                        VisibleText = TryGetString(t, "visibleText", null),
                    };
                }
                error = error with
                {
                    Precondition = new PreconditionFailure(
                        TryGetString(p, "condition", "unknown") ?? "unknown",
                        p.TryGetProperty("expected", out var expected) ? expected.Clone() : null,
                        p.TryGetProperty("actual", out var actual) ? actual.Clone() : null,
                        p.TryGetProperty("currentVersion", out var version) && version.ValueKind == JsonValueKind.Number
                            ? version.GetInt64() : 0,
                        target),
                };
            }
        }

        static IReadOnlyList<Anchor> Anchors(JsonElement root, string name)
        {
            if (!root.TryGetProperty(name, out var a) || a.ValueKind != JsonValueKind.Array)
                return Array.Empty<Anchor>();
            return a.EnumerateArray()
                .Where(x => x.ValueKind == JsonValueKind.Object)
                .Select(x => new Anchor(
                    TryGetString(x, "id", "") ?? "",
                    TryGetString(x, "kind", "") ?? "",
                    TryGetString(x, "scope", "") ?? "",
                    TryGetString(x, "unid", "") ?? ""))
                .ToArray();
        }

        MarkdownPatch? patch = null;
        if (root.TryGetProperty("patch", out var pch) && pch.ValueKind == JsonValueKind.Object)
            patch = new MarkdownPatch(
                TryGetString(pch, "scopeAnchorId", "") ?? "",
                TryGetString(pch, "markdown", "") ?? "");

        return new EditResult
        {
            Success = success.GetBoolean(),
            Error = error,
            Created = Anchors(root, "created"),
            Removed = Anchors(root, "removed"),
            Modified = Anchors(root, "modified"),
            TableAnchors = root.TryGetProperty("tableAnchors", out var tableAnchors)
                && tableAnchors.ValueKind == JsonValueKind.Object
                ? ParseTableAnchorMapping(tableAnchors)
                : null,
            AnnotationId = TryGetString(root, "annotationId", null),
            Patch = patch,
        };
    }

    /// <summary>
    /// Inverse of <see cref="AppendTableAnchorMapping"/>. Structural table ops report the
    /// before/after identity of every row, column and cell they touched; a batch adapter that
    /// dropped this left an agent unable to address the cells its own step had just created.
    /// </summary>
    private static TableAnchorMapping ParseTableAnchorMapping(JsonElement root) => new()
    {
        Retained = root.TryGetProperty("retained", out var retained)
            && retained.ValueKind == JsonValueKind.Array
            ? retained.EnumerateArray()
                .Where(x => x.ValueKind == JsonValueKind.Object)
                .Select(x => new RetainedTableAnchor(
                    ParseTableAnchorLocation(x, "before"),
                    ParseTableAnchorLocation(x, "after")))
                .ToArray()
            : Array.Empty<RetainedTableAnchor>(),
        Added = ParseTableAnchorLocations(root, "added"),
        Invalidated = ParseTableAnchorLocations(root, "invalidated"),
    };

    private static IReadOnlyList<TableAnchorLocation> ParseTableAnchorLocations(
        JsonElement root, string name) =>
        root.TryGetProperty(name, out var locations) && locations.ValueKind == JsonValueKind.Array
            ? locations.EnumerateArray()
                .Where(x => x.ValueKind == JsonValueKind.Object)
                .Select(ParseTableAnchorLocation)
                .ToArray()
            : Array.Empty<TableAnchorLocation>();

    private static TableAnchorLocation ParseTableAnchorLocation(JsonElement root, string name) =>
        root.TryGetProperty(name, out var location) && location.ValueKind == JsonValueKind.Object
            ? ParseTableAnchorLocation(location)
            : ParseTableAnchorLocation(default);

    private static TableAnchorLocation ParseTableAnchorLocation(JsonElement root) => new()
    {
        Anchor = ParseAnchorValue(root, "anchor"),
        // The serializer writes the member name lowercased, NOT snake_case — do not route this
        // through EnumToSnake, which would mis-map any future multi-word member.
        EntityKind = Enum.TryParse<TableAnchorEntityKind>(
            TryGetString(root, "entityKind", null), ignoreCase: true, out var entityKind)
            ? entityKind
            : TableAnchorEntityKind.Table,
        // Coordinates and spans are omitted when null, and isVirtual when false.
        RowIndex = TryGetIntNullable(root, "rowIndex"),
        ColumnIndex = TryGetIntNullable(root, "columnIndex"),
        RowSpan = TryGetIntNullable(root, "rowSpan"),
        ColumnSpan = TryGetIntNullable(root, "columnSpan"),
        IsVirtual = TryGetBoolNullable(root, "isVirtual") ?? false,
    };

    private static Anchor ParseAnchorValue(JsonElement root, string name) =>
        root.ValueKind == JsonValueKind.Object
        && root.TryGetProperty(name, out var anchor)
        && anchor.ValueKind == JsonValueKind.Object
            ? new Anchor(
                TryGetString(anchor, "id", "") ?? "",
                TryGetString(anchor, "kind", "") ?? "",
                TryGetString(anchor, "scope", "") ?? "",
                TryGetString(anchor, "unid", "") ?? "")
            : new Anchor(string.Empty, string.Empty, string.Empty, string.Empty);

    /// <summary>Common structured wire shape for core and transport mutation batches.</summary>
    public static string SerializeMutationBatchResult(MutationBatchResult result)
    {
        var sb = new StringBuilder(512);
        var mode = result.Mode == MutationBatchMode.Atomic ? "atomic" : "best_effort";
        var status = result.Success ? "ok"
            : result.Mode == MutationBatchMode.BestEffort && result.Steps.Any(s => s.Success)
                ? "partial" : "failed";
        sb.Append("{\"mode\":").Append(JsonString(mode))
          .Append(",\"status\":").Append(JsonString(status))
          .Append(",\"preview\":").Append(result.Preview ? "true" : "false")
          .Append(",\"success\":").Append(result.Success ? "true" : "false")
          .Append(",\"rolledBack\":").Append(result.RolledBack ? "true" : "false")
          .Append(",\"baseVersion\":").Append(result.BaseVersion)
          .Append(",\"resultVersion\":").Append(result.ResultVersion)
          .Append(",\"packageHash\":")
          .Append(result.PackageHash is null ? "null" : JsonString(result.PackageHash))
          .Append(",\"steps\":[");
        for (int i = 0; i < result.Steps.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var step = result.Steps[i];
            sb.Append("{\"index\":").Append(step.Index)
              .Append(",\"tool\":").Append(JsonString(step.Tool))
              .Append(",\"action\":").Append(JsonString(step.Action))
              .Append(",\"success\":").Append(step.Success ? "true" : "false")
              .Append(",\"rolledBack\":").Append(step.RolledBack ? "true" : "false")
              .Append(",\"results\":").Append(SerializeEditResults(step.Results))
              .Append('}');
        }
        sb.Append(']')
          .Append(",\"editsApplied\":").Append(
              result.RolledBack ? 0 : result.Steps.Count(s => s.Success))
          .Append(",\"results\":[");
        for (int i = 0; i < result.Steps.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var stepResults = result.Steps[i].Results;
            sb.Append(stepResults.Count == 1
                ? Serialize(stepResults[0])
                : SerializeEditResults(stepResults));
        }
        sb.Append("],\"errors\":[");
        bool firstError = true;
        foreach (var step in result.Steps.Where(s => !s.Success))
        {
            var failedError = step.Results.FirstOrDefault(r => !r.Success)?.Error;
            if (failedError is null) continue;
            if (!firstError) sb.Append(',');
            firstError = false;
            var failedJson = Serialize(new EditResult { Success = false, Error = failedError });
            using var failedDoc = JsonDocument.Parse(failedJson);
            sb.Append(failedDoc.RootElement.GetProperty("error").GetRawText());
        }
        sb.Append(']');
        if (result.Failure is { } failure)
        {
            var errorJson = Serialize(new EditResult { Success = false, Error = failure.Error });
            using var errorDoc = JsonDocument.Parse(errorJson);
            sb.Append(",\"failure\":{\"index\":").Append(failure.Index)
              .Append(",\"tool\":").Append(JsonString(failure.Tool))
              .Append(",\"action\":").Append(JsonString(failure.Action))
              .Append(",\"error\":").Append(errorDoc.RootElement.GetProperty("error").GetRawText())
              .Append(",\"rolledBack\":").Append(failure.RolledBack ? "true" : "false")
              .Append('}');
        }
        sb.Append(",\"revisionChanges\":");
        AppendChangeSet(sb, result.RevisionChanges, SerializeRevisionList);
        sb.Append(",\"commentChanges\":");
        AppendChangeSet(sb, result.CommentChanges, SerializeCommentList);
        sb.Append(",\"annotationChanges\":");
        AppendChangeSet(sb, result.AnnotationChanges, SerializeAnnotations);
        sb.Append(",\"warnings\":[");
        for (int i = 0; i < result.Warnings.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(JsonString(result.Warnings[i]));
        }
        sb.Append(']')
          .Append(",\"html\":")
          .Append(result.Html is null ? "null" : JsonString(result.Html));
        sb.Append('}');
        return sb.ToString();
    }

    private static void AppendChangeSet<T>(
        StringBuilder sb,
        MutationBatchChangeSet<T> changes,
        System.Func<IReadOnlyList<T>, string> serialize)
    {
        sb.Append("{\"added\":").Append(serialize(changes.Added))
          .Append(",\"removed\":").Append(serialize(changes.Removed))
          .Append(",\"modified\":").Append(serialize(changes.Modified))
          .Append('}');
    }

    public static void AppendAnchorArray(StringBuilder sb, IReadOnlyList<Anchor> anchors)
    {
        sb.Append('[');
        for (int i = 0; i < anchors.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var a = anchors[i];
            sb.Append('{')
              .Append("\"id\":").Append(JsonString(a.Id))
              .Append(",\"kind\":").Append(JsonString(a.Kind))
              .Append(",\"scope\":").Append(JsonString(a.Scope))
              .Append(",\"unid\":").Append(JsonString(a.Unid))
              .Append('}');
        }
        sb.Append(']');
    }

    private static string KindToString(PlaceholderKind kind) => kind switch
    {
        PlaceholderKind.BlankFill => "blank_fill",
        PlaceholderKind.AlternativeClause => "alternative_clause",
        PlaceholderKind.Instruction => "instruction",
        _ => "unknown",
    };

    public static string SerializeEditSummary(EditSummary summary)
    {
        var sb = new StringBuilder(1024);
        sb.Append("{\"totalAnchors\":").Append(summary.TotalAnchors)
          .Append(",\"remainingPlaceholders\":").Append(SerializePlaceholders(summary.RemainingPlaceholders))
          .Append(",\"bareUnderscoreRuns\":").Append(SerializeMatches(summary.BareUnderscoreRuns))
          .Append(",\"footnoteCount\":").Append(summary.FootnoteCount)
          .Append(",\"inlineFootnoteRefCount\":").Append(summary.InlineFootnoteRefCount)
          .Append(",\"commentCount\":").Append(summary.CommentCount)
          .Append('}');
        return sb.ToString();
    }

    public static string SerializePlaceholders(IReadOnlyList<TemplatePlaceholder> placeholders)
    {
        var sb = new StringBuilder(512);
        sb.Append('[');
        for (int i = 0; i < placeholders.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var p = placeholders[i];
            sb.Append("{\"kind\":\"").Append(KindToString(p.Kind)).Append('"');
            sb.Append(",\"alternativeKinds\":[");
            for (int a = 0; a < p.AlternativeKinds.Count; a++)
            {
                if (a > 0) sb.Append(',');
                sb.Append('"').Append(KindToString(p.AlternativeKinds[a])).Append('"');
            }
            sb.Append(']');
            if (p.Hint is not null)
                sb.Append(",\"hint\":").Append(JsonString(p.Hint));
            sb.Append(",\"match\":");
            AppendMatch(sb, p.Match);
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    public static void AppendMatch(StringBuilder sb, TextMatch m)
    {
        sb.Append("{\"text\":").Append(JsonString(m.Text))
          .Append(",\"enclosingAnchor\":{")
          .Append("\"id\":").Append(JsonString(m.EnclosingAnchor.Anchor.Id))
          .Append(",\"kind\":").Append(JsonString(m.EnclosingAnchor.Anchor.Kind))
          .Append(",\"scope\":").Append(JsonString(m.EnclosingAnchor.Anchor.Scope))
          .Append(",\"unid\":").Append(JsonString(m.EnclosingAnchor.Anchor.Unid))
          .Append('}')
          .Append(",\"span\":{\"start\":").Append(m.Span.Start).Append(",\"length\":").Append(m.Span.Length).Append('}')
          .Append(",\"contextBefore\":").Append(JsonString(m.ContextBefore))
          .Append(",\"contextAfter\":").Append(JsonString(m.ContextAfter))
          .Append(",\"fragments\":[");
        for (int f = 0; f < m.Fragments.Count; f++)
        {
            if (f > 0) sb.Append(',');
            var fr = m.Fragments[f];
            sb.Append("{\"unid\":").Append(JsonString(fr.Unid))
              .Append(",\"text\":").Append(JsonString(fr.Text))
              .Append(",\"spanInElement\":{\"start\":").Append(fr.SpanInElement.Start)
              .Append(",\"length\":").Append(fr.SpanInElement.Length).Append('}')
              .Append(",\"formatting\":{")
              .Append("\"bold\":").Append(fr.Formatting.Bold ? "true" : "false")
              .Append(",\"italic\":").Append(fr.Formatting.Italic ? "true" : "false")
              .Append(",\"underline\":").Append(fr.Formatting.Underline ? "true" : "false")
              .Append(",\"strike\":").Append(fr.Formatting.Strike ? "true" : "false")
              .Append(",\"code\":").Append(fr.Formatting.Code ? "true" : "false")
              .Append("}}");
        }
        sb.Append(']');
        // Groups omitted from placeholder serialization (rarely useful for this surface).
        sb.Append('}');
    }

    public static string SerializeMatches(IReadOnlyList<TextMatch> matches)
    {
        var sb = new StringBuilder(512);
        sb.Append('[');
        for (int i = 0; i < matches.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var m = matches[i];
            sb.Append("{\"text\":").Append(JsonString(m.Text))
              .Append(",\"enclosingAnchor\":");
            AppendAnchor(sb, m.EnclosingAnchor);
            sb.Append(",\"span\":{\"start\":").Append(m.Span.Start).Append(",\"length\":").Append(m.Span.Length).Append('}')
              .Append(",\"contextBefore\":").Append(JsonString(m.ContextBefore))
              .Append(",\"contextAfter\":").Append(JsonString(m.ContextAfter))
              .Append(",\"groups\":");
            AppendStringArray(sb, m.Groups);
            sb.Append(",\"fragments\":");
            AppendFragments(sb, m.Fragments);
            if (m.Citation is { } citation)
            {
                sb.Append(",\"citation\":");
                AppendPageCitation(sb, citation);
            }
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    public static string SerializeCrossBlockMatches(IReadOnlyList<CrossBlockMatch> matches)
    {
        var sb = new StringBuilder(512);
        sb.Append('[');
        for (int i = 0; i < matches.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var m = matches[i];
            sb.Append("{\"text\":").Append(JsonString(m.Text))
              .Append(",\"enclosingAnchors\":[");
            for (int a = 0; a < m.EnclosingAnchors.Count; a++)
            {
                if (a > 0) sb.Append(',');
                AppendAnchor(sb, m.EnclosingAnchors[a]);
            }
            sb.Append(']')
              .Append(",\"slices\":[");
            for (int sIdx = 0; sIdx < m.Slices.Count; sIdx++)
            {
                if (sIdx > 0) sb.Append(',');
                var slice = m.Slices[sIdx];
                sb.Append("{\"anchor\":");
                AppendAnchor(sb, slice.Anchor);
                sb.Append(",\"spanInBlock\":{\"start\":").Append(slice.SpanInBlock.Start)
                  .Append(",\"length\":").Append(slice.SpanInBlock.Length).Append('}')
                  .Append(",\"fragments\":");
                AppendFragments(sb, slice.Fragments);
                sb.Append('}');
            }
            sb.Append(']')
              .Append(",\"contextBefore\":").Append(JsonString(m.ContextBefore))
              .Append(",\"contextAfter\":").Append(JsonString(m.ContextAfter))
              .Append(",\"groups\":");
            AppendStringArray(sb, m.Groups);
            if (m.Citations is { } citations)
            {
                sb.Append(",\"citations\":[");
                for (int c = 0; c < citations.Count; c++)
                {
                    if (c > 0) sb.Append(',');
                    AppendPageCitation(sb, citations[c]);
                }
                sb.Append(']');
            }
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    public static void AppendAnchor(StringBuilder sb, AnchorTarget t) =>
        sb.Append("{\"id\":").Append(JsonString(t.Anchor.Id))
          .Append(",\"kind\":").Append(JsonString(t.Anchor.Kind))
          .Append(",\"scope\":").Append(JsonString(t.Anchor.Scope))
          .Append(",\"unid\":").Append(JsonString(t.Anchor.Unid))
          .Append('}');

    public static void AppendStringArray(StringBuilder sb, IReadOnlyList<string> items)
    {
        sb.Append('[');
        for (int i = 0; i < items.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(JsonString(items[i]));
        }
        sb.Append(']');
    }

    public static void AppendFragments(StringBuilder sb, IReadOnlyList<RunFragment> fragments)
    {
        sb.Append('[');
        for (int f = 0; f < fragments.Count; f++)
        {
            if (f > 0) sb.Append(',');
            var fr = fragments[f];
            sb.Append("{\"unid\":").Append(JsonString(fr.Unid))
              .Append(",\"text\":").Append(JsonString(fr.Text))
              .Append(",\"spanInElement\":{\"start\":").Append(fr.SpanInElement.Start)
              .Append(",\"length\":").Append(fr.SpanInElement.Length).Append('}')
              .Append(",\"formatting\":{")
              .Append("\"bold\":").Append(fr.Formatting.Bold ? "true" : "false")
              .Append(",\"italic\":").Append(fr.Formatting.Italic ? "true" : "false")
              .Append(",\"underline\":").Append(fr.Formatting.Underline ? "true" : "false")
              .Append(",\"strike\":").Append(fr.Formatting.Strike ? "true" : "false")
              .Append(",\"code\":").Append(fr.Formatting.Code ? "true" : "false");
            if (fr.Formatting.Color is not null)
                sb.Append(",\"color\":").Append(JsonString(fr.Formatting.Color));
            if (fr.Formatting.HyperlinkUrl is not null)
                sb.Append(",\"hyperlinkUrl\":").Append(JsonString(fr.Formatting.HyperlinkUrl));
            if (fr.Formatting.RunStyle is not null)
                sb.Append(",\"runStyle\":").Append(JsonString(fr.Formatting.RunStyle));
            sb.Append("}}");
        }
        sb.Append(']');
    }

    public static string SerializeProjection(MarkdownProjection p)
    {
        var sb = new StringBuilder(p.Markdown.Length + 200);
        sb.Append("{\"markdown\":").Append(JsonString(p.Markdown));
        sb.Append(",\"anchorIndex\":{");
        bool first = true;
        foreach (var kv in p.AnchorIndex)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(JsonString(kv.Key)).Append(":{")
              .Append("\"partUri\":").Append(JsonString(kv.Value.PartUri))
              .Append(",\"unid\":").Append(JsonString(kv.Value.Unid))
              .Append(",\"kind\":").Append(JsonString(kv.Value.Anchor.Kind))
              .Append(",\"scope\":").Append(JsonString(kv.Value.Anchor.Scope))
              .Append(",\"textPreview\":").Append(JsonString(kv.Value.TextPreview));
            if (kv.Value.AutoNumberPrefix is { } prefix)
                sb.Append(",\"autoNumberPrefix\":").Append(JsonString(prefix));
            sb.Append('}');
        }
        sb.Append('}');
        if (p.PageCitations is { } citations)
        {
            sb.Append(",\"pageCitations\":{");
            bool firstCitation = true;
            foreach (var (anchorId, citation) in citations)
            {
                if (!firstCitation) sb.Append(',');
                firstCitation = false;
                sb.Append(JsonString(anchorId)).Append(':');
                AppendPageCitation(sb, citation);
            }
            sb.Append('}');
        }
        sb.Append('}');
        return sb.ToString();
    }

    public static string JsonString(string s)
    {
        var sb = new StringBuilder(s.Length + 2);
        sb.Append('"');
        foreach (var c in s)
        {
            switch (c)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                default:
                    if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("X4"));
                    else sb.Append(c);
                    break;
            }
        }
        sb.Append('"');
        return sb.ToString();
    }

    public static string EnumToSnake(System.Enum code)
    {
        var s = code.ToString();
        var sb = new StringBuilder(s.Length + 4);
        for (int i = 0; i < s.Length; i++)
        {
            if (i > 0 && char.IsUpper(s[i])) sb.Append('_');
            sb.Append(char.ToLowerInvariant(s[i]));
        }
        return sb.ToString();
    }

    /// <summary>The wire shape for <see cref="DocxSession.ValidMoveTargets"/>:
    /// <c>[{"anchorId":…,"before":true,"after":false}]</c>.</summary>
    public static string SerializeMoveTargets(IReadOnlyList<MoveTarget> targets)
    {
        var sb = new StringBuilder(2 + targets.Count * 72);
        sb.Append('[');
        for (int i = 0; i < targets.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append("{\"anchorId\":").Append(JsonString(targets[i].AnchorId))
              .Append(",\"before\":").Append(targets[i].Before ? "true" : "false")
              .Append(",\"after\":").Append(targets[i].After ? "true" : "false")
              .Append('}');
        }
        return sb.Append(']').ToString();
    }

    /// <summary>
    /// The projection's anchor index WITHOUT the markdown payload — the shape the
    /// editor's anchor-map refresh needs. Emits the same <c>{"anchorIndex":{…}}</c>
    /// object <see cref="SerializeProjection"/> nests, so clients parse identically;
    /// serialized from the cheap index-only entries (previews empty).
    /// </summary>
    public static string SerializeAnchorIndex(IReadOnlyDictionary<string, AnchorTarget> index)
    {
        var sb = new StringBuilder(64 + index.Count * 96);
        sb.Append("{\"anchorIndex\":{");
        bool first = true;
        foreach (var kv in index)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(JsonString(kv.Key)).Append(":{")
              .Append("\"partUri\":").Append(JsonString(kv.Value.PartUri))
              .Append(",\"unid\":").Append(JsonString(kv.Value.Unid))
              .Append(",\"kind\":").Append(JsonString(kv.Value.Anchor.Kind))
              .Append(",\"scope\":").Append(JsonString(kv.Value.Anchor.Scope))
              .Append('}');
        }
        sb.Append("}}");
        return sb.ToString();
    }

    public static string SerializeRenderPlan(RenderPlan plan)
    {
        var sb = new StringBuilder(256);
        void Units(string key, System.Collections.Generic.IReadOnlyList<RenderUnit> units)
        {
            sb.Append('"').Append(key).Append("\":[");
            for (int i = 0; i < units.Count; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append("{\"id\":").Append(JsonString(units[i].Id))
                  .Append(",\"kind\":").Append(JsonString(units[i].Kind));
                if (units[i].Sig is { } sig)
                    sb.Append(",\"sig\":").Append(JsonString(sig));
                sb.Append('}');
            }
            sb.Append(']');
        }
        sb.Append('{');
        Units("body", plan.Body);
        sb.Append(',');
        Units("footnotes", plan.Footnotes);
        sb.Append(',');
        Units("endnotes", plan.Endnotes);
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeNoteList(IReadOnlyList<NoteListEntry> notes)
    {
        var sb = new StringBuilder(64 + notes.Count * 64);
        sb.Append('[');
        for (int i = 0; i < notes.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append("{\"id\":").Append(JsonString(notes[i].Id))
              .Append(",\"defAnchorId\":").Append(JsonString(notes[i].DefAnchorId))
              .Append(",\"ordinal\":").Append(notes[i].Ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture))
              .Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    public static string SerializeCommentList(IReadOnlyList<CommentListEntry> comments)
    {
        var sb = new StringBuilder(64 + comments.Count * 96);
        sb.Append('[');
        for (int i = 0; i < comments.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var c = comments[i];
            sb.Append("{\"anchorId\":").Append(JsonString(c.DefAnchorId))
              .Append(",\"author\":").Append(JsonString(c.Author));
            if (c.Initials is not null) sb.Append(",\"initials\":").Append(JsonString(c.Initials));
            if (c.Date is not null) sb.Append(",\"date\":").Append(JsonString(c.Date));
            sb.Append(",\"text\":").Append(JsonString(c.Text));
            if (c.ParentAnchorId is not null)
                sb.Append(",\"parentAnchorId\":").Append(JsonString(c.ParentAnchorId));
            if (c.Resolved.HasValue)
                sb.Append(",\"resolved\":").Append(c.Resolved.Value ? "true" : "false");
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    /// <summary>Serialize <see cref="DocxSession.ListRevisions"/> output:
    /// <c>[{"id","type","author","date"?,"text","anchorId"?}]</c> (date/anchorId
    /// omitted when null).</summary>
    public static string SerializeRevisionList(IReadOnlyList<RevisionListEntry> revisions)
    {
        var sb = new StringBuilder(64 + revisions.Count * 128);
        sb.Append('[');
        for (int i = 0; i < revisions.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var r = revisions[i];
            sb.Append("{\"id\":").Append(JsonString(r.Id))
              .Append(",\"type\":").Append(JsonString(r.Type))
              .Append(",\"author\":").Append(JsonString(r.Author));
            if (r.Date is not null) sb.Append(",\"date\":").Append(JsonString(r.Date));
            sb.Append(",\"text\":").Append(JsonString(r.Text));
            if (r.AnchorId is not null) sb.Append(",\"anchorId\":").Append(JsonString(r.AnchorId));
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    /// <summary>
    /// Parse an ISO-8601 comment date from the wire; null/empty → null (the deterministic
    /// no-date default). An unparseable string throws <see cref="System.FormatException"/> at
    /// the transport layer — the <see cref="ParseHeaderFooterKind"/> precedent, never a
    /// silent drop.
    /// </summary>
    public static System.DateTime? ParseCommentDate(string? iso) =>
        string.IsNullOrEmpty(iso)
            ? null
            : System.DateTime.Parse(iso, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind);

    public static string SerializeAnchorTargets(IReadOnlyList<AnchorTarget> targets)
    {
        var sb = new StringBuilder(targets.Count * 128 + 2);
        sb.Append('[');
        for (int i = 0; i < targets.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendAnchorTarget(sb, targets[i]);
        }
        sb.Append(']');
        return sb.ToString();
    }

    public static string SerializeAnchorTargetMap(
        IReadOnlyDictionary<string, IReadOnlyList<AnchorTarget>> map)
    {
        var sb = new StringBuilder(256);
        sb.Append('{');
        bool first = true;
        foreach (var kv in map)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(JsonString(kv.Key)).Append(':');
            sb.Append(SerializeAnchorTargets(kv.Value));
        }
        sb.Append('}');
        return sb.ToString();
    }

    public static void AppendAnchorTarget(StringBuilder sb, AnchorTarget t)
    {
        sb.Append("{\"id\":").Append(JsonString(t.Anchor.Id))
          .Append(",\"kind\":").Append(JsonString(t.Anchor.Kind))
          .Append(",\"scope\":").Append(JsonString(t.Anchor.Scope))
          .Append(",\"unid\":").Append(JsonString(t.Unid))
          .Append(",\"partUri\":").Append(JsonString(t.PartUri))
          .Append(",\"textPreview\":").Append(JsonString(t.TextPreview));
        if (t.AutoNumberPrefix is { } prefix)
            sb.Append(",\"autoNumberPrefix\":").Append(JsonString(prefix));
        if (t.Citation is { } citation)
        {
            sb.Append(",\"citation\":");
            AppendPageCitation(sb, citation);
        }
        sb.Append('}');
    }

    public static string SerializeAnchorTargetOrNull(AnchorTarget? target)
    {
        if (target is null) return "null";
        var sb = new StringBuilder(128);
        AppendAnchorTarget(sb, target);
        return sb.ToString();
    }

    public static string SerializeAnchorInfoMap(IReadOnlyDictionary<string, AnchorInfo?> map)
    {
        var sb = new StringBuilder(map.Count * 100 + 2);
        sb.Append('{');
        bool first = true;
        foreach (var kv in map)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(JsonString(kv.Key)).Append(':');
            sb.Append(SerializeAnchorInfoOrNull(kv.Value));
        }
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeAnchorInfoOrNull(AnchorInfo? info)
    {
        if (info is null) return "null";
        var sb = new StringBuilder(128);
        sb.Append("{\"id\":").Append(JsonString(info.Id))
          .Append(",\"kind\":").Append(JsonString(info.Kind))
          .Append(",\"scope\":").Append(JsonString(info.Scope))
          .Append(",\"textPreview\":").Append(JsonString(info.TextPreview))
          .Append(",\"contentHash\":").Append(JsonString(info.ContentHash))
          .Append(",\"visibleText\":").Append(JsonString(info.VisibleText));
        if (info.AutoNumberPrefix is { } prefix)
            sb.Append(",\"autoNumberPrefix\":").Append(JsonString(prefix));
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeBlockMetadataOrNull(BlockMetadata? meta)
    {
        if (meta is null) return "null";
        var sb = new StringBuilder(256);
        sb.Append("{\"anchorId\":").Append(JsonString(meta.AnchorId))
          .Append(",\"kind\":").Append(JsonString(meta.Kind))
          .Append(",\"scope\":").Append(JsonString(meta.Scope));
        if (meta.StyleId is not null)
            sb.Append(",\"styleId\":").Append(JsonString(meta.StyleId));
        if (meta.StyleName is not null)
            sb.Append(",\"styleName\":").Append(JsonString(meta.StyleName));
        if (meta.OutlineLevel.HasValue)
            sb.Append(",\"outlineLevel\":").Append(meta.OutlineLevel.Value);
        if (meta.List is not null)
            sb.Append(",\"list\":").Append(SerializeListMembershipOrNull(meta.List));
        sb.Append(",\"hasInlineFormatting\":").Append(meta.HasInlineFormatting ? "true" : "false");
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeBlockMetadataMap(System.Collections.Generic.IReadOnlyDictionary<string, BlockMetadata?> map)
    {
        var sb = new StringBuilder(map.Count * 200 + 2);
        sb.Append('{');
        bool first = true;
        foreach (var kv in map)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(JsonString(kv.Key)).Append(':');
            sb.Append(SerializeBlockMetadataOrNull(kv.Value));
        }
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeListMembershipOrNull(ListMembership? list)
    {
        if (list is null) return "null";
        var sb = new StringBuilder(256);
        sb.Append("{\"anchorId\":").Append(JsonString(list.AnchorId))
          .Append(",\"numId\":").Append(list.NumId)
          .Append(",\"abstractNumId\":").Append(list.AbstractNumId)
          .Append(",\"level\":").Append(list.Level)
          .Append(",\"format\":").Append(JsonString(NumberFormatToString(list.Format)))
          .Append(",\"start\":").Append(list.Start)
          .Append(",\"isAutoNumbered\":").Append(list.IsAutoNumbered ? "true" : "false")
          .Append(",\"fromStyle\":").Append(list.FromStyle ? "true" : "false");
        if (list.StartOverride.HasValue)
            sb.Append(",\"startOverride\":").Append(list.StartOverride.Value);
        if (list.LevelText is not null)
            sb.Append(",\"levelText\":").Append(JsonString(list.LevelText));
        if (list.LeftIndentTwips.HasValue)
            sb.Append(",\"leftIndentTwips\":").Append(list.LeftIndentTwips.Value);
        if (list.RightIndentTwips.HasValue)
            sb.Append(",\"rightIndentTwips\":").Append(list.RightIndentTwips.Value);
        if (list.FirstLineIndentTwips.HasValue)
            sb.Append(",\"firstLineIndentTwips\":").Append(list.FirstLineIndentTwips.Value);
        if (list.HangingIndentTwips.HasValue)
            sb.Append(",\"hangingIndentTwips\":").Append(list.HangingIndentTwips.Value);
        if (list.GeneratedLabel is not null)
            sb.Append(",\"generatedLabel\":").Append(JsonString(list.GeneratedLabel));
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeSectionInfoOrNull(SectionInfo? info)
    {
        if (info is null) return "null";
        var sb = new StringBuilder(256);
        sb.Append("{\"anchorId\":").Append(JsonString(info.AnchorId))
          .Append(",\"sectionUnid\":").Append(JsonString(info.SectionUnid))
          .Append(",\"pageWidthTwips\":").Append(info.PageWidthTwips)
          .Append(",\"pageHeightTwips\":").Append(info.PageHeightTwips)
          .Append(",\"landscape\":").Append(info.Landscape ? "true" : "false")
          .Append(",\"marginTopTwips\":").Append(info.MarginTopTwips)
          .Append(",\"marginBottomTwips\":").Append(info.MarginBottomTwips)
          .Append(",\"marginLeftTwips\":").Append(info.MarginLeftTwips)
          .Append(",\"marginRightTwips\":").Append(info.MarginRightTwips)
          .Append(",\"columns\":").Append(info.Columns)
          .Append(",\"headerPartUris\":[");
        for (int i = 0; i < info.HeaderPartUris.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(JsonString(info.HeaderPartUris[i]));
        }
        sb.Append("],\"footerPartUris\":[");
        for (int i = 0; i < info.FooterPartUris.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(JsonString(info.FooterPartUris[i]));
        }
        sb.Append(']');
        AppendHeaderFooterRefs(sb, ",\"headerRefs\":", info.HeaderRefs);
        AppendHeaderFooterRefs(sb, ",\"footerRefs\":", info.FooterRefs);
        // Omitted, not null-valued, when absent — an absent w:pgNumType attribute is "inherit", and
        // the optional TypeScript/Python fields read that as undefined/None.
        if (info.PageNumberStart is { } pnStart)
            sb.Append(",\"pageNumberStart\":").Append(pnStart);
        if (info.PageNumberFormat is { } pnFormat)
            sb.Append(",\"pageNumberFormat\":").Append(JsonString(NumberFormatToString(pnFormat)));
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeStyles(IReadOnlyList<StyleInfo> styles)
    {
        var sb = new StringBuilder(styles.Count * 400 + 2);
        sb.Append('[');
        for (int i = 0; i < styles.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var style = styles[i];
            sb.Append("{\"id\":").Append(JsonString(style.Id))
              .Append(",\"name\":").Append(JsonString(style.Name))
              .Append(",\"type\":").Append(JsonString(style.Type));
            if (style.BasedOn is not null)
                sb.Append(",\"basedOn\":").Append(JsonString(style.BasedOn));
            if (style.Next is not null)
                sb.Append(",\"next\":").Append(JsonString(style.Next));
            sb.Append(",\"isDefault\":").Append(style.IsDefault ? "true" : "false")
              .Append(",\"isCustom\":").Append(style.IsCustom ? "true" : "false")
              .Append(",\"hasLatentException\":").Append(style.HasLatentException ? "true" : "false");
            if (style.UiPriority.HasValue)
                sb.Append(",\"uiPriority\":").Append(style.UiPriority.Value);
            AppendNullableBool(sb, "semiHidden", style.SemiHidden);
            AppendNullableBool(sb, "unhideWhenUsed", style.UnhideWhenUsed);
            AppendNullableBool(sb, "quickFormat", style.QuickFormat);
            AppendNullableBool(sb, "locked", style.Locked);
            if (style.ResolvedParagraph is not null)
            {
                sb.Append(",\"resolvedParagraph\":");
                AppendParagraphFormatting(sb, style.ResolvedParagraph);
            }
            if (style.ResolvedRun is not null)
            {
                sb.Append(",\"resolvedRun\":");
                AppendRunFormattingInfo(sb, style.ResolvedRun);
            }
            if (style.ResolvedTable is not null)
            {
                sb.Append(",\"resolvedTable\":");
                AppendTableStyleFormatting(sb, style.ResolvedTable);
            }
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    public static string SerializeFormattingInspectionOrNull(FormattingInspection? inspection)
    {
        if (inspection is null) return "null";
        var sb = new StringBuilder(512 + inspection.Runs.Count * 300);
        sb.Append("{\"anchorId\":").Append(JsonString(inspection.AnchorId))
          .Append(",\"directParagraph\":");
        AppendParagraphFormatting(sb, inspection.DirectParagraph);
        sb.Append(",\"effectiveParagraph\":");
        AppendParagraphFormatting(sb, inspection.EffectiveParagraph);
        sb.Append(",\"runs\":");
        AppendInlineSpans(sb, inspection.Runs);
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeInlineSpans(IReadOnlyList<InlineSpan> spans)
    {
        var sb = new StringBuilder(spans.Count * 300 + 2);
        AppendInlineSpans(sb, spans);
        return sb.ToString();
    }

    private static void AppendInlineSpans(StringBuilder sb, IReadOnlyList<InlineSpan> spans)
    {
        sb.Append('[');
        for (int i = 0; i < spans.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var span = spans[i];
            sb.Append("{\"anchorId\":").Append(JsonString(span.AnchorId))
              .Append(",\"runUnid\":").Append(JsonString(span.RunUnid))
              .Append(",\"span\":{\"start\":").Append(span.Span.Start)
              .Append(",\"length\":").Append(span.Span.Length).Append('}')
              .Append(",\"text\":").Append(JsonString(span.Text))
              .Append(",\"direct\":");
            AppendRunFormattingInfo(sb, span.Direct);
            sb.Append(",\"effective\":");
            AppendRunFormattingInfo(sb, span.Effective);
            sb.Append('}');
        }
        sb.Append(']');
    }

    private static void AppendParagraphFormatting(StringBuilder sb, ParagraphFormatting f)
    {
        sb.Append('{');
        bool has = false;
        void StringValue(string name, string? value)
        {
            if (value is null) return;
            if (has) sb.Append(',');
            has = true;
            sb.Append(JsonString(name)).Append(':').Append(JsonString(value));
        }
        void IntValue(string name, int? value)
        {
            if (!value.HasValue) return;
            if (has) sb.Append(',');
            has = true;
            sb.Append(JsonString(name)).Append(':').Append(value.Value);
        }
        void BoolValue(string name, bool? value)
        {
            if (!value.HasValue) return;
            if (has) sb.Append(',');
            has = true;
            sb.Append(JsonString(name)).Append(':').Append(value.Value ? "true" : "false");
        }

        StringValue("styleId", f.StyleId);
        StringValue("alignment", f.Alignment switch
        {
            ParagraphAlignment.Center => "center",
            ParagraphAlignment.Right => "right",
            ParagraphAlignment.Justify => "justify",
            ParagraphAlignment.Left => "left",
            _ => null,
        });
        IntValue("leftIndentTwips", f.LeftIndentTwips);
        IntValue("rightIndentTwips", f.RightIndentTwips);
        IntValue("firstLineIndentTwips", f.FirstLineIndentTwips);
        IntValue("hangingIndentTwips", f.HangingIndentTwips);
        IntValue("spacingBeforeTwips", f.SpacingBeforeTwips);
        IntValue("spacingAfterTwips", f.SpacingAfterTwips);
        IntValue("lineSpacing", f.LineSpacing);
        StringValue("lineSpacingRule", f.LineSpacingRule switch
        {
            LineSpacingRule.Exact => "exact",
            LineSpacingRule.AtLeast => "atLeast",
            LineSpacingRule.Auto => "auto",
            _ => null,
        });
        BoolValue("keepNext", f.KeepNext);
        BoolValue("keepLines", f.KeepLines);
        BoolValue("pageBreakBefore", f.PageBreakBefore);
        IntValue("outlineLevel", f.OutlineLevel);
        StringValue("shadingFill", f.ShadingFill);
        AppendBorder("topBorder", f.TopBorder);
        AppendBorder("bottomBorder", f.BottomBorder);
        sb.Append('}');

        void AppendBorder(string name, ParagraphBorderEdge? edge)
        {
            if (edge is null) return;
            if (has) sb.Append(',');
            has = true;
            sb.Append(JsonString(name)).Append(":{");
            bool edgeHas = false;
            void EdgeString(string key, string? value)
            {
                if (value is null) return;
                if (edgeHas) sb.Append(',');
                edgeHas = true;
                sb.Append(JsonString(key)).Append(':').Append(JsonString(value));
            }
            void EdgeInt(string key, int? value)
            {
                if (!value.HasValue) return;
                if (edgeHas) sb.Append(',');
                edgeHas = true;
                sb.Append(JsonString(key)).Append(':').Append(value.Value);
            }
            EdgeString("style", edge.Style);
            EdgeInt("size", edge.Size);
            EdgeString("color", edge.Color);
            EdgeInt("space", edge.Space);
            sb.Append('}');
        }
    }

    private static void AppendRunFormattingInfo(StringBuilder sb, RunFormattingInfo f)
    {
        sb.Append('{');
        bool has = false;
        void StringValue(string name, string? value)
        {
            if (value is null) return;
            if (has) sb.Append(',');
            has = true;
            sb.Append(JsonString(name)).Append(':').Append(JsonString(value));
        }
        void BoolValue(string name, bool? value)
        {
            if (!value.HasValue) return;
            if (has) sb.Append(',');
            has = true;
            sb.Append(JsonString(name)).Append(':').Append(value.Value ? "true" : "false");
        }
        StringValue("styleId", f.StyleId);
        BoolValue("bold", f.Bold);
        BoolValue("italic", f.Italic);
        BoolValue("underline", f.Underline);
        StringValue("underlineStyle", f.UnderlineStyle);
        BoolValue("strike", f.Strike);
        BoolValue("code", f.Code);
        StringValue("color", f.Color);
        StringValue("highlight", f.Highlight);
        StringValue("vertAlign", f.VertAlign);
        if (f.FontSizePts.HasValue)
        {
            if (has) sb.Append(',');
            has = true;
            sb.Append("\"fontSizePts\":").Append(f.FontSizePts.Value.ToString(
                System.Globalization.CultureInfo.InvariantCulture));
        }
        StringValue("fontFamily", f.FontFamily);
        BoolValue("caps", f.Caps);
        BoolValue("smallCaps", f.SmallCaps);
        BoolValue("hidden", f.Hidden);
        sb.Append('}');
    }

    private static void AppendTableStyleFormatting(StringBuilder sb, TableStyleFormatting f)
    {
        sb.Append('{');
        bool has = false;
        void StringValue(string name, string? value)
        {
            if (value is null) return;
            if (has) sb.Append(',');
            has = true;
            sb.Append(JsonString(name)).Append(':').Append(JsonString(value));
        }
        void IntValue(string name, int? value)
        {
            if (!value.HasValue) return;
            if (has) sb.Append(',');
            has = true;
            sb.Append(JsonString(name)).Append(':').Append(value.Value);
        }
        StringValue("alignment", f.Alignment);
        IntValue("widthTwips", f.WidthTwips);
        IntValue("indentTwips", f.IndentTwips);
        StringValue("layout", f.Layout);
        if (f.HasBorders.HasValue)
        {
            if (has) sb.Append(',');
            has = true;
            sb.Append("\"hasBorders\":").Append(f.HasBorders.Value ? "true" : "false");
        }
        StringValue("cellShadingFill", f.CellShadingFill);
        sb.Append('}');
    }

    private static void AppendNullableBool(StringBuilder sb, string name, bool? value)
    {
        if (!value.HasValue) return;
        sb.Append(',').Append(JsonString(name)).Append(':')
          .Append(value.Value ? "true" : "false");
    }

    private static void AppendHeaderFooterRefs(
        StringBuilder sb, string key, IReadOnlyList<HeaderFooterRef> refs)
    {
        sb.Append(key).Append('[');
        for (int i = 0; i < refs.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append("{\"kind\":").Append(JsonString(HeaderFooterKindToString(refs[i].Kind)))
              .Append(",\"partUri\":").Append(JsonString(refs[i].PartUri))
              .Append(",\"inherited\":").Append(refs[i].Inherited ? "true" : "false").Append('}');
        }
        sb.Append(']');
    }

    /// <summary>Outbound counterpart of <see cref="ParseHeaderFooterKind"/>.</summary>
    private static string HeaderFooterKindToString(HeaderFooterKind kind) => kind switch
    {
        HeaderFooterKind.First => "first",
        HeaderFooterKind.Even => "even",
        _ => "default",
    };

    private static string NumberFormatToString(NumberFormat f) => NumberFormats.ToOoxml(f);

    public static string SerializeAnnotations(IReadOnlyList<DocumentAnnotation> anns)
    {
        var sb = new StringBuilder(anns.Count * 200 + 2);
        sb.Append('[');
        for (int i = 0; i < anns.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var a = anns[i];
            sb.Append("{\"id\":").Append(JsonString(a.Id ?? string.Empty))
              .Append(",\"labelId\":").Append(JsonString(a.LabelId ?? string.Empty))
              .Append(",\"label\":").Append(JsonString(a.Label ?? string.Empty))
              .Append(",\"color\":").Append(JsonString(a.Color ?? string.Empty))
              .Append(",\"bookmarkName\":").Append(JsonString(a.BookmarkName ?? string.Empty));
            if (a.Author is not null)
                sb.Append(",\"author\":").Append(JsonString(a.Author));
            if (a.Created.HasValue)
                sb.Append(",\"created\":").Append(JsonString(a.Created.Value.ToString("o")));
            if (a.AnnotatedText is not null)
                sb.Append(",\"annotatedText\":").Append(JsonString(a.AnnotatedText));
            if (a.Metadata is { Count: > 0 })
            {
                sb.Append(",\"metadata\":{");
                bool firstMeta = true;
                foreach (var kv in a.Metadata)
                {
                    if (!firstMeta) sb.Append(',');
                    firstMeta = false;
                    sb.Append(JsonString(kv.Key)).Append(':').Append(JsonString(kv.Value ?? string.Empty));
                }
                sb.Append('}');
            }
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    // ─── Deserializers ──────────────────────────────────────────────────

    /// <summary>
    /// Parses a camelCase annotation JSON object into a <see cref="DocumentAnnotation"/>
    /// using only <see cref="JsonDocument"/> — safe in trimmed WASM Release builds
    /// where reflection-based <see cref="JsonSerializer"/> is disabled. Both the
    /// WASM bridge and the stdio NDJSON host route annotation writes through here,
    /// so wire shape stays unified across transports.
    /// </summary>
    public static DocumentAnnotation DeserializeAnnotation(string json)
    {
        if (string.IsNullOrEmpty(json))
            throw new System.ArgumentException("annotation JSON is null or empty");

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            throw new System.ArgumentException("annotation JSON must be an object");

        var annotation = new DocumentAnnotation
        {
            Id            = TryGetString(root, "id", string.Empty) ?? string.Empty,
            LabelId       = TryGetString(root, "labelId", string.Empty) ?? string.Empty,
            Label         = TryGetString(root, "label", string.Empty) ?? string.Empty,
            Color         = TryGetString(root, "color", string.Empty) ?? string.Empty,
            Author        = TryGetString(root, "author", string.Empty) ?? string.Empty,
            BookmarkName  = TryGetString(root, "bookmarkName", string.Empty) ?? string.Empty,
            AnnotatedText = TryGetString(root, "annotatedText", string.Empty) ?? string.Empty,
            Metadata      = new System.Collections.Generic.Dictionary<string, string>(),
        };

        if (TryGetDateTime(root, "created", out var created))
            annotation.Created = created;

        if (TryGetIntNullable(root, "startPage", out var startPage))
            annotation.StartPage = startPage;

        if (TryGetIntNullable(root, "endPage", out var endPage))
            annotation.EndPage = endPage;

        // Honour explicit pageInfoStale if supplied; default already true on the model.
        if (root.TryGetProperty("pageInfoStale", out var stale) &&
            (stale.ValueKind == JsonValueKind.True || stale.ValueKind == JsonValueKind.False))
            annotation.PageInfoStale = stale.GetBoolean();

        if (TryGetDateTime(root, "pageInfoComputedAt", out var computedAt))
            annotation.PageInfoComputedAt = computedAt;

        if (root.TryGetProperty("metadata", out var meta) && meta.ValueKind == JsonValueKind.Object)
        {
            foreach (var kv in meta.EnumerateObject())
                if (kv.Value.ValueKind == JsonValueKind.String)
                    annotation.Metadata[kv.Name] = kv.Value.GetString()!;
        }

        return annotation;
    }

    /// <summary>
    /// Parses a camelCase annotation-update JSON object into an <see cref="AnnotationUpdate"/>
    /// using only <see cref="JsonDocument"/> — trim-safe under the WASM Release build.
    /// <see cref="AnnotationUpdate.MetadataPatch"/> honours explicit JSON nulls (a null
    /// value means "remove this key"), while missing fields leave the existing annotation
    /// value unchanged.
    /// </summary>
    public static AnnotationUpdate DeserializeAnnotationUpdate(string json)
    {
        if (string.IsNullOrEmpty(json))
            throw new System.ArgumentException("annotation update JSON is null or empty");

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            throw new System.ArgumentException("annotation update JSON must be an object");

        System.Collections.Generic.Dictionary<string, string?>? patch = null;
        if (root.TryGetProperty("metadataPatch", out var mp) && mp.ValueKind == JsonValueKind.Object)
        {
            patch = new System.Collections.Generic.Dictionary<string, string?>();
            foreach (var kv in mp.EnumerateObject())
                patch[kv.Name] = kv.Value.ValueKind == JsonValueKind.Null ? null : kv.Value.GetString();
        }

        return new AnnotationUpdate
        {
            LabelId       = TryGetString(root, "labelId", null),
            Label         = TryGetString(root, "label", null),
            Color         = TryGetString(root, "color", null),
            Author        = TryGetString(root, "author", null),
            MetadataPatch = patch,
        };
    }

    private static bool TryGetDateTime(JsonElement root, string name, out System.DateTime value)
    {
        if (root.TryGetProperty(name, out var prop) && prop.ValueKind == JsonValueKind.String
            && prop.TryGetDateTime(out var dt))
        {
            value = dt;
            return true;
        }
        value = default;
        return false;
    }

    private static bool TryGetIntNullable(JsonElement root, string name, out int? value)
    {
        if (root.TryGetProperty(name, out var prop))
        {
            if (prop.ValueKind == JsonValueKind.Number && prop.TryGetInt32(out var n))
            {
                value = n;
                return true;
            }
            if (prop.ValueKind == JsonValueKind.Null)
            {
                value = null;
                return true;
            }
        }
        value = null;
        return false;
    }
}
