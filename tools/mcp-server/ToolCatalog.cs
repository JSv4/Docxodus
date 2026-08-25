#nullable enable

using System.Collections.Generic;
using System.Text.Json;

namespace Docxodus.McpServer;

/// <summary>One MCP tool advertised by <c>tools/list</c>. <see cref="InputSchemaJson"/> is a raw
/// JSON Schema object (draft 2020-12 subset: type/properties/required/enum/description) embedded
/// verbatim into the response — see <see cref="Dispatcher"/> for what each action actually does.</summary>
internal sealed record ToolDefinition(string Name, string Description, string InputSchemaJson);

/// <summary>
/// The tool surface this server advertises: three lifecycle tools (open/save/close) plus seventeen
/// read or grouped-intent tools. Grouped tools accept an <c>action</c> discriminator and
/// action-specific arguments. See <c>docs/architecture/docx_agent_server.md</c> for the full contract, the
/// mapping of every action onto the underlying Docxodus API, and the documented capability gaps.
/// </summary>
internal static class ToolCatalog
{
    public static readonly IReadOnlyList<ToolDefinition> Tools = new[]
    {
        new ToolDefinition(
            "docxodus_open",
            "Open a .docx file into an in-memory editing session and return a session_id. All other tools (except docxodus_open itself) take that session_id.",
            """
            {
              "type": "object",
              "properties": {
                "path": { "type": "string", "description": "Location of the .docx within this server's configured document scope. Relative locations resolve under the scope root; an absolute path is accepted only if it falls inside that root. A location outside it is rejected — the scope is set by whoever launched the server and cannot be widened from here." },
                "trackedChanges": { "type": "string", "enum": ["accept", "render_inline", "strip_deletions"], "description": "How mutating ops record their own edits. 'accept' (default) applies them directly; 'render_inline' wraps them as w:ins/w:del tracked changes; 'strip_deletions' drops deleted content outright." },
                "revisionAuthor": { "type": "string", "description": "Author name stamped on tracked-change markup when trackedChanges is render_inline." },
                "undoDepth": { "type": "integer", "description": "Maximum undo steps retained. Default 20. Each step is a full document snapshot, so this is a step count, not a memory bound — see undoMemoryBudgetBytes." },
                "undoMemoryBudgetBytes": { "type": "integer", "description": "Approximate ceiling on memory held by undo/redo snapshots. Default 134217728 (128 MiB). Oldest history is discarded when exceeded, so on a large document undo may not reach the full undoDepth. Set 0 to bound by depth alone." },
                "persistAnchorIds": { "type": "boolean", "description": "Default false. When true, docxodus_save keeps the anchor-id bookkeeping in the written file, so a session opened over it later resolves the anchor ids this session hands out. Costs file size (hundreds of KB on a large document) — turn it on only when a workflow needs a close+reopen without losing its anchors (no longer needed just to switch trackedChanges mode — docxodus_track_changes set_mode does that in place). docxodus_save can also override this per call." },
                "captureInitialProjection": { "type": "boolean", "description": "Default true. Retains the opening package and projection as the session's comparison baseline, which is what docxodus_get_content format 'semantic_changes' and the markdown diff compare against — at the memory cost of one package copy per open session. Set false on a long-lived server that never asks for those, and they refuse with an error instead." }
              },
              "required": ["path"]
            }
            """),
        new ToolDefinition(
            "docxodus_save",
            "Write a session's current in-memory state back to the document store.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "path": { "type": "string", "description": "Destination within the server's document scope, resolved the same way docxodus_open resolves its path. Defaults to the location the session was opened from (overwrite in place)." },
                "persistAnchorIds": { "type": "boolean", "description": "Override the session's open-time persistAnchorIds for this save only. true: keep the anchor-id bookkeeping in the written file so reopening it resolves the same anchor ids (an anchor-stable checkpoint, at a file-size cost). false: strip it (a clean deliverable from a session opened with persistAnchorIds true). Absent: use the session's setting." }
              },
              "required": ["sessionId"]
            }
            """),
        new ToolDefinition(
            "docxodus_close",
            "Discard a session and free its in-memory document. Unsaved changes are lost.",
            """
            { "type": "object", "properties": { "sessionId": { "type": "string" } }, "required": ["sessionId"] }
            """),
        new ToolDefinition(
            "docxodus_get_content",
            "Read a session's document content in one of several formats.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "format": { "type": "string", "enum": ["markdown", "html", "text", "blocks", "info", "version", "semantic_changes", "verification", "check_preconditions", "styles", "formatting", "spans", "manifest"], "description": "markdown/text: anchor-addressed projection; html: rendered HTML; blocks: structural metadata; info: version plus edit and per-anchor section facts; version: monotonic document version; semantic_changes: stable versioned changes from the session's opening package to its current state; verification: bounded default deliverable gate over the session's clean-save checkpoint, using opening bytes as baseline; check_preconditions: read-only guard evaluation; styles: explicit style catalog and resolved properties; formatting: direct/effective paragraph and run formatting for anchorId; spans: mutation-compatible inline spans for anchorId; manifest: deterministic package entries, hashes, relationships, facts, and validation findings for the full logical package." },
                "anchorId": { "type": "string", "description": "Optional for markdown/html/text/info and guard evaluation; required for formatting/spans; not accepted for manifest, semantic_changes, or verification. Those formats always describe the complete package/document. Returned formatting/list/section anchors and span ranges can be passed unchanged to mutation tools." },
                "citation": {
                  "type": "object", "additionalProperties": false,
                  "properties": {
                    "documentVersion": { "type": "integer", "minimum": 0 },
                    "rendererFingerprint": { "type": "string", "minLength": 1 }
                  },
                  "required": ["documentVersion", "rendererFingerprint"]
                },
                "preconditions": { "type": "object", "description": "check_preconditions: expectedVersion and/or anchorId plus expectedContentHash, expectedText/expectedTextRange, expectedKind, expectedScope, or expectedMatchCount." }
              },
              "required": ["sessionId", "format"]
            }
            """),
        new ToolDefinition(
            "docxodus_preview",
            "Render a session's document (or a single block) to HTML for the host's inline preview widget. With an exact registered citation, the widget materializes the cited physical page and navigates to its highlighted fragment. The markup travels in the result's _meta for the widget only; call again after edits to refresh it.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "anchorId": { "type": "string", "description": "Optional. Render just this block (any addressable anchor, including hdr*/ftr* scopes) instead of the whole document. Whole-document renders include the converter's stylesheet; single-block renders are bare markup." }
                ,"citation": {
                  "type": "object", "additionalProperties": false,
                  "properties": {
                    "documentVersion": { "type": "integer", "minimum": 0 },
                    "rendererFingerprint": { "type": "string", "minLength": 1 }
                  },
                  "required": ["documentVersion", "rendererFingerprint"]
                }
              },
              "required": ["sessionId"]
            }
            """),
        new ToolDefinition(
            "docxodus_pagination",
            "Register or consume a browser-materialized PageMap. Core never estimates page numbers; unavailable, continuous, stale, and renderer-mismatched layouts are explicit.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "action": { "type": "string", "enum": ["register", "status", "cite"] },
                "pageMap": {
                  "type": "object", "additionalProperties": false,
                  "properties": {
                    "schemaVersion": { "type": "integer", "const": 1 },
                    "mode": { "type": "string", "enum": ["paginated", "continuous"] },
                    "availability": { "type": "string", "enum": ["available", "unavailable"] },
                    "documentVersion": { "type": "integer", "minimum": 0 },
                    "rendererFingerprint": { "type": "string", "minLength": 1 },
                    "pages": {
                      "type": "array",
                      "items": {
                        "type": "object", "additionalProperties": false,
                        "properties": {
                          "pageNumber": { "type": "integer", "minimum": 1 },
                          "pageInSection": { "type": "integer", "minimum": 1 },
                          "width": { "type": "number", "exclusiveMinimum": 0 },
                          "height": { "type": "number", "exclusiveMinimum": 0 },
                          "sectionIndex": { "type": "integer", "minimum": 0 },
                          "pageName": { "type": "string", "minLength": 1 }
                        },
                        "required": ["pageNumber", "pageInSection", "width", "height", "pageName"]
                      }
                    },
                    "fragments": {
                      "type": "array",
                      "items": {
                        "type": "object", "additionalProperties": false,
                        "properties": {
                          "fragmentId": { "type": "string", "minLength": 1 },
                          "anchorId": { "type": "string", "minLength": 1 },
                          "fragmentIndex": { "type": "integer", "minimum": 0 },
                          "pageNumber": { "type": "integer", "minimum": 1 },
                          "geometry": {
                            "type": "object", "additionalProperties": false,
                            "properties": {
                              "x": { "type": "number", "minimum": 0 },
                              "y": { "type": "number", "minimum": 0 },
                              "width": { "type": "number", "exclusiveMinimum": 0 },
                              "height": { "type": "number", "exclusiveMinimum": 0 }
                            },
                            "required": ["x", "y", "width", "height"]
                          },
                          "story": { "type": "string", "enum": ["body", "header", "footer", "footnote", "endnote", "comment"] },
                          "inTableCell": { "type": "boolean" }
                        },
                        "required": ["fragmentId", "anchorId", "fragmentIndex", "pageNumber", "geometry", "story", "inTableCell"]
                      }
                    }
                  },
                  "required": ["schemaVersion", "mode", "availability", "documentVersion", "rendererFingerprint", "pages", "fragments"]
                },
                "expectedRendererFingerprint": { "type": "string", "description": "register: optional independently expected fingerprint; mismatch rejects the map." },
                "anchorId": { "type": "string", "description": "cite: canonical kind:scope:unid anchor." },
                "citation": {
                  "type": "object", "additionalProperties": false,
                  "properties": {
                    "documentVersion": { "type": "integer", "minimum": 0 },
                    "rendererFingerprint": { "type": "string", "minLength": 1 }
                  },
                  "required": ["documentVersion", "rendererFingerprint"]
                }
              },
              "required": ["sessionId", "action"],
              "oneOf": [
                { "properties": { "action": { "const": "register" } }, "required": ["pageMap"] },
                { "properties": { "action": { "const": "status" } } },
                { "properties": { "action": { "const": "cite" } }, "required": ["anchorId", "citation"] }
              ]
            }
            """),
        new ToolDefinition(
            "docxodus_search",
            "Find text or structural nodes in a session's document. Returns anchor ids usable directly as the anchorId/cellAnchorId argument of every other tool.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "mode": { "type": "string", "enum": ["text", "regex", "kind", "annotation", "bookmark"], "description": "text: literal substring search (Grep). regex: .NET regex search. kind: all blocks of a structural kind (p, h, li, tbl, tc, ...). annotation: blocks touched by a given annotation id. bookmark: blocks anchored by a given Word bookmark name." },
                "query": { "type": "string", "description": "The needle: literal text, regex pattern, block kind, annotation id, or bookmark name depending on mode." },
                "caseSensitive": { "type": "boolean", "description": "Default false (case-insensitive)." },
                "contextChars": { "type": "integer", "description": "Characters of context captured on each side of a text/regex match. Default 80." },
                "scope": { "type": "string", "enum": ["body", "headers", "footers", "header_footer", "all"], "description": "text/regex only: package stories to search. Default body preserves existing behavior; headers/footers cover every hdr*/ftr* part, header_footer combines them, and all includes body, running stories, notes, and comments." },
                "maxResults": { "type": "integer", "description": "Cap the number of matches returned. Default unlimited." }
                ,"citation": {
                  "type": "object", "additionalProperties": false,
                  "properties": {
                    "documentVersion": { "type": "integer", "minimum": 0 },
                    "rendererFingerprint": { "type": "string", "minLength": 1 }
                  },
                  "required": ["documentVersion", "rendererFingerprint"]
                }
              },
              "required": ["sessionId", "mode", "query"]
            }
            """),
        new ToolDefinition(
            "docxodus_edit",
            "Insert, replace, move, delete, or undo/redo text and blocks, addressed by anchor id.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "preconditions": { "type": "object", "description": "Optional optimistic guards evaluated immediately before the mutation: expectedVersion, anchorId, expectedContentHash, expectedText/expectedTextRange, expectedKind, expectedScope, expectedMatchCount." },
                "action": { "type": "string", "enum": ["insert_paragraph", "replace_text", "replace_text_range", "delete_block", "move_block", "delete_range", "delete_section", "split_paragraph", "merge_paragraphs", "undo", "redo"] },
                "anchorId": { "type": "string", "description": "Target block. Required for every action except delete_range, delete_section, undo, redo." },
                "position": { "type": "string", "enum": ["before", "after"], "description": "insert_paragraph/move_block only." },
                "markdown": { "type": "string", "description": "insert_paragraph/replace_text payload, in the supported markdown subset (headings, bullet/ordered lists, bold/italic/code/strike, links, hard breaks)." },
                "find": { "type": "string", "description": "replace_text_range: literal text to find within the block. A find that matches nothing fails with text_not_found (both directly and as a batch step); pass preconditions.expectedMatchCount 0 to assert absence as a successful no-op instead." },
                "replace": { "type": "string", "description": "replace_text_range: replacement text." },
                "caseSensitive": { "type": "boolean", "description": "replace_text_range only. Default false." },
                "toAnchorIdExclusive": { "type": "string", "description": "delete_range: end boundary, exclusive." },
                "fromAnchorId": { "type": "string", "description": "delete_range: start boundary (use anchorId as the field name for the start would be ambiguous with toAnchorIdExclusive; both are required together)." },
                "headingAnchorId": { "type": "string", "description": "delete_section: heading whose section (heading + everything until the next same-or-higher heading) should be removed." },
                "characterOffset": { "type": "integer", "description": "split_paragraph: character offset to split at." },
                "secondAnchorId": { "type": "string", "description": "merge_paragraphs: the paragraph absorbed into anchorId." },
                "sourceAnchorId": { "type": "string", "description": "move_block: block to relocate." },
                "targetAnchorId": { "type": "string", "description": "move_block: reference block used with position." }
              },
              "required": ["sessionId", "action"]
            }
            """),
        new ToolDefinition(
            "docxodus_format",
            "Apply character or paragraph formatting, addressed by anchor id.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "preconditions": { "type": "object", "description": "Optional optimistic mutation guards; omitted preserves legacy behavior." },
                "action": { "type": "string", "enum": ["apply_format", "apply_format_by_substring", "set_paragraph_style", "set_paragraph_format", "set_list_level", "remove_list_membership", "apply_list_format"] },
                "anchorId": { "type": "string" },
                "span": { "type": "object", "properties": { "start": { "type": "integer" }, "length": { "type": "integer" } }, "description": "apply_format only. Omit to format the whole block." },
                "substring": { "type": "string", "description": "apply_format_by_substring: literal text within the block to format." },
                "format": {
                  "type": "object",
                  "description": "apply_format / apply_format_by_substring payload. Omitted fields are left unchanged.",
                  "properties": {
                    "bold": { "type": "boolean" }, "italic": { "type": "boolean" }, "underline": { "type": "boolean" },
                    "strike": { "type": "boolean" }, "code": { "type": "boolean" },
                    "color": { "type": "string", "description": "Hex RGB, no '#'." },
                    "vertAlign": { "type": "string", "enum": ["superscript", "subscript", "none"] },
                    "fontSizePts": { "type": "number" }, "fontFamily": { "type": "string" }
                  }
                },
                "styleId": { "type": "string", "description": "set_paragraph_style: a style id from the document's style definitions (e.g. Heading1)." },
                "paragraphFormat": {
                  "type": "object",
                  "description": "set_paragraph_format payload. Omitted fields are left unchanged.",
                  "properties": {
                    "alignment": { "type": "string", "enum": ["left", "center", "right", "justify"] },
                    "indentDelta": { "type": "integer", "description": "Twips to add to the current left indent (negative to outdent). 1440 twips = 1 inch." },
                    "firstLineIndent": { "type": "integer", "minimum": 0, "description": "Absolute first-line indent in twips (w:ind/@w:firstLine; 1440 = 1 inch, 720 = 0.5 inch). 0 = explicitly none. Mutually exclusive with hangingIndent (Word stores one or the other); setting it removes any hanging indent." },
                    "hangingIndent": { "type": "integer", "minimum": 0, "description": "Absolute hanging indent in twips (w:ind/@w:hanging; 1440 = 1 inch) — every line EXCEPT the first starts this far right of the left edge. Mutually exclusive with firstLineIndent; setting it removes any first-line indent." },
                    "spacingBefore": { "type": "integer", "minimum": 0, "description": "Absolute space above the paragraph in twips (w:spacing/@w:before). 20 twips = 1pt, so 240 = 12pt." },
                    "spacingAfter": { "type": "integer", "minimum": 0, "description": "Absolute space below the paragraph in twips (w:spacing/@w:after). 20 twips = 1pt, so 240 = 12pt." },
                    "lineSpacing": { "type": "integer", "minimum": 0, "description": "Line spacing (w:spacing/@w:line). Units depend on lineSpacingRule: under \"auto\" (the default) it is 240ths of a line (240 = single, 360 = 1.5x, 480 = double); under \"exact\"/\"atLeast\" it is twips (20 = 1pt, so 480 = 24pt)." },
                    "lineSpacingRule": { "type": "string", "enum": ["auto", "exact", "atLeast"], "description": "How lineSpacing is measured (w:spacing/@w:lineRule). Requires lineSpacing in the same call." },
                    "pageBreakBefore": { "type": "boolean" },
                    "topBorder": {
                      "type": "object",
                      "description": "Adds/replaces the paragraph's top border (w:pBdr/w:top).",
                      "properties": {
                        "style": { "type": "string", "description": "Border line style, e.g. single/double/thick/dotted/dashed. Default \"single\"." },
                        "size": { "type": "integer", "description": "Border weight in eighths of a point. Default 6 (≈0.75pt); a heavy rule ≈ 18-24." },
                        "color": { "type": "string", "description": "Hex RGB without '#', or \"auto\". Default \"auto\"." },
                        "space": { "type": "integer", "description": "Padding between border and text, in points. Default 1." }
                      }
                    },
                    "bottomBorder": {
                      "type": "object",
                      "description": "Adds/replaces the paragraph's bottom border (w:pBdr/w:bottom). This is what an S-1-style horizontal rule is: an (often empty) paragraph with a bottom border.",
                      "properties": {
                        "style": { "type": "string", "description": "Border line style, e.g. single/double/thick/dotted/dashed. Default \"single\"." },
                        "size": { "type": "integer", "description": "Border weight in eighths of a point. Default 6 (≈0.75pt); a heavy rule ≈ 18-24." },
                        "color": { "type": "string", "description": "Hex RGB without '#', or \"auto\". Default \"auto\"." },
                        "space": { "type": "integer", "description": "Padding between border and text, in points. Default 1." }
                      }
                    },
                    "clearBorders": { "type": "boolean", "description": "Remove the entire w:pBdr (all paragraph borders) before applying topBorder/bottomBorder in this same call." }
                  }
                },
                "levelDelta": { "type": "integer", "description": "set_list_level: +1 indents one level, -1 outdents one level." },
                "listFormat": { "type": "string", "enum": ["bullet", "decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman", "decimalParenthesis", "lowerLetterParenthesis", "upperLetterParenthesis", "lowerRomanParenthesis", "upperRomanParenthesis", "none"], "description": "apply_list_format: converts the paragraph into (or out of) a real, auto-numbered Word list. *Parenthesis variants render '(1)'/'(a)'/'(i)'." }
              },
              "required": ["sessionId", "action", "anchorId"]
            }
            """),
        new ToolDefinition(
            "docxodus_create",
            "Insert new structural content: paragraphs, headings, tables, horizontal rules, footnotes/endnotes, running headers/footers, page-number fields.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "preconditions": { "type": "object", "description": "Optional optimistic mutation guards; omitted preserves legacy behavior." },
                "action": { "type": "string", "enum": ["insert_paragraph", "insert_heading", "insert_table", "insert_horizontal_rule", "insert_footnote", "insert_endnote", "insert_page_number_field", "set_header_text", "set_footer_text", "ensure_header_footer_visible"] },
                "anchorId": { "type": "string", "description": "Reference block for insert_paragraph/insert_heading/insert_table/insert_horizontal_rule (paired with position), or the citing paragraph for insert_footnote/insert_endnote, or the target paragraph for insert_page_number_field." },
                "bodyAnchorId": { "type": "string", "description": "set_header_text/set_footer_text/ensure_header_footer_visible: a body block identifying the section whose running story or visibility flags should change." },
                "position": { "type": "string", "enum": ["before", "after"] },
                "text": { "type": "string", "description": "insert_heading: heading text (plain, not markdown)." },
                "level": { "type": "integer", "minimum": 1, "maximum": 6, "description": "insert_heading: 1-6." },
                "markdown": { "type": "string", "description": "insert_paragraph / insert_footnote / insert_endnote / set_header_text / set_footer_text payload." },
                "rows": { "type": "integer" }, "columns": { "type": "integer" },
                "cellContents": { "type": "array", "items": { "type": "string" }, "description": "insert_table: row-major markdown per cell." },
                "cellAlignment": { "type": "string", "enum": ["left", "center", "right", "justify"] },
                "columnWidths": { "type": "array", "items": { "type": "integer" }, "description": "insert_table: twips per column, left to right." },
                "borderless": { "type": "boolean" },
                "ruleStyle": { "type": "string", "enum": ["single", "double", "thick"], "description": "insert_horizontal_rule." },
                "characterOffset": { "type": "integer", "description": "insert_footnote/insert_endnote: character offset within the citing paragraph." },
                "kind": { "type": "string", "enum": ["default", "first", "even"], "description": "set_header_text/set_footer_text/ensure_header_footer_visible: running-story kind. first/even authoring also enables the corresponding Word visibility setting; ensure_header_footer_visible enables it for an already-referenced story." },
                "field": { "type": "string", "enum": ["current_page", "total_pages"], "description": "insert_page_number_field." },
                "numberFormat": { "type": "string", "enum": ["decimal", "upperLetter", "lowerLetter", "upperRoman", "lowerRoman"], "description": "insert_page_number_field: optional explicit \\* switch format." }
              },
              "required": ["sessionId", "action"]
            }
            """),
        new ToolDefinition(
            "docxodus_list",
            "Manage list membership: promote a paragraph (or a whole contiguous run of paragraphs) to a real auto-numbered Word list, renumber/indent, restart numbering, or drop membership. apply_format_range keeps every member in ONE shared w:num instance so the sequence numbers stay intact — use it instead of per-item apply_format when converting an existing list. set_start is Word's 'Set Numbering Value…': restart (or seed) the item's numbering at startValue — a mid-list restart splits the sequence like Word does; clear_start removes the restart from the item's whole sequence.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "preconditions": { "type": "object", "description": "Optional optimistic mutation guards; omitted preserves legacy behavior." },
                "action": { "type": "string", "enum": ["apply_format", "apply_format_range", "set_level", "set_start", "clear_start", "remove", "get_membership"] },
                "anchorId": { "type": "string", "description": "Target paragraph for every action except apply_format_range. remove accepts paragraph, heading, or list-item anchors and overrides style-inherited numbering when necessary." },
                "startValue": { "type": "integer", "description": "set_start: the number the item restarts at (>= 0), e.g. 5 to make this item render as '5.'" },
                "firstAnchorId": { "type": "string", "description": "apply_format_range: first paragraph of the contiguous sibling run (inclusive)." },
                "lastAnchorId": { "type": "string", "description": "apply_format_range: last paragraph of the run (inclusive; either document order)." },
                "listFormat": { "type": "string", "enum": ["bullet", "decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman", "decimalParenthesis", "lowerLetterParenthesis", "upperLetterParenthesis", "lowerRomanParenthesis", "upperRomanParenthesis", "none"], "description": "apply_format / apply_format_range: creates real w:numPr numbering. Plain numbered formats render '1.'/'a.'/'i.'; *Parenthesis variants render '(1)'/'(a)'/'(i)' (legal-drafting presets). 'none' strips membership." },
                "levelDelta": { "type": "integer", "description": "set_level: +1 indents, -1 outdents." }
              },
              "required": ["sessionId", "action"]
            }
            """),
        new ToolDefinition(
            "docxodus_comment",
            "Create and manage native Word review comments (real w:comment markup — visible in Word/Google Docs/LibreOffice's Reviewing pane): comment on a character span or tracked revision, reply in the same native thread, resolve/reopen, update, remove, or list. Comments are addressed by their cmt anchor (from add/reply's created list or the projection's # Comments section). list reports parentAnchorId/resolved when Word extension metadata exists. For the semantic highlight/label overlay see docxodus_annotate.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "preconditions": { "type": "object", "description": "Optional optimistic mutation guards; omitted preserves legacy behavior." },
                "action": { "type": "string", "enum": ["add", "reply", "resolve", "update", "remove", "list"] },
                "anchorId": { "type": "string", "description": "add: the body paragraph to comment on. Mutually exclusive with revisionId." },
                "span": { "type": "object", "properties": { "start": { "type": "integer" }, "length": { "type": "integer" } }, "description": "add: character range within the paragraph. Omit to comment on the whole block." },
                "revisionId": { "type": "string", "description": "add: alternatively, the id from docxodus_track_changes list. The exact live revision extent is targeted; unknown/already-resolved ids fail with revision_not_found." },
                "author": { "type": "string", "description": "add/reply: comment author (required)." },
                "initials": { "type": "string", "description": "add/reply: optional author initials." },
                "date": { "type": "string", "description": "add/reply: optional ISO-8601 timestamp; w:date is written only when provided (omitting keeps output deterministic)." },
                "markdown": { "type": "string", "description": "add/reply/update: the comment body (same markdown subset as other payloads)." },
                "commentAnchorId": { "type": "string", "description": "reply: parent comment; resolve/update/remove: target comment definition anchor (kind cmt)." },
                "resolved": { "type": "boolean", "description": "resolve: true marks done (default); false reopens while preserving thread parentage." }
              },
              "required": ["sessionId", "action"]
            }
            """),
        new ToolDefinition(
            "docxodus_annotate",
            "Create and manage anchor-addressed annotations: a highlight + label overlay stored in a bookmark and a custom-XML part, for semantically tagging regions for external tools (e.g. OpenContracts). Not a Word review comment — those live in docxodus_comment.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "preconditions": { "type": "object", "description": "Optional optimistic mutation guards; omitted preserves legacy behavior." },
                "action": { "type": "string", "enum": ["add", "update", "remove", "move", "list", "find"] },
                "anchorId": { "type": "string", "description": "add/move: block to attach the annotation to." },
                "span": { "type": "object", "properties": { "start": { "type": "integer" }, "length": { "type": "integer" } }, "description": "add/move: character range within the block. Omit to annotate the whole block." },
                "annotationId": { "type": "string", "description": "update/remove/move: id of an existing annotation. add: optional; a 16-char hex id is generated if omitted." },
                "labelId": { "type": "string", "description": "add: category id, e.g. 'REVIEW_NOTE'." },
                "label": { "type": "string", "description": "add: human-readable text." },
                "color": { "type": "string", "description": "add: hex highlight color, e.g. '#FFEB3B'." },
                "author": { "type": "string" },
                "update": {
                  "type": "object",
                  "description": "update action payload. Omitted fields are left unchanged.",
                  "properties": {
                    "labelId": { "type": "string" }, "label": { "type": "string" },
                    "color": { "type": "string" }, "author": { "type": "string" }
                  }
                },
                "newAnchorId": { "type": "string", "description": "move: new target block." },
                "newSpan": { "type": "object", "properties": { "start": { "type": "integer" }, "length": { "type": "integer" } }, "description": "move: new character range." },
                "query": { "type": "string", "description": "find: annotation id to look up." }
              },
              "required": ["sessionId", "action"]
            }
            """),
        new ToolDefinition(
            "docxodus_links",
            "Enumerate and safely mutate native Word hyperlinks and bookmarks across body, headers, footers, footnotes, and endnotes. Internal hyperlinks target an existing bookmark with relationship-free w:anchor markup; external links use the containing story part's relationship. Bookmark ranges may cross paragraphs in one story part but not package parts. Tracked render-inline mode rejects these metadata mutations explicitly.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "action": { "type": "string", "enum": ["list_hyperlinks", "add_hyperlink", "update_hyperlink", "remove_hyperlink", "list_bookmarks", "add_bookmark", "rename_bookmark", "move_bookmark", "remove_bookmark"] },
                "scope": { "type": "string", "enum": ["body", "headers", "footers", "footnotes", "endnotes", "comments", "all"], "description": "Listing only; default all." },
                "anchorId": { "type": "string", "description": "add_hyperlink: containing paragraph anchor." },
                "startOffset": { "type": "integer", "description": "add_hyperlink/add_bookmark/move_bookmark: zero-based character boundary." },
                "length": { "type": "integer", "minimum": 1, "description": "add_hyperlink: selected text length." },
                "kind": { "type": "string", "enum": ["external", "internal"], "description": "add/update_hyperlink target representation." },
                "target": { "type": "string", "description": "External URI or existing bookmark name (without '#')." },
                "hyperlinkId": { "type": "string", "description": "update/remove_hyperlink: id returned by list/add." },
                "name": { "type": "string", "description": "Bookmark name for add/rename/move/remove." },
                "newName": { "type": "string", "description": "rename_bookmark destination name; inbound internal links are retargeted atomically." },
                "startAnchorId": { "type": "string", "description": "add/move_bookmark range start paragraph." },
                "endAnchorId": { "type": "string", "description": "add/move_bookmark range end paragraph in the same story part." },
                "endOffset": { "type": "integer", "description": "add/move_bookmark exclusive end boundary." }
              },
              "required": ["sessionId", "action"]
            }
            """),
        new ToolDefinition(
            "docxodus_images",
            "Inspect and mutate native Word images. Binary payloads cross this JSON boundary only as base64; the server never fetches URLs or reads image paths. PNG, JPEG, GIF, BMP, and TIFF are writable; WebP, legacy VML, external links, and unsupported DrawingML remain inspection-only. Rendered dimensions are points; floating offsets/distances are exact EMUs at a documented 96-DPI default.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string", "description": "Required except for capabilities." },
                "action": { "type": "string", "enum": ["capabilities", "list", "insert", "replace", "set_dimensions", "set_metadata", "set_floating_layout", "remove"] },
                "scope": { "type": "string", "enum": ["body", "headers", "footers", "footnotes", "endnotes", "comments", "all"] },
                "anchorId": { "type": "string", "description": "insert: paragraph anchor." },
                "characterOffset": { "type": "integer", "minimum": 0 },
                "imageId": { "type": "string", "description": "replace/set/remove: id from list or insert." },
                "imageBase64": { "type": "string", "description": "insert/replace only; raw image bytes encoded as base64." },
                "options": { "type": "object", "description": "insert options: placement inline|floating, widthPoints, heightPoints, preserveAspect, altText, title, and optional floatingLayout." },
                "dimensions": { "type": "object", "description": "set_dimensions: widthPoints and/or heightPoints plus preserveAspect (default true)." },
                "altText": { "type": ["string", "null"], "description": "set_metadata full value; null removes it." },
                "title": { "type": ["string", "null"], "description": "set_metadata full value; null removes it." },
                "layout": { "type": "object", "description": "set_floating_layout: none/square wrap; typed references/alignments; exact EMU positions/distances and flags." }
              },
              "required": ["action"]
            }
            """),
        new ToolDefinition(
            "docxodus_content_controls",
            "Inspect and fill native Word content controls (structured-document tags) while preserving their wrappers and metadata. Bound controls fail closed unless bindingPolicy is detach_target, which removes only the selected control's own binding. Text, checkbox, date, and list whole-content replacements are refused for row/cell placements; nested targets and render_inline tracked-change mode fail closed for every whole-control fill.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "preconditions": { "type": "object", "description": "Optional optimistic mutation guards; omitted preserves legacy behavior." },
                "action": { "type": "string", "enum": ["list", "fill_text", "fill_rich_text", "set_checked", "set_date", "select_item", "fill_picture", "add_repeating_item", "remove_repeating_item"] },
                "scope": { "type": "string", "enum": ["body", "headers", "footers", "footnotes", "endnotes", "comments", "all"] },
                "anchorId": { "type": "string", "description": "Target sdt anchor returned by list." },
                "text": { "type": "string", "description": "fill_text payload." },
                "markdown": { "type": "string", "description": "fill_rich_text payload." },
                "checked": { "type": "boolean", "description": "set_checked value." },
                "value": { "type": "string", "description": "set_date ISO-8601 value or select_item value/display text." },
                "displayText": { "type": "string", "description": "Optional set_date displayed text." },
                "imageBase64": { "type": "string", "description": "fill_picture raw image bytes as base64." },
                "sectionAnchorId": { "type": "string", "description": "add_repeating_item section control." },
                "afterItemAnchorId": { "type": "string", "description": "Optional direct item after which the clone is inserted." },
                "itemAnchorId": { "type": "string", "description": "remove_repeating_item direct item." },
                "bindingPolicy": { "type": "string", "enum": ["preserve", "detach_target"], "description": "Default preserve. detach_target removes only the selected target's own native w:dataBinding or w15:dataBinding element; a bound ancestor still fails closed." }
              },
              "required": ["sessionId", "action"]
            }
            """),
        new ToolDefinition(
            "docxodus_track_changes",
            "List, selectively accept/reject (by revisionId), or atomically bulk-resolve live tracked changes including structural cell, content-control, and numbering families — switch how the session records its OWN subsequent edits (set_mode) — or prove this redline accepts to an intended final and rejects to a baseline (prove_reversibility).",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "preconditions": { "type": "object", "description": "Optional optimistic guards for accept/reject/accept_all/reject_all." },
                "action": { "type": "string", "enum": ["list", "accept", "reject", "accept_all", "reject_all", "set_mode", "prove_reversibility"], "description": "'list' returns the live part-aware registry, including all affected anchors and fail-closed diagnostics. Individual and bulk resolution use the same resolver and are undoable. Bulk resolution is atomic and REFUSES the whole document (revision_unsupported/revision_malformed/revision_ambiguous, nothing mutated) on the first entry it cannot resolve safely — a missing or non-numeric w:id, one w:id shared by two live revisions in a part, customXml move ranges, revisions under m:ctrlPr, w:del on a run's w:rPr or a paragraph's w:numPr, a w:sdt envelope whose range topology is not Word's two-pair shape, an unattached w:numberingChange, or a malformed cell marker. There is no force mode; run 'list' and read each entry's diagnostic to see what blocks it." },
                "revisionId": { "type": "string", "description": "accept/reject: the opaque stable id from 'list' (e.g. 'rev2-a1b2c3d4e5f60718293a'). Legacy revNNN ids remain accepted only when uniquely resolvable. Unknown or already-resolved ids fail with revision_not_found — re-list for the current set." },
                "author": { "type": "string", "description": "list: only return revisions by this author." },
                "changeType": { "type": "string", "enum": ["insert", "delete", "move", "format", "structure"], "description": "list: only return revisions of this coarse type." },
                "family": { "type": "string", "description": "list: exact family filter, such as cell_delete, content_control_insert, or numbering_change." },
                "resolutionStatus": { "type": "string", "enum": ["supported", "unsupported", "malformed", "ambiguous"], "description": "list: fail-closed resolution status filter." },
                "partUri": { "type": "string", "description": "list: exact owning package-part URI." },
                "mode": { "type": "string", "enum": ["accept", "render_inline", "strip_deletions"], "description": "set_mode: how SUBSEQUENT mutations are recorded (same values as docxodus_open's trackedChanges). Never touches already-applied edits — accept does not resolve existing revisions (use accept/accept_all), render_inline does not retroactively track prior direct edits. Not undoable." },
                "revisionAuthor": { "type": "string", "description": "set_mode: author stamped on subsequent tracked-change markup. Absent = leave the current author unchanged; empty string = reset to the 'docxodus' default." },
                "baselinePath": { "type": "string", "description": "prove_reversibility: the document this session's redline was generated against, resolved inside the server's document scope like docxodus_open's path. Rejecting only the generated revisions must reproduce it." },
                "intendedFinalPath": { "type": "string", "description": "prove_reversibility: the document accepting only the generated revisions must reproduce, resolved the same way. Together with baselinePath this is what makes the result a proof rather than a diff." }
              },
              "required": ["sessionId", "action"]
            }
            """),
        new ToolDefinition(
            "docxodus_compare",
            "Compare stored document versions into one native tracked-changes redline, without an open session: two-way (baselinePath vs revisedPath) or N-way consolidate (baselinePath vs revisedPaths, each reviewer's changes attributed to their author). Every path resolves inside the server's document scope exactly like docxodus_open's; the redline is written to outputPath. The result summarizes the generated revisions by author; open the output with docxodus_open to inspect, comment on, or resolve them.",
            """
            {
              "type": "object",
              "properties": {
                "baselinePath": { "type": "string", "description": "The earlier version both forms diff against." },
                "revisedPath": { "type": "string", "description": "Two-way: the later version. Exactly one of revisedPath / revisedPaths." },
                "revisedPaths": { "type": "array", "items": { "type": "string" }, "minItems": 2, "description": "Consolidate: two or more revised versions merged into one redline over the baseline." },
                "author": { "type": "string", "description": "Two-way: the author stamped on generated revisions; absent uses the engine default." },
                "authors": { "type": "array", "items": { "type": "string" }, "description": "Consolidate: reviewer name per revisedPaths entry, same order and length. Absent uses each file's name without extension." },
                "outputPath": { "type": "string", "description": "Where the redline is written, resolved in the document scope like docxodus_save's path." }
              },
              "required": ["baselinePath", "outputPath"]
            }
            """),
        new ToolDefinition(
            "docxodus_mutations",
            "Apply or safely preview a batch of mutating edit/format/create/table/list/comment/link/image/content-control/track-changes actions. Atomic mode commits as one unit. An optional transactionId makes applying retries idempotent within this open session; preview is isolated and cannot carry a transactionId.",
            $$"""
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "transactionId": { "type": "string", "minLength": 1, "maxLength": {{MutationTransactions.MaxTransactionIdLength}}, "pattern": {{JsonSerializer.Serialize(MutationTransactions.TransactionIdNonBlankPattern)}}, "description": {{JsonSerializer.Serialize(MutationTransactions.TransactionIdSchemaDescription)}} },
                "preconditions": { "type": "object", "description": "Optional batch-start guards. Each step args object may also carry its own preconditions." },
                "mode": { "type": "string", "enum": ["atomic", "best_effort", "apply", "preview"], "default": "atomic", "description": "atomic (default): all steps commit as one undo/version unit or fully roll back. best_effort: explicitly retain successful steps after failures. apply: deprecated alias for best_effort. preview: isolated dry-run shorthand using atomic policy unless previewPolicy says best_effort." },
                "preview": { "type": "boolean", "default": false, "description": "Dry-run mode for mode=atomic or mode=best_effort. The complete package is cloned and the live document, version, caches, configuration, and undo/redo history are never touched." },
                "previewPolicy": { "type": "string", "enum": ["atomic", "best_effort"], "default": "atomic", "description": "Policy used by legacy mode=preview. Ignored when mode itself is atomic/best_effort." },
                "previewHtml": { "type": "string", "enum": ["none", "scoped", "full"], "default": "none", "description": "Optionally render shadow-only HTML from the predicted package." },
                "previewAnchorId": { "type": "string", "description": "Required when previewHtml=scoped; the live anchor is resolved only inside the cloned package." },
                "steps": {
                  "type": "array",
                  "items": {
                    "type": "object",
                    "properties": {
                      "tool": { "type": "string", "enum": ["docxodus_edit", "docxodus_format", "docxodus_create", "docxodus_table", "docxodus_list", "docxodus_comment", "docxodus_links", "docxodus_images", "docxodus_content_controls", "docxodus_track_changes"] },
                      "args": { "type": "object", "description": "The same arguments that tool's action takes, minus sessionId (inherited from the batch). transactionId is forbidden here; it belongs only at the batch root." }
                    },
                    "required": ["tool", "args"]
                  }
                }
              },
              "required": ["sessionId", "steps"]
            }
            """),
        new ToolDefinition(
            "docxodus_deliver",
            "Build one verified delivery bundle from a named baseline and the current session. Returns canonical manifest bytes and every available artifact as base64, bounded to 64 MiB before base64 expansion. Production rendering uses the process-owned DOCXODUS_NODE_PATH and DOCXODUS_EXPORT_HOST_PATH configuration; authoritative change-receipt evidence remains available through the programmatic transaction API.",
            """
            {
              "type": "object",
              "additionalProperties": false,
              "properties": {
                "sessionId": { "type": "string" },
                "baselinePath": { "type": "string", "description": "Store-scoped baseline DOCX location." },
                "baselineDocumentVersion": { "type": "integer", "minimum": 0 },
                "finalDocumentName": { "type": "string", "minLength": 1 },
                "finalDocumentVersion": { "type": "integer", "minimum": 0 },
                "revisionPolicy": {
                  "type": "object",
                  "additionalProperties": false,
                  "properties": {
                    "preExistingRevisions": { "type": "string", "enum": ["preserve", "accept", "reject"] },
                    "generatedRevisions": { "type": "string", "enum": ["preserve", "accept", "reject"] }
                  },
                  "required": ["preExistingRevisions", "generatedRevisions"]
                },
                "artifacts": {
                  "type": "array",
                  "minItems": 1,
                  "items": {
                    "type": "object",
                    "additionalProperties": false,
                    "properties": {
                      "artifactId": { "type": "string", "minLength": 1 },
                      "kind": { "type": "string", "enum": ["baselineDocx", "policyBaselineDocx", "workingDocx", "reviewDocx", "finalDocx", "standaloneHtml", "finalPdf", "reviewPdf", "pageMap", "renderReport", "baselinePackageManifest", "finalPackageManifest", "semanticDelta", "packageDelta", "validationReport", "reversibilityProof", "changeReceipt"] },
                      "requiredness": { "type": "string", "enum": ["required", "optional"] },
                      "reviewProfile": { "type": "string", "enum": ["final", "original", "markup"] },
                      "commentProfile": { "type": "string", "enum": ["hidden", "inline", "endnotes", "margin"] }
                    },
                    "required": ["artifactId", "kind", "requiredness"]
                  }
                },
                "failOnDeliverableValidationFailure": { "type": "boolean", "default": true },
                "returnIncompleteBundle": { "type": "boolean", "default": false }
              },
              "required": ["sessionId", "baselinePath", "baselineDocumentVersion", "finalDocumentName", "finalDocumentVersion", "revisionPolicy", "artifacts"]
            }
            """),
        new ToolDefinition(
            "docxodus_table",
            "Inspect canonical table identities and coordinates; create tables; edit rows, columns, and cell content; merge/unmerge cells; and style them after insert.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "preconditions": { "type": "object", "description": "Optional optimistic mutation guards; omitted preserves legacy behavior." },
                "action": { "type": "string", "enum": ["get_metadata", "resolve_cell_anchor", "resolve_cell_coordinate", "insert", "insert_row", "insert_column", "delete_row", "delete_column", "replace_cell_content", "merge_cells", "unmerge_cells", "set_column_widths", "set_borders", "set_shading", "set_repeat_header_row", "set_row_options"] },
                "anchorId": { "type": "string", "description": "insert: reference block (paired with position)." },
                "tableAnchorId": { "type": "string", "description": "get_metadata/resolve_cell_coordinate: the table's canonical tbl anchor." },
                "position": { "type": "string", "enum": ["before", "after"], "description": "insert: relative to anchorId. insert_row/insert_column: relative to cellAnchorId." },
                "rows": { "type": "integer" }, "columns": { "type": "integer" },
                "cellContents": { "type": "array", "items": { "type": "string" } },
                "cellAlignment": { "type": "string", "enum": ["left", "center", "right", "justify"] },
                "columnWidths": { "type": "array", "items": { "type": "integer" } },
                "borderless": { "type": "boolean" },
                "cellAnchorId": { "type": "string", "description": "resolve_cell_anchor and every cell mutation: the cell's canonical tc anchor, returned by insert/insert_row/insert_column, get_metadata, resolve_cell_coordinate, or docxodus_search mode=kind query=tc. Legacy p/h/li anchors physically inside a cell are translated temporarily for migration; new callers must use tc." },
                "rowIndex": { "type": "integer", "minimum": 0, "description": "resolve_cell_coordinate: zero-based physical row index." },
                "columnIndex": { "type": "integer", "minimum": 0, "description": "resolve_cell_coordinate: zero-based table-grid column (honors gridBefore/gridAfter and gridSpan)." },
                "markdown": { "type": "string", "description": "replace_cell_content payload." },
                "rowSpan": { "type": "integer", "minimum": 1, "description": "merge_cells: how many rows down the merged rectangle runs from the anchor's cell (default 1). Becomes w:vMerge restart/continue." },
                "colSpan": { "type": "integer", "minimum": 1, "description": "merge_cells: how many cells right the rectangle runs from the anchor's cell (default 1). Becomes w:gridSpan. rowSpan x colSpan must be > 1. unmerge_cells addressed at a vertical continuation unmerges the whole run." },
                "mergeContent": { "type": "string", "enum": ["append", "discard", "reject"], "description": "merge_cells: what to do with the absorbed cells' content — append it to the surviving cell (default, lossless), discard it, or refuse the merge when any absorbed cell is non-empty." },
                "widths": { "type": "array", "items": { "type": "integer" }, "description": "set_column_widths: one positive twip width per column, left→right (1440 = 1 inch). Rewrites w:tblGrid + every cell width and pins the table to fixed layout." },
                "borderScope": { "type": "string", "enum": ["all", "outside", "inside"], "description": "set_borders: which edges to write (default all). Untargeted edges are left unchanged." },
                "borderStyle": { "type": "string", "description": "set_borders: OOXML border style — single (default), double, thick, dotted, dashed, …, or none to remove the targeted edges." },
                "borderSize": { "type": "integer", "description": "set_borders: weight in eighths of a point (default 4 = 0.5pt)." },
                "borderColor": { "type": "string", "description": "set_borders: hex RRGGBB (no '#') or auto (default)." },
                "fill": { "type": "string", "description": "set_shading: hex RRGGBB (leading '#' tolerated) or auto; omit/empty to remove the shading." },
                "shadingScope": { "type": "string", "enum": ["cell", "row"], "description": "set_shading: just the anchor's cell (default) or every cell of its row — header-row banding." },
                "repeat": { "type": "boolean", "description": "set_repeat_header_row/set_row_options: true marks the anchor's row as a repeating header row (w:tblHeader; Word honors it on a run of rows starting at row 1), false unmarks." },
                "allowBreakAcrossPages": { "type": "boolean", "description": "set_row_options: true permits a row to split across pages; false writes w:cantSplit." },
                "heightTwips": { "type": "integer", "minimum": 0, "description": "set_row_options: explicit row height in twips (20 twips = 1pt); zero removes an existing height." },
                "heightRule": { "type": "string", "enum": ["auto", "atLeast", "exact"], "description": "set_row_options: how Word interprets a positive heightTwips value (default atLeast)." }
              },
              "required": ["sessionId", "action"]
            }
            """),
    };
}
