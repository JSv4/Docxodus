#nullable enable

using System.Collections.Generic;

namespace Docxodus.McpServer;

/// <summary>One MCP tool advertised by <c>tools/list</c>. <see cref="InputSchemaJson"/> is a raw
/// JSON Schema object (draft 2020-12 subset: type/properties/required/enum/description) embedded
/// verbatim into the response — see <see cref="Dispatcher"/> for what each action actually does.</summary>
internal sealed record ToolDefinition(string Name, string Description, string InputSchemaJson);

/// <summary>
/// The tool surface this server advertises: three lifecycle tools (open/save/close) plus eleven
/// grouped-intent tools, each accepting an <c>action</c> discriminator and action-specific
/// arguments. See <c>docs/architecture/docx_agent_server.md</c> for the full contract, the
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
                "undoDepth": { "type": "integer", "description": "Bounded undo-ring depth. Default 50." }
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
                "path": { "type": "string", "description": "Destination within the server's document scope, resolved the same way docxodus_open resolves its path. Defaults to the location the session was opened from (overwrite in place)." }
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
                "format": { "type": "string", "enum": ["markdown", "html", "text", "blocks", "info"], "description": "markdown/text: anchor-addressed markdown projection (text strips the markdown syntax). html: fully rendered HTML. blocks: structural metadata for every addressable block. info: section/page-setup facts plus a document edit summary." },
                "anchorId": { "type": "string", "description": "Optional. Scope markdown/html output to one block and its descendants instead of the whole document." }
              },
              "required": ["sessionId", "format"]
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
                "maxResults": { "type": "integer", "description": "Cap the number of matches returned. Default unlimited." }
              },
              "required": ["sessionId", "mode", "query"]
            }
            """),
        new ToolDefinition(
            "docxodus_edit",
            "Insert, replace, delete, or undo/redo text and blocks, addressed by anchor id.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "action": { "type": "string", "enum": ["insert_paragraph", "replace_text", "replace_text_range", "delete_block", "delete_range", "delete_section", "split_paragraph", "merge_paragraphs", "undo", "redo"] },
                "anchorId": { "type": "string", "description": "Target block. Required for every action except delete_range, delete_section, undo, redo." },
                "position": { "type": "string", "enum": ["before", "after"], "description": "insert_paragraph only." },
                "markdown": { "type": "string", "description": "insert_paragraph/replace_text payload, in the supported markdown subset (headings, bullet/ordered lists, bold/italic/code/strike, links, hard breaks)." },
                "find": { "type": "string", "description": "replace_text_range: literal text to find within the block." },
                "replace": { "type": "string", "description": "replace_text_range: replacement text." },
                "caseSensitive": { "type": "boolean", "description": "replace_text_range only. Default false." },
                "toAnchorIdExclusive": { "type": "string", "description": "delete_range: end boundary, exclusive." },
                "fromAnchorId": { "type": "string", "description": "delete_range: start boundary (use anchorId as the field name for the start would be ambiguous with toAnchorIdExclusive; both are required together)." },
                "headingAnchorId": { "type": "string", "description": "delete_section: heading whose section (heading + everything until the next same-or-higher heading) should be removed." },
                "characterOffset": { "type": "integer", "description": "split_paragraph: character offset to split at." },
                "secondAnchorId": { "type": "string", "description": "merge_paragraphs: the paragraph absorbed into anchorId." }
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
                    "indentDelta": { "type": "integer", "description": "Twips to add to the current left indent (negative to outdent)." },
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
                "listFormat": { "type": "string", "enum": ["bullet", "decimal", "none"], "description": "apply_list_format: converts the paragraph into (or out of) a real, auto-numbered Word list." }
              },
              "required": ["sessionId", "action", "anchorId"]
            }
            """),
        new ToolDefinition(
            "docxodus_create",
            "Insert new structural content: paragraphs, headings, tables, horizontal rules, footnotes/endnotes, page-number fields.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "action": { "type": "string", "enum": ["insert_paragraph", "insert_heading", "insert_table", "insert_horizontal_rule", "insert_footnote", "insert_endnote", "insert_page_number_field"] },
                "anchorId": { "type": "string", "description": "Reference block for insert_paragraph/insert_heading/insert_table/insert_horizontal_rule (paired with position), or the citing paragraph for insert_footnote/insert_endnote, or the target paragraph for insert_page_number_field." },
                "position": { "type": "string", "enum": ["before", "after"] },
                "text": { "type": "string", "description": "insert_heading: heading text (plain, not markdown)." },
                "level": { "type": "integer", "minimum": 1, "maximum": 6, "description": "insert_heading: 1-6." },
                "markdown": { "type": "string", "description": "insert_paragraph / insert_footnote / insert_endnote payload." },
                "rows": { "type": "integer" }, "columns": { "type": "integer" },
                "cellContents": { "type": "array", "items": { "type": "string" }, "description": "insert_table: row-major markdown per cell." },
                "cellAlignment": { "type": "string", "enum": ["left", "center", "right", "justify"] },
                "columnWidths": { "type": "array", "items": { "type": "integer" }, "description": "insert_table: twips per column, left to right." },
                "borderless": { "type": "boolean" },
                "ruleStyle": { "type": "string", "enum": ["single", "double", "thick"], "description": "insert_horizontal_rule." },
                "characterOffset": { "type": "integer", "description": "insert_footnote/insert_endnote: character offset within the citing paragraph." },
                "field": { "type": "string", "enum": ["current_page", "total_pages"], "description": "insert_page_number_field." },
                "numberFormat": { "type": "string", "enum": ["decimal", "upperLetter", "lowerLetter", "upperRoman", "lowerRoman"], "description": "insert_page_number_field: optional explicit \\* switch format." }
              },
              "required": ["sessionId", "action"]
            }
            """),
        new ToolDefinition(
            "docxodus_list",
            "Manage list membership: promote a paragraph to a real auto-numbered Word list, renumber/indent, or drop membership.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "action": { "type": "string", "enum": ["apply_format", "set_level", "remove", "get_membership"] },
                "anchorId": { "type": "string" },
                "listFormat": { "type": "string", "enum": ["bullet", "decimal"], "description": "apply_format: creates real w:numPr numbering on the paragraph." },
                "levelDelta": { "type": "integer", "description": "set_level: +1 indents, -1 outdents." }
              },
              "required": ["sessionId", "action", "anchorId"]
            }
            """),
        new ToolDefinition(
            "docxodus_comment",
            "Create and manage native Word review comments (real w:comment markup — visible in Word/Google Docs/LibreOffice's Reviewing pane): comment on a character span of a body paragraph, update a comment's body, remove one, or list them. Comments are addressed by their cmt anchor (from add's created list or the projection's # Comments section). For the semantic highlight/label overlay see docxodus_annotate. Reply threading / resolve state is not yet supported (v2).",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "action": { "type": "string", "enum": ["add", "update", "remove", "list"] },
                "anchorId": { "type": "string", "description": "add: the body paragraph to comment on." },
                "span": { "type": "object", "properties": { "start": { "type": "integer" }, "length": { "type": "integer" } }, "description": "add: character range within the paragraph. Omit to comment on the whole block." },
                "author": { "type": "string", "description": "add: comment author (required)." },
                "initials": { "type": "string", "description": "add: optional author initials." },
                "date": { "type": "string", "description": "add: optional ISO-8601 timestamp; w:date is written only when provided (omitting keeps output deterministic)." },
                "markdown": { "type": "string", "description": "add/update: the comment body (same markdown subset as other payloads)." },
                "commentAnchorId": { "type": "string", "description": "update/remove: the comment definition anchor (kind cmt)." }
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
            "docxodus_track_changes",
            "List, accept, or reject tracked changes (w:ins/w:del/w:moveFrom/w:moveTo/w:rPrChange) already present in the document.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "action": { "type": "string", "enum": ["list", "accept_all", "reject_all"], "description": "Only whole-document accept/reject is supported; there is no selective per-author or per-type accept/reject — see the capability-gap note in docs/architecture/docx_agent_server.md. 'list' accepts author/changeType as client-side display filters." },
                "author": { "type": "string", "description": "list: only return revisions by this author." },
                "changeType": { "type": "string", "enum": ["insert", "delete", "move", "format"], "description": "list: only return revisions of this type." }
              },
              "required": ["sessionId", "action"]
            }
            """),
        new ToolDefinition(
            "docxodus_mutations",
            "Apply (or preview) a batch of docxodus_edit/docxodus_format/docxodus_create/docxodus_table/docxodus_list/docxodus_comment actions as one atomic-feeling sequence, with a single aggregate result.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "mode": { "type": "string", "enum": ["apply", "preview"], "description": "preview applies every step, records the result, then undoes them all before returning — nothing is left changed." },
                "steps": {
                  "type": "array",
                  "items": {
                    "type": "object",
                    "properties": {
                      "tool": { "type": "string", "enum": ["docxodus_edit", "docxodus_format", "docxodus_create", "docxodus_table", "docxodus_list", "docxodus_comment"] },
                      "args": { "type": "object", "description": "The same arguments that tool's action takes, minus sessionId (inherited from the batch)." }
                    },
                    "required": ["tool", "args"]
                  }
                }
              },
              "required": ["sessionId", "mode", "steps"]
            }
            """),
        new ToolDefinition(
            "docxodus_table",
            "Create tables and edit their rows/columns/cell content.",
            """
            {
              "type": "object",
              "properties": {
                "sessionId": { "type": "string" },
                "action": { "type": "string", "enum": ["insert", "insert_row", "insert_column", "delete_row", "delete_column", "replace_cell_content"] },
                "anchorId": { "type": "string", "description": "insert: reference block (paired with position)." },
                "position": { "type": "string", "enum": ["before", "after"], "description": "insert: relative to anchorId. insert_row/insert_column: relative to cellAnchorId." },
                "rows": { "type": "integer" }, "columns": { "type": "integer" },
                "cellContents": { "type": "array", "items": { "type": "string" } },
                "cellAlignment": { "type": "string", "enum": ["left", "center", "right", "justify"] },
                "columnWidths": { "type": "array", "items": { "type": "integer" } },
                "borderless": { "type": "boolean" },
                "cellAnchorId": { "type": "string", "description": "insert_row/insert_column/delete_row/delete_column: a 'p' (paragraph-inside-the-cell) anchor in the target row/column, e.g. from docxodus_table's own insert result or docxodus_search. replace_cell_content: the cell's own 'tc' anchor instead (e.g. from docxodus_search with mode kind, query 'tc') — these two anchor kinds are not interchangeable." },
                "markdown": { "type": "string", "description": "replace_cell_content payload." }
              },
              "required": ["sessionId", "action"]
            }
            """),
    };
}
