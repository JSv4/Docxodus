#nullable enable

using System;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Docxodus;
using Docxodus.Internal;
using Docxodus.Verification;

namespace Docxodus.PyHost;

/// <summary>
/// Op-name → <see cref="DocxSessionOps"/> routing. Each case parses the
/// <c>args</c> JsonElement into the primitive/value-type arguments the Ops
/// facade expects and returns the JSON fragment to embed as the response
/// <c>result</c>.
///
/// Op names mirror the snake_case Python API. The argument keys on the wire
/// are camelCase to match the existing WASM bridge serialization, so the same
/// JSON shapes are interchangeable between TypeScript and Python clients —
/// the Python wrapper normalizes camelCase to snake_case on the decode side.
/// </summary>
internal static class Dispatcher
{
    public static string Dispatch(string op, JsonElement args)
    {
        var preconditions = ParsePreconditions(args);
        if (preconditions is not null && IsMutation(op) && op != "replace_text_range")
        {
            if (op == "undo") return DocxSessionOps.UndoChecked(Handle(args), preconditions);
            if (op == "redo") return DocxSessionOps.RedoChecked(Handle(args), preconditions);

            // The stdio host dispatches one complete request at a time, so the check and
            // mutation below cannot be interleaved by another protocol request.
            var check = DocxSessionOps.CheckPreconditions(Handle(args), preconditions);
            using var parsed = JsonDocument.Parse(check);
            if (!parsed.RootElement.GetProperty("success").GetBoolean()) return check;
        }

        return DispatchCore(op, args);
    }

    private static string DispatchCore(string op, JsonElement args) => op switch
    {
        "ping" => Ping(),
        "open_session" => OpenSession(args),
        "close_session" => CloseSession(args),
        "save" => Save(args),
        "convert_to_html" => ConvertToHtml(args),
        "session_to_html" => SessionToHtml(args),
        "generate_package_manifest" => GeneratePackageManifest(args),
        "get_package_manifest" => VerificationOps.GetPackageManifest(Handle(args)),
        "verify_deliverable" => VerifyDeliverable(args),
        "prove_redline_reversibility" => ProveRedlineReversibility(args),
        "verify_delivery_receipt" => VerifyDeliveryReceipt(args),

        "docx_diff_compare" => DocxDiffCompare(args),
        "docx_diff_compare_products" => DocxDiffCompareProducts(args),
        "docx_diff_compare_batch" => DocxDiffCompareBatch(args),
        "docx_diff_get_revisions" => DocxDiffGetRevisions(args),
        "docx_diff_get_edit_script" => JsonString(DocxDiffGetEditScript(args)),
        "docx_diff_get_semantic_changes" => DocxDiffGetSemanticChanges(args),
        "docx_diff_accept_revisions" => DocxDiffAcceptRevisions(args),
        "docx_diff_reject_revisions" => DocxDiffRejectRevisions(args),
        "docx_diff_consolidate" => DocxDiffConsolidate(args),
        "docx_diff_get_conflicts" => DocxDiffOps.GetConflictsJson(BaseB(args), ReviewersJson(args), DiffSettingsJson(args)),
        "docx_diff_get_consolidated_revisions" => DocxDiffOps.GetConsolidatedRevisionsJson(BaseB(args), ReviewersJson(args), DiffSettingsJson(args)),
        "docx_diff_get_consolidated_edit_script" => JsonString(DocxDiffOps.GetConsolidatedEditScriptJson(BaseB(args), ReviewersJson(args), DiffSettingsJson(args))),
        "project" => DocxSessionOps.Project(Handle(args)),
        "project_anchor" => DocxSessionOps.ProjectAnchor(
            Handle(args), Str(args, "anchorId"),
            (ProjectionDepth)IntOptional(args, "depth", 2),
            DocxSessionJson.ParsePageCitationRequest(args)),
        "get_version" => DocxSessionOps.GetVersionJson(Handle(args)),
        "register_page_map" => DocxSessionOps.RegisterPageMap(
            Handle(args),
            DocxSessionJson.ParsePageMap(JsonObjectElement(args, "pageMap")),
            OptStr(args, "expectedRendererFingerprint")),
        "get_page_map_status" => DocxSessionOps.GetPageMapStatus(
            Handle(args), DocxSessionJson.ParsePageCitationRequest(args)),
        "get_page_citation" => DocxSessionOps.GetPageCitation(
            Handle(args), Str(args, "anchorId"),
            DocxSessionJson.ParsePageCitationRequest(args)
                ?? throw new FormatException("args missing object \"citation\"")),
        "check_preconditions" => DocxSessionOps.CheckPreconditions(Handle(args), ParsePreconditions(args)),
        "execute_batch" => ExecuteBatch(args),
        "preview_batch" => ExecuteBatch(args, preview: true),

        "replace_text" => DocxSessionOps.ReplaceText(Handle(args), Str(args, "anchorId"), Str(args, "markdown")),
        "delete_block" => DocxSessionOps.DeleteBlock(Handle(args), Str(args, "anchorId")),
        "move_block" => DocxSessionOps.MoveBlock(
            Handle(args), Str(args, "sourceAnchorId"), Str(args, "targetAnchorId"),
            DocxSessionJson.ParsePos(Str(args, "position"))),
        "delete_range" => DocxSessionOps.DeleteRange(
            Handle(args), Str(args, "fromAnchorId"), Str(args, "toAnchorIdExclusive")),
        "delete_section" => DocxSessionOps.DeleteSection(
            Handle(args), Str(args, "headingAnchorId")),
        "replace_text_range" => DocxSessionOps.ReplaceTextRange(
            Handle(args), Str(args, "anchorId"), Str(args, "find"), Str(args, "replace"), ParseReplaceOptions(args)),
        "replace_text_at_span" => DocxSessionOps.ReplaceTextAtSpan(
            Handle(args), Str(args, "anchorId"), Int(args, "spanStart"), Int(args, "spanLength"), Str(args, "replace")),
        "replace_inner" => DocxSessionOps.ReplaceInner(
            Handle(args), Str(args, "matchText"), Str(args, "anchorId"),
            Int(args, "spanStart"), Int(args, "spanLength"), Str(args, "newInner")),

        "insert_paragraph" => DocxSessionOps.InsertParagraph(
            Handle(args), Str(args, "anchorId"), ParsePos(args, "position"), Str(args, "markdown")),
        "split_paragraph" => DocxSessionOps.SplitParagraph(
            Handle(args), Str(args, "anchorId"), Int(args, "characterOffset")),
        "merge_paragraphs" => DocxSessionOps.MergeParagraphs(
            Handle(args), Str(args, "firstAnchorId"), Str(args, "secondAnchorId")),

        "set_header_text" => DocxSessionOps.SetHeaderText(
            Handle(args), Str(args, "anchorId"),
            DocxSessionJson.ParseHeaderFooterKind(Str(args, "kind")), Str(args, "markdown")),
        "set_footer_text" => DocxSessionOps.SetFooterText(
            Handle(args), Str(args, "anchorId"),
            DocxSessionJson.ParseHeaderFooterKind(Str(args, "kind")), Str(args, "markdown")),
        "insert_page_number_field" => DocxSessionOps.InsertPageNumberField(
            Handle(args), Str(args, "anchorId"),
            DocxSessionJson.ParsePageNumberField(Str(args, "field")),
            DocxSessionJson.ParseNumberFormatOrNull(OptStr(args, "format"))),
        "ensure_header_footer_visible" => DocxSessionOps.EnsureHeaderFooterVisible(
            Handle(args), Str(args, "anchorId"),
            DocxSessionJson.ParseHeaderFooterKind(Str(args, "kind"))),
        "set_page_numbering" => DocxSessionOps.SetPageNumbering(
            Handle(args), Str(args, "anchorId"), ParsePageNumberingOp(args, "op")),
        "clear_page_numbering" => DocxSessionOps.ClearPageNumbering(
            Handle(args), Str(args, "anchorId")),
        "set_header_footer_kind_enabled" => DocxSessionOps.SetHeaderFooterKindEnabled(
            Handle(args), Str(args, "anchorId"), Str(args, "kind"),
            OptBool(args, "enabled") ?? throw new FormatException("args missing boolean \"enabled\"")),
        "set_page_setup" => DocxSessionOps.SetPageSetup(
            Handle(args), Str(args, "anchorId"), ParsePageSetupOp(args, "op")),

        // Reference fields (issue #607): the switches never cross the wire — typed options only.
        "insert_table_of_contents" => DocxSessionOps.InsertTableOfContents(
            Handle(args), Str(args, "anchorId"), DocxSessionJson.ParsePos(OptStr(args, "position") ?? "before"),
            OptionsOrNull(args, DocxSessionJson.ParseTableOfContentsOptions)),
        "insert_table_of_figures" => DocxSessionOps.InsertTableOfFigures(
            Handle(args), Str(args, "anchorId"), DocxSessionJson.ParsePos(OptStr(args, "position") ?? "before"),
            OptionsOrNull(args, DocxSessionJson.ParseTableOfFiguresOptions)),
        "insert_table_of_authorities" => DocxSessionOps.InsertTableOfAuthorities(
            Handle(args), Str(args, "anchorId"), DocxSessionJson.ParsePos(OptStr(args, "position") ?? "before"),
            OptionsOrNull(args, DocxSessionJson.ParseTableOfAuthoritiesOptions)),

        "insert_footnote" => DocxSessionOps.InsertFootnote(
            Handle(args), Str(args, "anchorId"), Int(args, "characterOffset"), Str(args, "markdown")),
        "insert_endnote" => DocxSessionOps.InsertEndnote(
            Handle(args), Str(args, "anchorId"), Int(args, "characterOffset"), Str(args, "markdown")),
        "insert_cross_reference" => DocxSessionOps.InsertCrossReference(
            Handle(args), Str(args, "anchorId"), Int(args, "characterOffset"),
            Str(args, "bookmarkName"), ParseCrossReferenceOptions(args)),

        "add_comment" => AddComment(args),
        "add_comment_reply" => DocxSessionOps.AddCommentReply(
            Handle(args), Str(args, "parentAnchorId"), Str(args, "author"),
            OptStr(args, "initials"), OptStr(args, "date"), Str(args, "markdown")),
        "update_comment" => DocxSessionOps.UpdateComment(
            Handle(args), Str(args, "anchorId"), Str(args, "markdown")),
        "set_comment_resolved" => DocxSessionOps.SetCommentResolved(
            Handle(args), Str(args, "anchorId"),
            OptBool(args, "resolved") ?? throw new FormatException("args missing boolean \"resolved\"")),
        "remove_comment" => DocxSessionOps.RemoveComment(
            Handle(args), Str(args, "anchorId")),
        "list_comments" => DocxSessionOps.ListComments(Handle(args)),

        "list_hyperlinks" => DocxSessionOps.ListHyperlinks(
            Handle(args), (ProjectionScopes)IntOptional(args, "scopes", (int)ProjectionScopes.All)),
        "add_hyperlink" => DocxSessionOps.AddHyperlink(
            Handle(args), Str(args, "anchorId"), Int(args, "start"), Int(args, "length"),
            Str(args, "kind"), Str(args, "target")),
        "update_hyperlink" => DocxSessionOps.UpdateHyperlink(
            Handle(args), Str(args, "hyperlinkId"), Str(args, "kind"), Str(args, "target")),
        "remove_hyperlink" => DocxSessionOps.RemoveHyperlink(Handle(args), Str(args, "hyperlinkId")),
        "get_image_capabilities" => DocxSessionOps.GetImageCapabilities(),
        "list_images" => DocxSessionOps.ListImages(
            Handle(args), (ProjectionScopes)IntOptional(args, "scopes", (int)ProjectionScopes.All)),
        "insert_image" => DocxSessionOps.InsertImage(
            Handle(args), Str(args, "anchorId"), Int(args, "characterOffset"),
            Str(args, "imageBase64"), JsonObjectOrEmpty(args, "options")),
        "replace_image" => DocxSessionOps.ReplaceImage(
            Handle(args), Str(args, "imageId"), Str(args, "imageBase64")),
        "set_image_dimensions" => DocxSessionOps.SetImageDimensions(
            Handle(args), Str(args, "imageId"), JsonObjectOrEmpty(args, "dimensions")),
        "set_image_metadata" => DocxSessionOps.SetImageMetadata(
            Handle(args), Str(args, "imageId"), OptStr(args, "altText"), OptStr(args, "title")),
        "set_image_floating_layout" => DocxSessionOps.SetImageFloatingLayout(
            Handle(args), Str(args, "imageId"), JsonObject(args, "layout")),
        "remove_image" => DocxSessionOps.RemoveImage(Handle(args), Str(args, "imageId")),
        "list_content_controls" => DocxSessionOps.ListContentControls(
            Handle(args), (ProjectionScopes)IntOptional(args, "scopes", (int)ProjectionScopes.All)),
        "fill_content_control_text" => DocxSessionOps.FillContentControlText(
            Handle(args), Str(args, "anchorId"), Str(args, "text"),
            JsonObjectOrEmpty(args, "options")),
        "fill_content_control_rich_text" => DocxSessionOps.FillContentControlRichText(
            Handle(args), Str(args, "anchorId"), Str(args, "markdown"),
            JsonObjectOrEmpty(args, "options")),
        "set_content_control_checked" => DocxSessionOps.SetContentControlChecked(
            Handle(args), Str(args, "anchorId"),
            OptBool(args, "checked") ?? throw new FormatException("args missing boolean \"checked\""),
            JsonObjectOrEmpty(args, "options")),
        "set_content_control_date" => DocxSessionOps.SetContentControlDate(
            Handle(args), Str(args, "anchorId"), Str(args, "value"),
            OptStr(args, "displayText"), JsonObjectOrEmpty(args, "options")),
        "select_content_control_item" => DocxSessionOps.SelectContentControlItem(
            Handle(args), Str(args, "anchorId"), Str(args, "value"),
            JsonObjectOrEmpty(args, "options")),
        "fill_content_control_picture" => DocxSessionOps.FillContentControlPicture(
            Handle(args), Str(args, "anchorId"), Str(args, "imageBase64"),
            JsonObjectOrEmpty(args, "options")),
        "add_repeating_section_item" => DocxSessionOps.AddRepeatingSectionItem(
            Handle(args), Str(args, "sectionAnchorId"), OptStr(args, "afterItemAnchorId"),
            JsonObjectOrEmpty(args, "options")),
        "remove_repeating_section_item" => DocxSessionOps.RemoveRepeatingSectionItem(
            Handle(args), Str(args, "itemAnchorId")),
        "list_bookmarks" => DocxSessionOps.ListBookmarks(
            Handle(args), (ProjectionScopes)IntOptional(args, "scopes", (int)ProjectionScopes.All)),
        "add_bookmark" => DocxSessionOps.AddBookmark(
            Handle(args), Str(args, "name"), Str(args, "startAnchorId"), Int(args, "startOffset"),
            Str(args, "endAnchorId"), Int(args, "endOffset")),
        "rename_bookmark" => DocxSessionOps.RenameBookmark(
            Handle(args), Str(args, "name"), Str(args, "newName")),
        "move_bookmark" => DocxSessionOps.MoveBookmark(
            Handle(args), Str(args, "name"), Str(args, "startAnchorId"), Int(args, "startOffset"),
            Str(args, "endAnchorId"), Int(args, "endOffset")),
        "remove_bookmark" => DocxSessionOps.RemoveBookmark(Handle(args), Str(args, "name")),

        "list_revisions" => DocxSessionOps.ListRevisions(Handle(args)),
        "accept_revision" => DocxSessionOps.AcceptRevision(Handle(args), Str(args, "revisionId")),
        "reject_revision" => DocxSessionOps.RejectRevision(Handle(args), Str(args, "revisionId")),
        "accept_all_revisions" => DocxSessionOps.AcceptAllRevisions(Handle(args)),
        "reject_all_revisions" => DocxSessionOps.RejectAllRevisions(Handle(args)),

        "apply_format" => DocxSessionOps.ApplyFormat(
            Handle(args), Str(args, "anchorId"), ParseOptionalSpan(args, "span"), ParseFormatOp(args, "op")),
        "apply_format_by_substring" => DocxSessionOps.ApplyFormatBySubstring(
            Handle(args), Str(args, "anchorId"), Str(args, "substring"), ParseFormatOp(args, "op")),
        "set_paragraph_style" => DocxSessionOps.SetParagraphStyle(
            Handle(args), Str(args, "anchorId"), Str(args, "styleId")),
        "set_paragraph_format" => DocxSessionOps.SetParagraphFormat(
            Handle(args), Str(args, "anchorId"), ParseParagraphFormatOp(args, "op")),
        "set_list_level" => DocxSessionOps.SetListLevel(
            Handle(args), Str(args, "anchorId"), Int(args, "levelDelta")),
        "remove_list_membership" => DocxSessionOps.RemoveListMembership(
            Handle(args), Str(args, "anchorId")),
        "apply_list_format" => DocxSessionOps.ApplyListFormat(
            Handle(args), Str(args, "anchorId"), DocxSessionJson.ParseListFormat(OptStr(args, "listFormat"))),
        "apply_list_format_range" => DocxSessionOps.ApplyListFormatRange(
            Handle(args), Str(args, "firstAnchorId"), Str(args, "lastAnchorId"),
            DocxSessionJson.ParseListFormat(OptStr(args, "listFormat"))),
        "set_list_start_override" => DocxSessionOps.SetListStartOverride(
            Handle(args), Str(args, "anchorId"), Int(args, "value")),
        "clear_list_start_override" => DocxSessionOps.ClearListStartOverride(
            Handle(args), Str(args, "anchorId")),

        "get_table_metadata" => DocxSessionOps.GetTableMetadata(
            Handle(args), Str(args, "tableAnchorId")),
        "resolve_table_cell_anchor" => DocxSessionOps.ResolveTableCellAnchor(
            Handle(args), Str(args, "cellAnchorId")),
        "resolve_table_cell_coordinate" => DocxSessionOps.ResolveTableCellCoordinate(
            Handle(args), Str(args, "tableAnchorId"), Int(args, "rowIndex"), Int(args, "columnIndex")),
        "insert_table" => DocxSessionOps.InsertTable(
            Handle(args), Str(args, "anchorId"), ParsePos(args, "position"),
            Int(args, "rows"), Int(args, "columns"), RawObjectOrEmpty(args, "options")),
        "insert_table_row" => DocxSessionOps.InsertTableRow(
            Handle(args), Str(args, "cellAnchorId"), ParsePos(args, "position")),
        "insert_table_column" => DocxSessionOps.InsertTableColumn(
            Handle(args), Str(args, "cellAnchorId"), ParsePos(args, "position")),
        "delete_table_row" => DocxSessionOps.DeleteTableRow(Handle(args), Str(args, "cellAnchorId")),
        "delete_table_column" => DocxSessionOps.DeleteTableColumn(Handle(args), Str(args, "cellAnchorId")),
        "merge_cells" => DocxSessionOps.MergeCells(
            Handle(args), Str(args, "cellAnchorId"), Int(args, "rowSpan"), Int(args, "columnSpan"),
            OptStr(args, "content")),
        "unmerge_cells" => DocxSessionOps.UnmergeCells(Handle(args), Str(args, "cellAnchorId")),
        "set_column_widths" => DocxSessionOps.SetColumnWidths(
            Handle(args), Str(args, "cellAnchorId"), RawArray(args, "widths")),
        "set_table_borders" => DocxSessionOps.SetTableBorders(
            Handle(args), Str(args, "cellAnchorId"), RawObjectOrEmpty(args, "spec")),
        "set_cell_shading" => DocxSessionOps.SetCellShading(
            Handle(args), Str(args, "cellAnchorId"), OptStr(args, "fill") ?? "",
            OptStr(args, "scope") ?? "cell"),
        "set_repeat_header_row" => DocxSessionOps.SetRepeatHeaderRow(
            Handle(args), Str(args, "cellAnchorId"), OptBool(args, "repeat") ?? true),
        "set_table_row_options" => DocxSessionOps.SetTableRowOptions(
            Handle(args), Str(args, "cellAnchorId"), OptBool(args, "repeatHeader"),
            OptBool(args, "allowBreakAcrossPages"), OptInt(args, "heightTwips"),
            OptStr(args, "heightRule")),
        "replace_cell_content" => DocxSessionOps.ReplaceCellContent(
            Handle(args), Str(args, "cellAnchorId"), Str(args, "markdown")),

        "raw_get_xml" => JsonString(DocxSessionOps.RawGetXml(Handle(args), Str(args, "anchorId"))),
        "raw_insert_xml" => DocxSessionOps.RawInsertXml(
            Handle(args), Str(args, "anchorId"), ParsePos(args, "position"), Str(args, "xml")),
        "raw_replace_xml" => DocxSessionOps.RawReplaceXml(
            Handle(args), Str(args, "anchorId"), Str(args, "xml")),

        "grep" => Grep(args, crossBlock: false),
        "grep_cross_block" => Grep(args, crossBlock: true),

        "find_placeholders" => DocxSessionOps.FindPlaceholders(
            Handle(args),
            (PlaceholderKinds)IntOptional(args, "kinds", (int)PlaceholderKinds.All),
            (ProjectionScopes)IntOptional(args, "scope", (int)ProjectionScopes.Body),
            IntOptional(args, "contextChars", 80),
            (ContextBoundary)IntOptional(args, "boundary", (int)ContextBoundary.Char),
            DocxSessionJson.ParsePageCitationRequest(args)),
        "get_edit_summary" => DocxSessionOps.GetEditSummary(Handle(args)),
        "remaining_placeholders" => DocxSessionOps.RemainingPlaceholders(
            Handle(args), (PlaceholderKinds)IntOptional(args, "kinds", 7)),
        "get_diff" => DocxSessionOps.GetDiff(
            Handle(args), (DiffFormat)IntOptional(args, "format", 0)),
        "get_semantic_changes" => DocxSessionOps.GetSemanticChanges(Handle(args)),
        "find_by_annotation" => DocxSessionOps.FindByAnnotation(
            Handle(args), Str(args, "annotationId"), DocxSessionJson.ParsePageCitationRequest(args)),
        "find_by_label" => DocxSessionOps.FindByLabel(
            Handle(args), Str(args, "labelId"), DocxSessionJson.ParsePageCitationRequest(args)),
        "find_by_bookmark" => DocxSessionOps.FindByBookmark(
            Handle(args), Str(args, "bookmarkName"), DocxSessionJson.ParsePageCitationRequest(args)),
        "list_annotations" => DocxSessionOps.ListAnnotations(Handle(args)),
        "add_annotation" => DocxSessionOps.AddAnnotation(
            Handle(args),
            Str(args, "anchorId"),
            ParseOptionalSpan(args, "span"),
            JsonObject(args, "annotation")),
        "remove_annotation" => DocxSessionOps.RemoveAnnotation(
            Handle(args), Str(args, "annotationId")),
        "update_annotation" => DocxSessionOps.UpdateAnnotation(
            Handle(args), Str(args, "annotationId"), JsonObject(args, "update")),
        "move_annotation" => DocxSessionOps.MoveAnnotation(
            Handle(args),
            Str(args, "annotationId"),
            Str(args, "newAnchorId"),
            ParseOptionalSpan(args, "newSpan")),

        "exists" => DocxSessionOps.Exists(Handle(args), Str(args, "anchorId")) ? "true" : "false",
        "get_anchor_info" => DocxSessionOps.GetAnchorInfo(Handle(args), Str(args, "anchorId")),
        "get_anchor_infos" => DocxSessionOps.GetAnchorInfos(Handle(args), ParseAnchorIdArray(args)),
        "get_block_metadata" => DocxSessionOps.GetBlockMetadata(Handle(args), Str(args, "anchorId")),
        "get_block_metadatas" => DocxSessionOps.GetBlockMetadatas(Handle(args), ParseAnchorIdArray(args)),
        "get_list_membership" => DocxSessionOps.GetListMembership(Handle(args), Str(args, "anchorId")),
        "get_section_info" => DocxSessionOps.GetSectionInfo(Handle(args), Str(args, "anchorId")),
        "list_styles" => DocxSessionOps.ListStyles(Handle(args)),
        "get_formatting" => DocxSessionOps.GetFormatting(Handle(args), Str(args, "anchorId")),
        "list_inline_spans" => DocxSessionOps.ListInlineSpans(Handle(args), Str(args, "anchorId")),
        "find_by_text" => DocxSessionOps.FindByText(Handle(args), Str(args, "needle"), ParseFindOptions(args)),
        "find_all_by_text" => DocxSessionOps.FindAllByText(Handle(args), Str(args, "needle"), ParseFindOptions(args)),
        "find_by_regex" => DocxSessionOps.FindByRegex(
            Handle(args), Str(args, "pattern"),
            (RegexOptions)IntOptional(args, "regexOptions", 0),
            ParseFindOptions(args)),
        "find_by_kind" => DocxSessionOps.FindByKind(
            Handle(args), Str(args, "kind"),
            args.ValueKind == JsonValueKind.Object && args.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.String
                ? sc.GetString() : null,
            DocxSessionJson.ParsePageCitationRequest(args)),

        "undo" => DocxSessionOps.Undo(Handle(args)) ? "true" : "false",
        "redo" => DocxSessionOps.Redo(Handle(args)) ? "true" : "false",

        "set_tracked_changes" => SetTrackedChanges(args),
        "set_revision_author" => SetRevisionAuthor(args),

        _ => throw new UnknownOpException(op),
    };

    /// <summary>
    /// Prove that a redline's generated revisions accept to the intended final and reject to the
    /// baseline. All three packages are required: unlike verify_deliverable there is no
    /// session-scoped form, because the proof compares three distinct packages rather than a
    /// live document against its own opening bytes.
    /// </summary>
    private static string ProveRedlineReversibility(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object)
            throw new FormatException("prove_redline_reversibility args must be an object");

        return VerificationOps.ProveRedlineReversibility(
            RequiredPackage(args, "baselineB64"),
            RequiredPackage(args, "intendedFinalB64"),
            RequiredPackage(args, "redlineB64"));
    }

    private static byte[] RequiredPackage(JsonElement args, string property)
    {
        if (!args.TryGetProperty(property, out var encoded)
            || encoded.ValueKind != JsonValueKind.String)
            throw new FormatException($"args property \"{property}\" must be a string");

        return Convert.FromBase64String(Str(args, property));
    }

    /// <summary>Portable delivery-receipt verification (issue #520): receiptJson plus an
    /// optional artifactsB64 object of {artifactId: base64} — the exact wire shape
    /// <see cref="DeliveryOps.VerifyChangeReceiptJson"/> owns, passed through verbatim.</summary>
    private static string VerifyDeliveryReceipt(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object)
            throw new FormatException("verify_delivery_receipt args must be an object");
        var receiptJson = Str(args, "receiptJson");
        string? artifactsJson = null;
        if (args.TryGetProperty("artifactsB64", out var artifacts)
            && artifacts.ValueKind != JsonValueKind.Null)
        {
            if (artifacts.ValueKind != JsonValueKind.Object)
                throw new FormatException("args property \"artifactsB64\" must be an object");
            artifactsJson = artifacts.GetRawText();
        }

        return DeliveryOps.VerifyChangeReceiptJson(receiptJson, artifactsJson);
    }

    private static string VerifyDeliverable(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object)
            throw new FormatException("verify_deliverable args must be an object");

        if (args.TryGetProperty("docxB64", out var encoded))
        {
            if (encoded.ValueKind != JsonValueKind.String)
                throw new FormatException("args property \"docxB64\" must be a string");

            var packageBytes = Convert.FromBase64String(Str(args, "docxB64"));
            if (!args.TryGetProperty("baselineB64", out var baseline))
                return VerificationOps.VerifyDeliverable(packageBytes);
            if (baseline.ValueKind != JsonValueKind.String)
                throw new FormatException("args property \"baselineB64\" must be a string");

            return VerificationOps.VerifyDeliverable(
                Convert.FromBase64String(Str(args, "baselineB64")),
                packageBytes);
        }

        if (args.TryGetProperty("baselineB64", out _))
            throw new FormatException("args property \"baselineB64\" requires string \"docxB64\"");

        return DocxSessionOps.VerifyDeliverable(Handle(args));
    }

    private static string Ping()
    {
        var version = typeof(DocxSession).Assembly.GetName().Version?.ToString() ?? "0.0.0";
        var sb = new System.Text.StringBuilder(96);
        sb.Append("{\"pong\":true,\"version\":");
        sb.Append(DocxSessionJson.JsonString(version));
        sb.Append(",\"dotnet\":");
        sb.Append(DocxSessionJson.JsonString(Environment.Version.ToString()));
        sb.Append(",\"sessions\":");
        sb.Append(SessionRegistry.Count);
        sb.Append('}');
        return sb.ToString();
    }

    private static string OpenSession(JsonElement args)
    {
        var b64 = Str(args, "docxB64");
        var bytes = Convert.FromBase64String(b64);
        DocxSessionSettings? settings = null;
        if (args.TryGetProperty("settings", out var s) && s.ValueKind == JsonValueKind.Object)
        {
            settings = DocxSessionJson.ParseSettings(s.GetRawText());
        }
        var handle = DocxSessionOps.OpenSession(bytes, settings);
        return handle.ToString(CultureInfo.InvariantCulture);
    }

    private static string CloseSession(JsonElement args)
    {
        DocxSessionOps.CloseSession(Handle(args));
        return "null";
    }

    private static string SetTrackedChanges(JsonElement args)
    {
        DocxSessionOps.SetTrackedChanges(Handle(args),
            DocxSessionJson.ParseTrackedChangeMode(Str(args, "mode")));
        return "null";
    }

    private static string SetRevisionAuthor(JsonElement args)
    {
        string? author = args.TryGetProperty("author", out var a) && a.ValueKind == JsonValueKind.String
            ? a.GetString()
            : null;
        DocxSessionOps.SetRevisionAuthor(Handle(args), author);
        return "null";
    }

    private static string Save(JsonElement args)
    {
        // Tri-state: absent → the session's open-time persistAnchorIds setting; explicit
        // true/false overrides it for this save only (mirrors docxodus-mcp's docxodus_save).
        var bytes = OptBool(args, "persistAnchorIds") is { } persist
            ? DocxSessionOps.Save(Handle(args), persist)
            : DocxSessionOps.Save(Handle(args));
        return "{\"docxB64\":" + DocxSessionJson.JsonString(Convert.ToBase64String(bytes)) + "}";
    }

    private static string ConvertToHtml(JsonElement args)
    {
        var bytes = Convert.FromBase64String(Str(args, "docxB64"));
        var html = HtmlConversionOps.ConvertToHtml(bytes, ParseHtmlOptions(args));
        return JsonString(html);
    }

    private static string SessionToHtml(JsonElement args)
    {
        var html = HtmlConversionOps.ConvertToHtml(Handle(args), ParseHtmlOptions(args));
        return JsonString(html);
    }

    // ─── DocxDiff (IR diff engine) — stateless byte-in ops ──────────────

    private static string DocxDiffCompare(JsonElement args)
    {
        var left = Convert.FromBase64String(Str(args, "leftB64"));
        var right = Convert.FromBase64String(Str(args, "rightB64"));
        var bytes = DocxDiffOps.Compare(left, right, DiffSettingsJson(args));
        return "{\"docxB64\":" + DocxSessionJson.JsonString(Convert.ToBase64String(bytes)) + "}";
    }

    private static string DocxDiffGetRevisions(JsonElement args)
    {
        var left = Convert.FromBase64String(Str(args, "leftB64"));
        var right = Convert.FromBase64String(Str(args, "rightB64"));
        // Already a JSON object ({"revisions":[…]}) — embed verbatim as the result.
        return DocxDiffOps.GetRevisionsJson(left, right, DiffSettingsJson(args));
    }

    private static string DocxDiffCompareProducts(JsonElement args)
    {
        var left = Convert.FromBase64String(Str(args, "leftB64"));
        var right = Convert.FromBase64String(Str(args, "rightB64"));
        var products = args.TryGetProperty("products", out var selected)
            && selected.ValueKind == JsonValueKind.Array
            ? selected.GetRawText()
            : null;
        // One memoized comparison pass serving every requested product (issue #594);
        // already the complete JSON envelope — embed verbatim as the result.
        return DocxDiffOps.CompareProductsJson(left, right, DiffSettingsJson(args), products);
    }

    /// <summary>
    /// One baseline against many candidates, reading the baseline once (issue #617). The candidates
    /// arrive as <c>[{"name":…,"docB64":…}]</c>, and the reply is already the complete envelope.
    /// </summary>
    private static string DocxDiffCompareBatch(JsonElement args)
    {
        var baseline = Convert.FromBase64String(Str(args, "baselineB64"));
        var candidates = args.TryGetProperty("candidates", out var list)
            && list.ValueKind == JsonValueKind.Array
            ? list.GetRawText()
            : "[]";
        var products = args.TryGetProperty("products", out var selected)
            && selected.ValueKind == JsonValueKind.Array
            ? selected.GetRawText()
            : null;
        return DocxDiffOps.CompareBatchJson(baseline, candidates, DiffSettingsJson(args), products);
    }

    private static string DocxDiffGetEditScript(JsonElement args)
    {
        var left = Convert.FromBase64String(Str(args, "leftB64"));
        var right = Convert.FromBase64String(Str(args, "rightB64"));
        return DocxDiffOps.GetEditScriptJson(left, right, DiffSettingsJson(args));
    }

    private static string DocxDiffGetSemanticChanges(JsonElement args)
    {
        var left = Convert.FromBase64String(Str(args, "leftB64"));
        var right = Convert.FromBase64String(Str(args, "rightB64"));
        // Already the public semantic-change JSON object — embed verbatim in the host result.
        return DocxDiffOps.GetSemanticChangesJson(left, right, DiffSettingsJson(args));
    }

    private static string DocxDiffAcceptRevisions(JsonElement args)
    {
        var bytes = DocxDiffOps.AcceptRevisions(Convert.FromBase64String(Str(args, "docxB64")));
        return "{\"docxB64\":" + DocxSessionJson.JsonString(Convert.ToBase64String(bytes)) + "}";
    }

    private static string DocxDiffRejectRevisions(JsonElement args)
    {
        var bytes = DocxDiffOps.RejectRevisions(Convert.FromBase64String(Str(args, "docxB64")));
        return "{\"docxB64\":" + DocxSessionJson.JsonString(Convert.ToBase64String(bytes)) + "}";
    }

    private static string DocxDiffConsolidate(JsonElement args)
    {
        var bytes = DocxDiffOps.Consolidate(BaseB(args), ReviewersJson(args), DiffSettingsJson(args));
        return "{\"docxB64\":" + DocxSessionJson.JsonString(Convert.ToBase64String(bytes)) + "}";
    }

    private static byte[] BaseB(JsonElement args) => Convert.FromBase64String(Str(args, "baseB64"));

    private static string ReviewersJson(JsonElement args) =>
        args.TryGetProperty("reviewers", out var r) && r.ValueKind == JsonValueKind.Array ? r.GetRawText() : "[]";

    private static string? DiffSettingsJson(JsonElement args) =>
        args.ValueKind == JsonValueKind.Object
        && args.TryGetProperty("settings", out var s)
        && s.ValueKind == JsonValueKind.Object
            ? s.GetRawText()
            : null;

    private static HtmlConversionOptions ParseHtmlOptions(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object
            || !args.TryGetProperty("options", out var o)
            || o.ValueKind != JsonValueKind.Object)
        {
            return new HtmlConversionOptions();
        }

        string StrOpt(string name, string fallback) =>
            o.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String
                ? v.GetString()! : fallback;
        int IntOpt(string name, int fallback) =>
            o.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
                ? v.GetInt32() : fallback;
        double DblOpt(string name, double fallback) =>
            o.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
                ? v.GetDouble() : fallback;
        bool BoolOpt(string name, bool fallback) =>
            o.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False)
                ? v.GetBoolean() : fallback;
        string? StrOptNullable(string name, string? fallback) =>
            o.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String
                ? v.GetString() : fallback;

        var defaults = new HtmlConversionOptions();
        return new HtmlConversionOptions
        {
            PageTitle = StrOpt("pageTitle", defaults.PageTitle),
            CssClassPrefix = StrOpt("cssClassPrefix", defaults.CssClassPrefix),
            FabricateCssClasses = BoolOpt("fabricateCssClasses", defaults.FabricateCssClasses),
            AdditionalCss = StrOpt("additionalCss", defaults.AdditionalCss),
            CommentRenderMode = IntOpt("commentRenderMode", defaults.CommentRenderMode),
            CommentCssClassPrefix = StrOpt("commentCssClassPrefix", defaults.CommentCssClassPrefix),
            PaginationMode = IntOpt("paginationMode", defaults.PaginationMode),
            PaginationScale = DblOpt("paginationScale", defaults.PaginationScale),
            PaginationCssClassPrefix = StrOpt("paginationCssClassPrefix", defaults.PaginationCssClassPrefix),
            RenderAnnotations = BoolOpt("renderAnnotations", defaults.RenderAnnotations),
            AnnotationLabelMode = IntOpt("annotationLabelMode", defaults.AnnotationLabelMode),
            AnnotationCssClassPrefix = StrOpt("annotationCssClassPrefix", defaults.AnnotationCssClassPrefix),
            RenderFootnotesAndEndnotes = BoolOpt("renderFootnotesAndEndnotes", defaults.RenderFootnotesAndEndnotes),
            RenderHeadersAndFooters = BoolOpt("renderHeadersAndFooters", defaults.RenderHeadersAndFooters),
            RenderTrackedChanges = BoolOpt("renderTrackedChanges", defaults.RenderTrackedChanges),
            ShowDeletedContent = BoolOpt("showDeletedContent", defaults.ShowDeletedContent),
            RenderMoveOperations = BoolOpt("renderMoveOperations", defaults.RenderMoveOperations),
            RenderUnsupportedContentPlaceholders = BoolOpt("renderUnsupportedContentPlaceholders", defaults.RenderUnsupportedContentPlaceholders),
            DocumentLanguage = StrOptNullable("documentLanguage", defaults.DocumentLanguage),
        };
    }

    private static string Grep(JsonElement args, bool crossBlock)
    {
        var pattern = Str(args, "pattern");
        var regexOpts = (RegexOptions)IntOptional(args, "regexOptions", 0);
        var scope = (ProjectionScopes)IntOptional(args, "scope", (int)ProjectionScopes.Body);
        var contextChars = IntOptional(args, "contextChars", 80);
        var whitespace = (WhitespaceMode)IntOptional(args, "whitespace", (int)WhitespaceMode.Preserve);
        var boundary = (ContextBoundary)IntOptional(args, "boundary", (int)ContextBoundary.Char);
        var citation = DocxSessionJson.ParsePageCitationRequest(args);
        return crossBlock
            ? DocxSessionOps.GrepCrossBlock(Handle(args), pattern, regexOpts, scope, contextChars, whitespace, boundary, citation)
            : DocxSessionOps.Grep(Handle(args), pattern, regexOpts, scope, contextChars, whitespace, boundary, citation);
    }

    private static string AddComment(JsonElement args)
    {
        var anchorId = OptStr(args, "anchorId");
        var revisionId = OptStr(args, "revisionId");
        var hasSpan = args.ValueKind == JsonValueKind.Object && args.TryGetProperty("span", out _);
        if ((anchorId is null) == (revisionId is null) || (revisionId is not null && hasSpan))
            throw new FormatException(
                "add_comment requires exactly one target: anchorId (with optional span) or revisionId");

        return revisionId is not null
            ? DocxSessionOps.AddCommentToRevision(
                Handle(args), revisionId, Str(args, "author"), OptStr(args, "initials"),
                OptStr(args, "date"), Str(args, "markdown"))
            : DocxSessionOps.AddComment(
                Handle(args), anchorId!, ParseOptionalSpan(args, "span"), Str(args, "author"),
                OptStr(args, "initials"), OptStr(args, "date"), Str(args, "markdown"));
    }

    // ─── Arg helpers ────────────────────────────────────────────────────

    private static int Handle(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty("handle", out var h) || h.ValueKind != JsonValueKind.Number)
            throw new FormatException("args missing numeric \"handle\"");
        return h.GetInt32();
    }

    /// <summary>
    /// Generate a manifest, optionally under the caller's lowered inspection ceilings so a
    /// constrained caller limits the inspection itself rather than rejecting after expansion.
    /// </summary>
    private static string GeneratePackageManifest(JsonElement args)
    {
        var bytes = Convert.FromBase64String(Str(args, "docxB64"));
        if (args.ValueKind != JsonValueKind.Object
            || !args.TryGetProperty("limits", out var limits)
            || limits.ValueKind != JsonValueKind.Object)
        {
            return VerificationOps.GeneratePackageManifest(bytes);
        }

        var defaults = new PackageManifestOptions();
        return VerificationOps.GeneratePackageManifest(bytes, new PackageManifestOptions
        {
            MaxEntryCount = IntOptional(limits, "opcEntries", defaults.MaxEntryCount),
            MaxTotalUncompressedBytes = LongOptional(
                limits, "expandedOpcBytes", defaults.MaxTotalUncompressedBytes),
            MaxXmlPartBytes = LongOptional(limits, "xmlPartBytes", defaults.MaxXmlPartBytes),
            MaxUriLength = IntOptional(limits, "opcUriCharacters", defaults.MaxUriLength),
            MaxCompressionRatio = DoubleOptional(
                limits, "opcCompressionRatio", defaults.MaxCompressionRatio),
        });
    }

    private static long LongOptional(JsonElement args, string name, long fallback)
    {
        if (args.ValueKind != JsonValueKind.Object) return fallback;
        return args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
            ? v.GetInt64()
            : fallback;
    }

    private static double DoubleOptional(JsonElement args, string name, double fallback)
    {
        if (args.ValueKind != JsonValueKind.Object) return fallback;
        return args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
            ? v.GetDouble()
            : fallback;
    }

    private static string Str(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var v) || v.ValueKind != JsonValueKind.String)
            throw new FormatException($"args missing string \"{name}\"");
        return v.GetString()!;
    }

    private static int Int(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var v) || v.ValueKind != JsonValueKind.Number)
            throw new FormatException($"args missing number \"{name}\"");
        return v.GetInt32();
    }

    private static int IntOptional(JsonElement args, string name, int fallback)
    {
        if (args.ValueKind != JsonValueKind.Object) return fallback;
        return args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : fallback;
    }

    private static string? OptStr(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object) return null;
        return args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;
    }

    private static bool? OptBool(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object) return null;
        return args.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False)
            ? v.GetBoolean() : null;
    }

    private static int? OptInt(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object) return null;
        return args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
            ? v.GetInt32() : null;
    }

    private static string RawArray(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var v)
            || v.ValueKind != JsonValueKind.Array)
            throw new FormatException($"args missing array \"{name}\"");
        return v.GetRawText();
    }

    private static string RawObjectOrEmpty(JsonElement args, string name) =>
        args.ValueKind == JsonValueKind.Object && args.TryGetProperty(name, out var v)
            && v.ValueKind == JsonValueKind.Object
                ? v.GetRawText() : "";

    private static PageNumberingOp ParsePageNumberingOp(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var op))
            return new PageNumberingOp();
        return DocxSessionJson.ParsePageNumberingOp(op);
    }

    private static PageSetupOp ParsePageSetupOp(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var op))
            return new PageSetupOp();
        return DocxSessionJson.ParsePageSetupOp(op);
    }

    private static Position ParsePos(JsonElement args, string name) =>
        DocxSessionJson.ParsePos(Str(args, name));

    private static CharSpan? ParseOptionalSpan(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object) return null;
        if (!args.TryGetProperty(name, out var s) || s.ValueKind != JsonValueKind.Object) return null;
        return new CharSpan(
            s.TryGetProperty("start", out var st) && st.ValueKind == JsonValueKind.Number ? st.GetInt32() : 0,
            s.TryGetProperty("length", out var ln) && ln.ValueKind == JsonValueKind.Number ? ln.GetInt32() : 0);
    }

    private static FormatOp ParseFormatOp(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var op) || op.ValueKind != JsonValueKind.Object)
            return new FormatOp();
        return DocxSessionJson.ParseFormatOp(op.GetRawText());
    }

    private static ParagraphFormatOp ParseParagraphFormatOp(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var op) || op.ValueKind != JsonValueKind.Object)
            return new ParagraphFormatOp();
        return DocxSessionJson.ParseParagraphFormatOp(op.GetRawText());
    }

    private static string[] ParseAnchorIdArray(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty("anchorIds", out var arr) || arr.ValueKind != JsonValueKind.Array)
            throw new FormatException("args missing array \"anchorIds\"");
        var result = new string[arr.GetArrayLength()];
        int i = 0;
        foreach (var el in arr.EnumerateArray())
        {
            if (el.ValueKind != JsonValueKind.String)
                throw new FormatException("\"anchorIds\" entries must be strings");
            result[i++] = el.GetString()!;
        }
        return result;
    }

    private static FindOptions? ParseFindOptions(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty("options", out var o) || o.ValueKind != JsonValueKind.Object)
            return null;
        return DocxSessionJson.ParseFindOptions(o);
    }

    private static CrossReferenceOptions? ParseCrossReferenceOptions(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object
            || !args.TryGetProperty("options", out var options)
            || options.ValueKind != JsonValueKind.Object)
            return null;
        return new CrossReferenceOptions
        {
            ReferenceNumber = options.TryGetProperty("referenceNumber", out var number)
                && number.ValueKind == JsonValueKind.True,
            Hyperlink = options.TryGetProperty("hyperlink", out var link)
                && link.ValueKind == JsonValueKind.True,
            IncludePosition = options.TryGetProperty("includePosition", out var position)
                && position.ValueKind == JsonValueKind.True,
        };
    }

    private static ReplaceOptions? ParseReplaceOptions(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object) return null;
        var hasOptions = args.TryGetProperty("options", out var o) && o.ValueKind == JsonValueKind.Object;
        var preconditions = ParsePreconditions(args);
        if (preconditions is null && hasOptions
            && o.TryGetProperty("preconditions", out var nestedPreconditions))
            preconditions = DocxSessionJson.ParseMutationPreconditions(nestedPreconditions);
        if (!hasOptions && preconditions is null) return null;
        return new ReplaceOptions
        {
            IgnoreCase = hasOptions && DocxSessionJson.TryGetBool(o, "ignoreCase", false),
            MaxReplacements = hasOptions && o.TryGetProperty("maxReplacements", out var mr) && mr.ValueKind == JsonValueKind.Number
                ? mr.GetInt32() : (int?)null,
            ExpectedMatchCount = hasOptions && o.TryGetProperty("expectedMatchCount", out var emc) && emc.ValueKind == JsonValueKind.Number
                ? emc.GetInt32() : (int?)null,
            Preconditions = preconditions,
        };
    }

    private static string ExecuteBatch(JsonElement args, bool preview = false)
    {
        var liveHandle = Handle(args);
        var mode = args.TryGetProperty("mode", out var m) && m.ValueKind == JsonValueKind.String
            ? m.GetString() : "atomic";
        var batchMode = mode switch
        {
            "atomic" => MutationBatchMode.Atomic,
            "best_effort" => MutationBatchMode.BestEffort,
            _ => throw new ArgumentException($"unknown batch mode: {mode}"),
        };
        if (!args.TryGetProperty("steps", out var steps) || steps.ValueKind != JsonValueKind.Array)
            throw new ArgumentException("execute_batch requires an array 'steps'");

        IEnumerable<MutationBatchStep> ParseSteps(int targetHandle)
        {
            var parsed = new List<MutationBatchStep>();
            foreach (var step in steps.EnumerateArray())
            {
                if (step.ValueKind != JsonValueKind.Object)
                    throw new ArgumentException("each batch step must be an object");
                var operation = step.TryGetProperty("operation", out var op) && op.ValueKind == JsonValueKind.String
                    ? op.GetString()! : throw new ArgumentException("batch step missing string 'operation'");
                var stepArgs = step.TryGetProperty("args", out var a) && a.ValueKind == JsonValueKind.Object
                    ? WithHandle(a, targetHandle) : WithHandle(default, targetHandle);
                EditError? preflight = IsBatchableMutation(operation)
                    ? null
                    : new EditError(EditErrorCode.InvalidBatchStep,
                        $"unsupported or non-mutation batch operation: {operation}");
                parsed.Add(DocxSessionOps.SerializedBatchStep(
                    "docx_scalpel",
                    operation,
                    () => Dispatch(operation, stepArgs),
                    preflight is null ? null : () => preflight));
            }
            return parsed;
        }

        if (!preview)
            return DocxSessionOps.ExecuteBatch(liveHandle, batchMode, ParseSteps(liveHandle));

        var htmlMode = args.TryGetProperty("htmlMode", out var html) && html.ValueKind == JsonValueKind.String
            ? html.GetString() switch
            {
                "scoped" => MutationPreviewHtmlMode.Scoped,
                "full" => MutationPreviewHtmlMode.Full,
                null or "none" => MutationPreviewHtmlMode.None,
                var value => throw new ArgumentException($"unknown preview html mode: {value}"),
            }
            : MutationPreviewHtmlMode.None;
        return DocxSessionOps.PreviewBatch(
            liveHandle,
            batchMode,
            ParseSteps,
            new MutationBatchPreviewOptions
            {
                HtmlMode = htmlMode,
                HtmlAnchorId = args.TryGetProperty("htmlAnchorId", out var anchor)
                    && anchor.ValueKind == JsonValueKind.String ? anchor.GetString() : null,
            });
    }

    private static JsonElement WithHandle(JsonElement args, int handle)
    {
        var values = args.ValueKind == JsonValueKind.Object
            ? JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(args.GetRawText())!
            : new Dictionary<string, JsonElement>();
        values["handle"] = JsonSerializer.SerializeToElement(handle);
        return JsonSerializer.SerializeToElement(values);
    }

    private static MutationPreconditions? ParsePreconditions(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object
            || !args.TryGetProperty("preconditions", out var p)
            || p.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        var parsed = DocxSessionJson.ParseMutationPreconditions(p);
        if (parsed is null || parsed.AnchorId is not null) return parsed;
        foreach (var targetName in new[]
        {
            "anchorId", "cellAnchorId", "sourceAnchorId", "fromAnchorId",
            "firstAnchorId", "headingAnchorId", "parentAnchorId", "newAnchorId",
        })
        {
            if (args.TryGetProperty(targetName, out var target) && target.ValueKind == JsonValueKind.String)
                return parsed with { AnchorId = target.GetString() };
        }
        return parsed;
    }

    /// <summary>The request's optional <c>options</c> object, or null for "all defaults".</summary>
    private static T? OptionsOrNull<T>(JsonElement args, Func<JsonElement, T> parse)
        where T : class =>
        args.TryGetProperty("options", out var options) && options.ValueKind == JsonValueKind.Object
            ? parse(options)
            : null;

    private static bool IsMutation(string op) => op is
        "replace_text" or "delete_block" or "move_block" or "delete_range" or "delete_section"
        or "replace_text_range" or "replace_text_at_span" or "replace_inner"
        or "insert_paragraph" or "split_paragraph" or "merge_paragraphs"
        or "set_header_text" or "set_footer_text" or "insert_page_number_field"
        or "ensure_header_footer_visible" or "set_page_numbering" or "clear_page_numbering"
        or "set_header_footer_kind_enabled" or "set_page_setup"
        or "insert_table_of_contents" or "insert_table_of_figures" or "insert_table_of_authorities"
        or "insert_footnote" or "insert_endnote" or "insert_cross_reference"
        or "add_comment" or "add_comment_reply" or "update_comment"
        or "set_comment_resolved" or "remove_comment"
        or "accept_revision" or "reject_revision"
        or "accept_all_revisions" or "reject_all_revisions"
        or "apply_format" or "apply_format_by_substring" or "set_paragraph_style"
        or "set_paragraph_format" or "set_list_level" or "remove_list_membership"
        or "apply_list_format" or "apply_list_format_range" or "set_list_start_override"
        or "clear_list_start_override" or "replace_cell_content"
        or "raw_insert_xml" or "raw_replace_xml"
        or "add_annotation" or "remove_annotation" or "update_annotation" or "move_annotation"
        or "undo" or "redo";

    /// <summary>
    /// Operations a batch step may run (issue #445). Deliberately NOT <see cref="IsMutation"/>:
    /// that predicate also gates the single-call precondition pre-check, and every structural
    /// table op is absent from it, so reusing it rejected <c>insert_table</c> and friends as
    /// <c>invalid_batch_step</c> while the identical MCP step was accepted. Undo/redo are
    /// excluded because a batch step must not move the shared history cursor, and
    /// <c>set_tracked_changes</c>/<c>set_revision_author</c> because they are session
    /// configuration rather than document mutations — both matching the MCP batch allowlist.
    /// </summary>
    private static bool IsBatchableMutation(string op) =>
        (IsMutation(op) && op is not ("undo" or "redo"))
        || op is "insert_table" or "insert_table_row" or "insert_table_column"
            or "delete_table_row" or "delete_table_column"
            or "merge_cells" or "unmerge_cells" or "set_column_widths"
            or "set_table_borders" or "set_cell_shading"
            or "set_repeat_header_row" or "set_table_row_options";

    private static string JsonString(string s) => DocxSessionJson.JsonString(s);

    private static string JsonObject(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var v) || v.ValueKind != JsonValueKind.Object)
            throw new FormatException($"args missing object \"{name}\"");
        return v.GetRawText();
    }

    private static JsonElement JsonObjectElement(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var v)
            || v.ValueKind != JsonValueKind.Object)
            throw new FormatException($"args missing object \"{name}\"");
        return v;
    }

    private static string JsonObjectOrEmpty(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var value))
            return "{}";
        if (value.ValueKind != JsonValueKind.Object)
            throw new FormatException($"optional argument \"{name}\" must be an object when present");
        return value.GetRawText();
    }
}
