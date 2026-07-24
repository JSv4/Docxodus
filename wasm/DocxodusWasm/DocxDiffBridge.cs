#nullable enable

using System;
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using Docxodus.Internal;

namespace DocxodusWasm;

/// <summary>
/// JSExport bridge for the <c>Docxodus.DocxDiff</c> facade — the IR diff engine's
/// three public entry points (Compare, GetRevisions, GetEditScriptJson). Like
/// <see cref="DocxSessionBridge"/>, this is a thin JSExport-attributed shell over
/// the shared <see cref="DocxDiffOps"/> facade so the WASM and stdio NDJSON
/// transports stay byte-for-byte identical.
///
/// <para><see cref="DocumentComparer"/> also defaults to this engine for primary
/// comparisons. This bridge exposes its additional anchor-addressed revisions and
/// diff-as-data edit script. Settings arrive as a JSON object (the transport mirror
/// of <c>DocxDiffSettings</c>); an empty/whitespace string uses the defaults.</para>
/// </summary>
[SupportedOSPlatform("browser")]
public static partial class DocxDiffBridge
{
    /// <summary>
    /// Compare two DOCX byte arrays and return the redlined DOCX as bytes
    /// (native <c>w:ins</c>/<c>w:del</c>/<c>w:moveFrom</c>/<c>w:moveTo</c>/<c>w:rPrChange</c>
    /// markup). Returns an empty array on error.
    /// </summary>
    /// <param name="leftBytes">The earlier/original document.</param>
    /// <param name="rightBytes">The later/revised document.</param>
    /// <param name="settingsJson">JSON object mirroring <c>DocxDiffSettings</c>; empty string for defaults.</param>
    [JSExport]
    public static byte[] Compare(byte[] leftBytes, byte[] rightBytes, string settingsJson)
    {
        try
        {
            return DocxDiffOps.Compare(leftBytes, rightBytes, settingsJson);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"DocxDiff.Compare error: {ex.GetType().Name}: {ex.Message}");
            return Array.Empty<byte>();
        }
    }

    /// <summary>
    /// Compare two DOCX byte arrays and return the anchor-addressed revision list
    /// as a JSON object (<c>{"revisions":[…]}</c>), or a JSON error object.
    /// </summary>
    [JSExport]
    public static string GetRevisionsJson(byte[] leftBytes, byte[] rightBytes, string settingsJson)
    {
        try
        {
            return DocxDiffOps.GetRevisionsJson(leftBytes, rightBytes, settingsJson);
        }
        catch (Exception ex)
        {
            return DocumentConverter.SerializeError(ex.Message, ex.GetType().Name);
        }
    }

    /// <summary>
    /// Compare two DOCX byte arrays and return the engine's edit script as a JSON
    /// string (the diff-as-data differentiator), or a JSON error object.
    /// </summary>
    [JSExport]
    public static string GetEditScriptJson(byte[] leftBytes, byte[] rightBytes, string settingsJson)
    {
        try
        {
            return DocxDiffOps.GetEditScriptJson(leftBytes, rightBytes, settingsJson);
        }
        catch (Exception ex)
        {
            return DocumentConverter.SerializeError(ex.Message, ex.GetType().Name);
        }
    }

    /// <summary>
    /// Accept every tracked revision in a redlined DOCX and return the resulting
    /// bytes (materializes the "right"/revised side). The byte-in, byte-out
    /// counterpart of <see cref="Compare"/> — <c>AcceptRevisions(Compare(left, right))</c>
    /// ≡ <c>right</c> at the per-block text level — so clients can verify the
    /// round-trip contract, not just the redline's shape. Empty array on error.
    /// </summary>
    [JSExport]
    public static byte[] AcceptRevisions(byte[] bytes)
    {
        try
        {
            return DocxDiffOps.AcceptRevisions(bytes);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"DocxDiff.AcceptRevisions error: {ex.GetType().Name}: {ex.Message}");
            return Array.Empty<byte>();
        }
    }

    /// <summary>
    /// Reject every tracked revision in a redlined DOCX and return the resulting
    /// bytes (materializes the "left"/original side) — <c>RejectRevisions(Compare(left, right))</c>
    /// ≡ <c>left</c> at the per-block text level. Empty array on error.
    /// </summary>
    [JSExport]
    public static byte[] RejectRevisions(byte[] bytes)
    {
        try
        {
            return DocxDiffOps.RejectRevisions(bytes);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"DocxDiff.RejectRevisions error: {ex.GetType().Name}: {ex.Message}");
            return Array.Empty<byte>();
        }
    }

    /// <summary>
    /// Consolidate multiple reviewers' edits against a shared base DOCX and return
    /// the merged redlined DOCX as bytes (native tracked-changes markup, per-author
    /// attribution). Returns an empty array on error.
    /// </summary>
    /// <param name="baseBytes">The shared base/original document.</param>
    /// <param name="reviewersJson">JSON array of reviewers — each <c>{"author": string, "docB64": base64-docx}</c> (the shape <see cref="DocxDiffOps.ParseReviewers"/> reads; an element lacking <c>docB64</c> is skipped).</param>
    /// <param name="settingsJson">JSON object mirroring <c>DocxDiffSettings</c>; empty string for defaults.</param>
    [JSExport]
    public static byte[] Consolidate(byte[] baseBytes, string reviewersJson, string settingsJson)
    {
        try
        {
            return DocxDiffOps.Consolidate(baseBytes, reviewersJson, settingsJson);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"DocxDiff.Consolidate error: {ex.GetType().Name}: {ex.Message}");
            return Array.Empty<byte>();
        }
    }

    /// <summary>
    /// Consolidate multiple reviewers' edits against a shared base DOCX and return
    /// the merged anchor-addressed revision list as a JSON object
    /// (<c>{"revisions":[…]}</c>), or a JSON error object.
    /// </summary>
    [JSExport]
    public static string GetConsolidatedRevisionsJson(byte[] baseBytes, string reviewersJson, string settingsJson)
    {
        try
        {
            return DocxDiffOps.GetConsolidatedRevisionsJson(baseBytes, reviewersJson, settingsJson);
        }
        catch (Exception ex)
        {
            return DocumentConverter.SerializeError(ex.Message, ex.GetType().Name);
        }
    }

    /// <summary>
    /// Consolidate multiple reviewers' edits against a shared base DOCX and return
    /// the merged edit script as a JSON string (the diff-as-data differentiator),
    /// or a JSON error object.
    /// </summary>
    [JSExport]
    public static string GetConsolidatedEditScriptJson(byte[] baseBytes, string reviewersJson, string settingsJson)
    {
        try
        {
            return DocxDiffOps.GetConsolidatedEditScriptJson(baseBytes, reviewersJson, settingsJson);
        }
        catch (Exception ex)
        {
            return DocumentConverter.SerializeError(ex.Message, ex.GetType().Name);
        }
    }

    /// <summary>
    /// Consolidate multiple reviewers' edits against a shared base DOCX and return
    /// the detected conflicts (overlapping reviewer edits) as a JSON string, or a
    /// JSON error object.
    /// </summary>
    [JSExport]
    public static string GetConflictsJson(byte[] baseBytes, string reviewersJson, string settingsJson)
    {
        try
        {
            return DocxDiffOps.GetConflictsJson(baseBytes, reviewersJson, settingsJson);
        }
        catch (Exception ex)
        {
            return DocumentConverter.SerializeError(ex.Message, ex.GetType().Name);
        }
    }
}
