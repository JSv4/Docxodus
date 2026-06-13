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
/// <para><c>WmlComparer</c> remains the default comparison surface
/// (<see cref="DocumentComparer"/>); this bridge exposes the NEW engine, whose
/// differentiators are anchor-addressed revisions and the diff-as-data edit
/// script. Settings arrive as a JSON object (the transport mirror of
/// <c>DocxDiffSettings</c>); an empty/whitespace string uses the defaults.</para>
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
}
