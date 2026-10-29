using System.Xml;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>Character checks shared by text payloads and XML attribute values.</summary>
internal static class XmlText
{
    internal static string SymbolText(XElement symbol) =>
        ushort.TryParse((string?)symbol.Attribute(W.w + "char"), System.Globalization.NumberStyles.HexNumber,
            System.Globalization.CultureInfo.InvariantCulture, out var code)
        && code >= 0x20 && XmlConvert.IsXmlChar((char)code) ? ((char)code).ToString() : string.Empty;

    /// <summary>Trim markdown padding while retaining tabs, which are document content.</summary>
    internal static string TrimWhitespaceExceptTabs(string text, bool trimStart = true)
    {
        int start = 0, end = text.Length;
        while (trimStart && start < end && text[start] != '\t' && char.IsWhiteSpace(text[start])) start++;
        while (end > start && text[end - 1] != '\t' && char.IsWhiteSpace(text[end - 1])) end--;
        return text[start..end];
    }

    internal static bool IsCharacterBoundary(string value, int offset) =>
        offset == 0 || offset == value.Length || !char.IsSurrogatePair(value, offset - 1);

    internal static bool IsValid(string? value)
    {
        if (value is null) return true;
        try { XmlConvert.VerifyXmlChars(value); return true; }
        catch (XmlException) { return false; }
    }

    internal static EditError? ValidatePayload(string? value, string? anchorId = null) =>
        value is null
            ? new EditError(EditErrorCode.MalformedMarkdown, "null payload", anchorId)
            : !IsValid(value)
                ? new EditError(EditErrorCode.MalformedMarkdown,
                    "text contains a character XML cannot represent, such as an unpaired surrogate or a forbidden control character", anchorId)
                : null;
}
