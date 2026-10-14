using System.Xml;

namespace Docxodus.Internal;

/// <summary>Character checks shared by text payloads and XML attribute values.</summary>
internal static class XmlText
{
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
