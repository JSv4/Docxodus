using System.IO;
using System.Xml;
using Docxodus.Verification;
using DocumentFormat.OpenXml;

namespace Docxodus.Internal;

/// <summary>
/// Measures an SDK tree through WriteTo without producing output. Materialized nodes stop at the
/// first start element past the limit; lazy nodes stream their raw XML through the bounded reader
/// instead of forcing the SDK's recursive parse. Comments and text add no element depth.
/// </summary>
internal sealed class DepthMeasuringXmlWriter(int limit) : XmlWriter
{
    private int _depth;
    private bool _attribute;
    private bool _probeRoot;

    // A materialized SDK root hoists namespace declarations by walking every descendant before
    // writing its children. Stop at its first start tag, then measure the children directly so that
    // hoisting cannot force an unchecked lazy subtree to parse. A lazy root writes raw XML instead.
    internal void Measure(OpenXmlPartRootElement root)
    {
        _probeRoot = true;
        try { root.WriteTo(this); }
        catch (MaterializedRootException)
        {
            _probeRoot = false;
            _depth = 1;
            foreach (var child in root.ChildElements) child.WriteTo(this);
        }
        finally { _probeRoot = false; }
    }

    private sealed class MaterializedRootException : System.Exception;

    private static readonly XmlReaderSettings RawSettings = new()
    {
        ConformanceLevel = ConformanceLevel.Auto,
        DtdProcessing = DtdProcessing.Parse,
        MaxCharactersFromEntities = 10_000_000,
        XmlResolver = null,
        IgnoreWhitespace = true,
    };

    public override void WriteStartElement(string? prefix, string localName, string? ns)
    {
        if (_probeRoot) throw new MaterializedRootException();
        if (++_depth > limit) XmlSemanticNormalizer.ThrowDepthLimit(limit);
    }

    public override void WriteEndElement() => _depth--;
    public override void WriteFullEndElement() => _depth--;

    public override void WriteRaw(string data)
    {
        if (_attribute || string.IsNullOrEmpty(data)) return;
        using var reader = new DepthLimitedXmlReader(
            XmlReader.Create(new StringReader(data), RawSettings), limit - _depth);
        while (reader.Read()) { }
    }

    public override void WriteRaw(char[] buffer, int index, int count) => WriteRaw(new string(buffer, index, count));
    public override void WriteStartAttribute(string? prefix, string localName, string? ns) => _attribute = true;
    public override void WriteEndAttribute() => _attribute = false;
    public override WriteState WriteState => _attribute ? WriteState.Attribute : WriteState.Content;
    public override string? LookupPrefix(string ns) => null;
    public override void Flush() { }
    public override void WriteStartDocument() { }
    public override void WriteStartDocument(bool standalone) { }
    public override void WriteEndDocument() { }
    public override void WriteDocType(string name, string? pubid, string? sysid, string? subset) { }
    public override void WriteCData(string? text) { }
    public override void WriteComment(string? text) { }
    public override void WriteProcessingInstruction(string name, string? text) { }
    public override void WriteEntityRef(string name) { }
    public override void WriteCharEntity(char ch) { }
    public override void WriteWhitespace(string? ws) { }
    public override void WriteString(string? text) { }
    public override void WriteSurrogateCharEntity(char lowChar, char highChar) { }
    public override void WriteChars(char[] buffer, int index, int count) { }
    public override void WriteBase64(byte[] buffer, int index, int count) { }
}
