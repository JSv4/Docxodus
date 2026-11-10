using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
#if !WASM_BUILD
using SkiaSharp;
#endif

namespace Docxodus;

public static partial class WmlToHtmlConverter
{
    private static readonly HashSet<string> ImageContentTypes = new(StringComparer.OrdinalIgnoreCase) {
        "image/png", "image/gif", "image/tiff", "image/jpeg" };

    private static XElement? RenderImagePart(ImagePart part, XElement element, Func<ImageInfo, XElement?> handler,
        WmlToHtmlConverterSettings? settings, XAttribute? style = null, string? alt = null)
    {
        string contentType = part.ContentType.ToLowerInvariant();
        var imageType = contentType switch {
            "image/x-wmf" or "image/wmf" => UnsupportedContentType.WmfImage,
            "image/x-emf" => UnsupportedContentType.EmfImage,
            "image/svg+xml" => UnsupportedContentType.SvgImage,
            _ => UnsupportedContentType.Other };
        XElement? Unsupported(string? reason = null)
        {
            if (settings == null) return null;
            string text = imageType switch {
                UnsupportedContentType.WmfImage => "[WMF IMAGE]", UnsupportedContentType.EmfImage => "[EMF IMAGE]",
                UnsupportedContentType.SvgImage => "[SVG IMAGE]", _ => "[UNSUPPORTED IMAGE]" };
            var placeholder = CreateUnsupportedContentPlaceholder(settings, element, imageType, text);
            if (reason != null && settings.IncludeUnsupportedContentMetadata)
                placeholder?.SetAttributeValue("title", contentType + ": " + reason);
            return placeholder;
        }
        if (!ImageContentTypes.Contains(contentType) && imageType is not (UnsupportedContentType.WmfImage or UnsupportedContentType.EmfImage)) return Unsupported();
        byte[] imageBytes;
        EmfImage? emf = null;
        try
        {
            using var input = part.GetStream();
            using var buffer = new MemoryStream();
            var chunk = new byte[8192];
            int count;
            while ((count = input.Read(chunk, 0, chunk.Length)) > 0)
            {
                if (buffer.Length + count > DocxSession.MaxImageInputBytes)
                    throw new InvalidDataException("Image exceeds the 64 MiB input limit.");
                buffer.Write(chunk, 0, count);
            }
            imageBytes = buffer.ToArray();
            if (contentType == "image/tiff") { imageBytes = new TiffImage(imageBytes).ToPng(); contentType = "image/png"; }
            else if (imageType == UnsupportedContentType.WmfImage) { imageBytes = WmfImage.ToSvg(imageBytes); contentType = "image/svg+xml"; }
            else if (imageType == UnsupportedContentType.EmfImage)
            {
                emf = EmfImage.Read(imageBytes);
                imageBytes = System.Text.Encoding.UTF8.GetBytes(emf.Drawing.ToXmlString(SaveOptions.DisableFormatting));
                contentType = "image/svg+xml";
            }
        }
        catch (Exception error) when (error is InvalidDataException or NotSupportedException or OverflowException)
        {
            return Unsupported(error.Message);
        }
#if !WASM_BUILD
        SKBitmap? bitmap = null;
        try { bitmap = SKBitmap.Decode(imageBytes); }
        catch (Exception) { /* A custom handler can still use the encoded resource. */ }
        using (bitmap)
#endif
        {
        var image = handler(new ImageInfo {
#if !WASM_BUILD
            Bitmap = bitmap,
#endif
            ImageBytes = imageBytes, ContentType = contentType, DrawingElement = element, InlineVector = emf,
            ImgStyleAttribute = style, AltText = alt });
        if (emf == null || image == null) return image;
        return image.DescendantsAndSelf().Any(e => ReferenceEquals(e.Annotation<EmfImage>(), emf))
            ? image : Unsupported("EMF paint overflow requires ImageInfo.CreateHtmlImage in the image handler.");
        }
    }
}
