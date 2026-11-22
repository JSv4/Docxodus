using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

[CollectionDefinition("Date fill time zone", DisableParallelization = true)]
public class DateFillTimeZoneCollection { }

[Collection("Date fill time zone")]
public class ContentControlDateTests
{
    [Theory]
    [InlineData("2032-03-04", null, "2032-03-04T00:00:00Z", "2032-03-04")]
    [InlineData("2032-03-04T01:00:00+05:00", null, "2032-03-03T20:00:00Z", "2032-03-04")]
    [InlineData(" 2032-03-04 ", null, "2032-03-04T00:00:00Z", "2032-03-04")]
    [InlineData("2032-03-04T01:00:00", null, null, "2032-03-04")]
    [InlineData("2032-03-04", "", "2032-03-04T00:00:00Z", "")]
    public void DateFillStoresInstantAndKeepsDisplayOverride(string input, string? display,
        string? expectedUtc, string expectedText)
    {
        int handle = DocxSessionOps.OpenSession(DocxSessionContentControlTests.BuildFixture(), null);
        string? previousZone = Environment.GetEnvironmentVariable("TZ");
        try
        {
            if (!OperatingSystem.IsWindows())
            {
                Environment.SetEnvironmentVariable("TZ", "Pacific/Auckland");
                TimeZoneInfo.ClearCachedData();
                Assert.True(TimeZoneInfo.Local.GetUtcOffset(new DateTime(2032, 3, 4)) > TimeSpan.Zero);
            }
            var control = SessionRegistry.Get(handle).ListContentControls().Single(c => c.Type == ContentControlType.Date);
            using var result = JsonDocument.Parse(DocxSessionOps.SetContentControlDate(
                handle, control.AnchorId, input, display, "{}"));
            Assert.True(result.RootElement.GetProperty("success").GetBoolean());
            using var bytes = new MemoryStream(DocxSessionOps.Save(handle));
            using var saved = WordprocessingDocument.Open(bytes, false);
            var date = saved.MainDocumentPart!.GetXDocument().Descendants(W.date).Single();
            expectedUtc ??= new DateTime(2032, 3, 4, 1, 0, 0, DateTimeKind.Local)
                .ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", System.Globalization.CultureInfo.InvariantCulture);
            Assert.Equal(expectedUtc, (string?)date.Attribute(W.fullDate));
            Assert.Equal(expectedText, date.Ancestors(W.sdt).First().Element(W.sdtContent)!.Value);
        }
        finally
        {
            if (!OperatingSystem.IsWindows())
            {
                Environment.SetEnvironmentVariable("TZ", previousZone);
                TimeZoneInfo.ClearCachedData();
            }
            DocxSessionOps.CloseSession(handle);
        }
    }
}
