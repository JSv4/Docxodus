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

// A date control holds a calendar day. The desktop word processor stores the day it shows at
// midnight UTC and keeps no time of day, so the stored day must match the displayed one
// whatever the input's offset or the host's time zone.
[Collection("Date fill time zone")]
public class ContentControlDateTests
{
    [Theory]
    [InlineData("2032-03-04", null, "2032-03-04")]
    [InlineData(" 2032-03-04 ", null, "2032-03-04")]
    [InlineData("2032-03-04T01:00:00+05:00", null, "2032-03-04")]
    [InlineData("2032-03-04T00:30:00+14:00", null, "2032-03-04")]
    [InlineData("2032-03-04T23:30:00-12:00", null, "2032-03-04")]
    [InlineData("2032-03-04T01:00:00", null, "2032-03-04")]
    [InlineData("2032-03-04T23:00:00Z", null, "2032-03-04")]
    [InlineData("2032-03-04", "", "")]
    public void DateFillStoresTheWrittenDayAtMidnightUtc(string input, string? display, string expectedText)
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
            AssertStored(DocxSessionOps.Save(handle), "2032-03-04T00:00:00Z", expectedText);
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

    [Theory]
    [InlineData(0, 0, 1)]
    [InlineData(0, 0, 14)]
    [InlineData(23, 59, -12)]
    public void SessionDateFillStoresTheValuesOwnDay(int hour, int minute, int offsetHours)
    {
        using var session = new DocxSession(DocxSessionContentControlTests.BuildFixture());
        var control = session.ListContentControls().Single(c => c.Type == ContentControlType.Date);
        var value = new DateTimeOffset(2032, 3, 4, hour, minute, 0, TimeSpan.FromHours(offsetHours));
        Assert.True(session.SetContentControlDate(control.AnchorId, value).Success);
        AssertStored(session.Save(), "2032-03-04T00:00:00Z", "2032-03-04");
    }

    private static void AssertStored(byte[] package, string expectedFullDate, string expectedText)
    {
        using var bytes = new MemoryStream(package);
        using var saved = WordprocessingDocument.Open(bytes, false);
        var date = saved.MainDocumentPart!.GetXDocument().Descendants(W.date).Single();
        Assert.Equal(expectedFullDate, (string?)date.Attribute(W.fullDate));
        Assert.Equal(expectedText, date.Ancestors(W.sdt).First().Element(W.sdtContent)!.Value);
    }
}
