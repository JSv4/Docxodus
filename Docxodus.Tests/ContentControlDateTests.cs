using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
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
    [InlineData("2032-03-04", null, "March 4, 2032")]
    [InlineData(" 2032-03-04 ", null, "March 4, 2032")]
    [InlineData("2032-03-04T01:00:00+05:00", null, "March 4, 2032")]
    [InlineData("2032-03-04T00:30:00+14:00", null, "March 4, 2032")]
    [InlineData("2032-03-04T23:30:00-12:00", null, "March 4, 2032")]
    [InlineData("2032-03-04T01:00:00", null, "March 4, 2032")]
    [InlineData("2032-03-04T23:00:00Z", null, "March 4, 2032")]
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
        AssertStored(session.Save(), "2032-03-04T00:00:00Z", "March 4, 2032");
    }

    // What the desktop word processor (build 16.0.20326) shows after 15 December 2026 is typed into a
    // control with each format, language and calendar. A control with no language is US English there.
    [Theory]
    [InlineData("dddd, d MMMM yyyy", "en-GB", "gregorian", "Tuesday, 15 December 2026")]
    [InlineData("'Due' d MMM yy", "en-GB", "gregorian", "Due 15 Dec 26")]
    [InlineData("dddd, d. MMMM yyyy", "de-DE", "gregorian", "Dienstag, 15. Dezember 2026")]
    [InlineData("d MMMM yyyy", "pl-PL", "gregorian", "15 grudnia 2026")]
    [InlineData("yyyy年M月d日", "ja-JP", "gregorian", "2026年12月15日")]
    [InlineData("dd/MM/yyyy", "fr-FR", "gregorian", "15/12/2026")]
    [InlineData("M/d/yyyy h:mm am/pm", "en-US", "gregorian", "12/15/2026 12:00 AM")]
    [InlineData("d/M/yy", "en-GB", "gregorian", "15/12/26")]
    [InlineData(@"d MMM yyyy \x", "en-GB", "gregorian", @"15 Dec 2026 \x")]
    [InlineData("d MMMM yyyy", null, "gregorian", "15 December 2026")]
    [InlineData("d MMMM yyyy", "0809", "gregorian", "15 December 2026")]
    [InlineData("d MMMM yyyy", "en-GB", null, "15 December 2026")]
    [InlineData(null, "en-GB", "gregorian", "15/12/2026")]
    // The reference shows a Hijri date (06/07/1448); Docxodus formats only the Gregorian calendar.
    [InlineData("dd/MM/yyyy", "ar-SA", "hijri", "2026-12-15")]
    // For a language whose usual calendar is another one it shows English month names (Thai) or an
    // Arabic calendar variant's; Docxodus keeps the ISO day rather than show a Buddhist or Hijri year.
    [InlineData("d MMMM yyyy", "th-TH", "gregorian", "2026-12-15")]
    [InlineData("dd MMMM yyyy", "ar-SA", "gregorian", "2026-12-15")]
    public void DefaultDisplayUsesTheControlsFormatAndLanguage(string? format, string? language,
        string? calendar, string expectedText)
    {
        using var session = new DocxSession(FixtureWithDate(format, language, calendar));
        var control = session.ListContentControls().Single(c => c.Type == ContentControlType.Date);
        Assert.True(session.SetContentControlDate(control.AnchorId,
            new DateTimeOffset(2026, 12, 15, 0, 0, 0, TimeSpan.Zero)).Success);
        AssertStored(session.Save(), "2026-12-15T00:00:00Z", expectedText);
    }

    // Without the language's month and day names (the browser build has none), an English or
    // all-number format still formats, and another language's names fall back to the ISO day. A
    // language whose usual calendar is another one keeps the ISO day here too, as it does with names,
    // and English given as a hexadecimal language code is still English.
    [Theory]
    [InlineData("dddd, d MMMM yyyy", "en-GB", "Tuesday, 15 December 2026")]
    [InlineData("d MMMM yyyy", null, "15 December 2026")]
    [InlineData("d MMMM yyyy", "0809", "15 December 2026")]
    [InlineData("d MMMM yyyy", "0409", "15 December 2026")]
    [InlineData("dd.MM.yyyy", "de-DE", "15.12.2026")]
    [InlineData("d MMMM yyyy", "de-DE", "2026-12-15")]
    [InlineData("dd/MM/yyyy", "th-TH", "2026-12-15")]
    [InlineData("dd/MM/yyyy", "041E", "2026-12-15")]
    [InlineData("dd/MM/yyyy", "ar-SA", "2026-12-15")]
    [InlineData("dd/MM/yyyy", "fa-IR", "2026-12-15")]
    [InlineData("dd/MM/yyyy", "ar-EG", "15/12/2026")]
    [InlineData(null, "en-GB", "2026-12-15")]
    public void DefaultDisplayWithoutLanguageDataStaysReadable(string? format, string? language, string expectedText)
    {
        var date = DateElement(format, language, "gregorian");
        Assert.Equal(expectedText, ContentControlDateText.Format(new DateTime(2026, 12, 15), date, _ => null));
    }

    private static XElement DateElement(string? format, string? language, string? calendar) =>
        new(W.date,
            format is null ? null : new XElement(W.dateFormat, new XAttribute(W.val, format)),
            language is null ? null : new XElement(W.lid, new XAttribute(W.val, language)),
            calendar is null ? null : new XElement(W.calendar, new XAttribute(W.val, calendar)));

    private static byte[] FixtureWithDate(string? format, string? language, string? calendar)
    {
        using var stream = new MemoryStream();
        stream.Write(DocxSessionContentControlTests.BuildFixture());
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            main.GetXDocument().Descendants(W.date).Single().ReplaceWith(DateElement(format, language, calendar));
            main.PutXDocument();
        }
        return stream.ToArray();
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
