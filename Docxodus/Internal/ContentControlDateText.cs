using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// The text a date control shows for a day when no display text is given. The desktop word
/// processor (build 16.0.20326) shows a chosen date in the control's own format (<c>w:dateFormat</c>)
/// and language (<c>w:lid</c>, US English when absent), or in that language's short date when the
/// control has no format. Only the Gregorian calendar is formatted. A control with another calendar, a
/// language whose usual calendar is another one, or a format that needs month or day names this
/// runtime does not have for its language (the browser build has only English), shows the ISO day,
/// yyyy-MM-dd.
/// </summary>
internal static class ContentControlDateText
{
    public static string Format(DateTime day, XElement? date) => Format(day, date, Resolve);

    internal static string Format(DateTime day, XElement? date, Func<string, CultureInfo?> resolve)
    {
        string iso = day.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        string? calendar = (string?)date?.Element(W.calendar)?.Attribute(W.val);
        if (calendar is not null && calendar != "gregorian")
            return iso;
        string language = (string?)date?.Element(W.lid)?.Attribute(W.val) ?? "en-US";
        string? picture = (string?)date?.Element(W.dateFormat)?.Attribute(W.val);
        var culture = resolve(language);
        // A language whose usual calendar is another one (Thai, Saudi Arabic) is shown there through
        // calendar variants this formatter does not model, such as English month names or a Buddhist
        // year in the short date.
        if (culture is not null ? culture.Calendar is not GregorianCalendar : UsesAnotherCalendar(language))
            return iso;
        if (string.IsNullOrEmpty(picture))
            return culture is null ? iso : day.ToString(culture.DateTimeFormat.ShortDatePattern, culture);
        var (format, usesNames) = Translate(picture);
        if (culture is null)
        {
            // English names are the invariant culture's; numbers need no names at all.
            if (usesNames && !IsEnglish(language))
                return iso;
            culture = CultureInfo.InvariantCulture;
        }
        return day.ToString(format, culture);
    }

    // The culture for a w:lid value (a language tag such as en-GB, or a four-digit hex LCID), or null
    // when the runtime has no data for it, as in globalization-invariant mode.
    private static CultureInfo? Resolve(string language)
    {
        if (AppContext.TryGetSwitch("System.Globalization.Invariant", out bool invariant) && invariant)
            return null;
        try
        {
            return TryLcid(language, out int lcid) ? CultureInfo.GetCultureInfo(lcid) : CultureInfo.GetCultureInfo(language);
        }
        catch (ArgumentException)
        {
            return null;
        }
    }

    private static bool TryLcid(string language, out int lcid) =>
        int.TryParse(language, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out lcid) && language.Length == 4;

    // English by language tag (en, en-GB) or by LCID, whose low ten bits are the language (0x09 is English).
    private static bool IsEnglish(string language) =>
        TryLcid(language, out int lcid)
            ? (lcid & 0x3FF) == 0x09
            : language.Equals("en", StringComparison.OrdinalIgnoreCase)
              || language.StartsWith("en-", StringComparison.OrdinalIgnoreCase);

    // Without culture data, the languages whose usual calendar in .NET's culture data is not the
    // Gregorian one: Umm al-Qura for Saudi Arabic, Persian for Persian, Pashto and their relatives,
    // and Thai Buddhist for Thai. Other Arabic regions use the Gregorian calendar.
    private static bool UsesAnotherCalendar(string language) =>
        TryLcid(language, out int lcid) ? OtherCalendarLcids.Contains(lcid) : OtherCalendarLanguages.Contains(language);

    private static readonly HashSet<string> OtherCalendarLanguages =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "ar", "ar-SA", "fa", "fa-AF", "fa-IR", "ku-Arab-IR", "lrc", "lrc-IR", "mzn", "mzn-IR",
            "ps", "ps-AF", "th", "th-TH", "uz-Arab", "uz-Arab-AF",
        };

    private static readonly HashSet<int> OtherCalendarLcids =
        new() { 0x0001, 0x0401, 0x0029, 0x0429, 0x048C, 0x0063, 0x0463, 0x001E, 0x041E };

    // A date picture as a .NET custom format. Day, month, year, hour, minute and second letters and
    // AM/PM (either case, shown as the culture's designator) become specifiers; text in single quotes
    // and every other character, a backslash included, are copied as they are.
    private static (string Format, bool UsesNames) Translate(string picture)
    {
        var format = new StringBuilder();
        bool usesNames = false;
        for (int i = 0; i < picture.Length;)
        {
            char c = picture[i];
            if (c == '\'')
            {
                int end = picture.IndexOf('\'', i + 1);
                if (end < 0)
                    end = picture.Length;
                Literal(format, picture.AsSpan(i + 1, end - i - 1));
                i = Math.Min(end + 1, picture.Length);
            }
            else if (string.Compare(picture, i, "AM/PM", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
            {
                format.Append("tt");
                usesNames = true;
                i += 5;
            }
            else if (c is 'd' or 'M' or 'y' or 'h' or 'H' or 'm' or 's')
            {
                int run = 1;
                while (i + run < picture.Length && picture[i + run] == c)
                    run++;
                int length = c == 'y' ? (run <= 2 ? 2 : 4) : Math.Min(run, c is 'd' or 'M' ? 4 : 2);
                format.Append(c, length);
                usesNames |= c is 'd' or 'M' && length >= 3;
                i += run;
            }
            else
            {
                Literal(format, picture.AsSpan(i, 1));
                i++;
            }
        }
        // A one-letter format would be read as a standard format ("d" is the short date).
        return (format.Length == 1 ? "%" + format : format.ToString(), usesNames);
    }

    private static void Literal(StringBuilder format, ReadOnlySpan<char> text)
    {
        foreach (char c in text)
            format.Append('\\').Append(c);
    }
}
