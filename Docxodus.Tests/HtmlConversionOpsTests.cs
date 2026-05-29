#nullable enable
using System.IO;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlConversionOpsTests
{
    private static byte[] TourPlanBytes() =>
        File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC001-5DayTourPlanTemplate.docx"));

    [Fact]
    public void HCO001_ConvertBytes_ProducesHtmlWithPrefix()
    {
        var options = new HtmlConversionOptions { CssClassPrefix = "zz-" };

        string html = HtmlConversionOps.ConvertToHtml(TourPlanBytes(), options);

        Assert.Contains("<html", html);
        Assert.Contains("zz-", html);
    }
}
