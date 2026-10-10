using S95.Blazor.Export.Services;
using Xunit;

namespace S95.Blazor.Export.Tests;

public class PdfLayoutTests
{
    private static double Measure(string text) => text.Length;

    [Fact]
    public void GetColumnWidths_WhenContentFits_FillsWholeWidthProportionally()
    {
        var widths = PdfService.GetColumnWidths([10, 30], 200);

        Assert.Equal([50, 150], widths);
    }

    [Fact]
    public void GetColumnWidths_WhenContentIsTooWide_KeepsNarrowColumnsAndSplitsRest()
    {
        var widths = PdfService.GetColumnWidths([20, 500, 300], 300);

        Assert.Equal([20, 140, 140], widths);
    }

    [Fact]
    public void GetColumnWidths_WhenAllColumnsAreWide_SplitsWidthEqually()
    {
        var widths = PdfService.GetColumnWidths([400, 500, 600], 300);

        Assert.Equal([100, 100, 100], widths);
    }

    [Fact]
    public void GetColumnWidths_WithEmptyColumns_SplitsWidthEqually()
    {
        var widths = PdfService.GetColumnWidths([0, 0], 100);

        Assert.Equal([50, 50], widths);
    }

    [Fact]
    public void GetColumnWidths_WithNoColumns_ReturnsEmpty()
    {
        Assert.Empty(PdfService.GetColumnWidths([], 100));
    }

    [Fact]
    public void WrapText_WhenTextFits_ReturnsSingleLine()
    {
        Assert.Equal(["short text"], PdfService.WrapText("short text", 20, Measure));
    }

    [Fact]
    public void WrapText_BreaksLinesBetweenWords()
    {
        var lines = PdfService.WrapText("one two three four", 9, Measure);

        Assert.Equal(["one two", "three", "four"], lines);
    }

    [Fact]
    public void WrapText_KeepsExplicitLineBreaks()
    {
        var lines = PdfService.WrapText("first\nsecond\r\nthird", 20, Measure);

        Assert.Equal(["first", "second", "third"], lines);
    }

    [Fact]
    public void WrapText_BreaksWordsLongerThanWidth()
    {
        var lines = PdfService.WrapText("abcdefghij xy", 4, Measure);

        Assert.Equal(["abcd", "efgh", "ij", "xy"], lines);
    }

    [Fact]
    public void WrapText_WithEmptyText_ReturnsSingleEmptyLine()
    {
        Assert.Equal([""], PdfService.WrapText(string.Empty, 10, Measure));
    }
}