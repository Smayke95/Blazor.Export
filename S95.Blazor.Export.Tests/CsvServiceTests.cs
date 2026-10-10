using S95.Blazor.Export.Models;
using S95.Blazor.Export.Services;
using S95.Blazor.Export.Tests.Models;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Text;
using Xunit;

namespace S95.Blazor.Export.Tests;

public class CsvServiceTests
{
    private readonly CsvService CsvService = new();

    [Fact]
    public void Export_KeepsNonAsciiCharacters()
    {
        var csv = Export(["Name", "Note"], [new Person { Name = "Šešić Đorđe, Čačak Žepče" }]);

        Assert.Contains("Šešić Đorđe, Čačak Žepče", csv);
    }

    [Fact]
    public void Export_StartsWithUtf8Bom()
    {
        var bytes = CsvService.Export(["Name", "Note"], [new Person { Name = "A" }]);

        Assert.Equal(Encoding.UTF8.GetPreamble(), bytes.Take(3));
    }

    [Fact]
    public void Export_QuotesValuesWithComma()
    {
        var csv = Export(["Name", "Note"], [new Person { Name = "Smith, John", Note = "3,5" }]);

        Assert.Equal("Name,Note\r\n\"Smith, John\",\"3,5\"\r\n", csv);
    }

    [Fact]
    public void Export_EscapesQuotesAndNewLines()
    {
        var csv = Export(["Name", "Note"], [new Person { Name = "He said \"hello\"", Note = "First line\nSecond line" }]);

        Assert.Equal("Name,Note\r\n\"He said \"\"hello\"\"\",\"First line\nSecond line\"\r\n", csv);
    }

    [Fact]
    public void Export_QuotesColumnNames()
    {
        var csv = Export(["Last name, first name", "Note"], [new Person { Name = "A", Note = "B" }]);

        Assert.StartsWith("\"Last name, first name\",Note\r\n", csv);
    }

    [Fact]
    public void Export_WritesNullAsEmptyValue()
    {
        var csv = Export(["Name", "Note"], [new Person { Name = "A", Note = null }]);

        Assert.Equal("Name,Note\r\nA,\r\n", csv);
    }

    [Fact]
    public void Export_WithFewerColumnNames_UsesPropertyNamesForMissingOnes()
    {
        var csv = Export(["Name"], [new Person { Name = "A", Note = "B" }]);

        Assert.Equal("Name,Note\r\nA,B\r\n", csv);
    }

    [Fact]
    public void Export_WithMoreColumnNames_IgnoresExtraNames()
    {
        var csv = Export(["Name", "Note", "Extra"], [new Person { Name = "A", Note = "B" }]);

        Assert.Equal("Name,Note\r\nA,B\r\n", csv);
    }

    [Fact]
    public void Export_FormatsValuesIndependentOfCurrentCulture()
    {
        var originalCulture = CultureInfo.CurrentCulture;

        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("bs-Latn-BA");

            var csv = Export(FormattedRow.Columns, [new FormattedRow()]);

            Assert.Equal("2015-02-17,2015-02-17 13:45:00,1537.5,2020-01-02,08:30:00,Monday,,\r\n", csv.Split("\r\n", 2)[1]);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }
    }

    [Fact]
    public void Export_WithOptions_UsesCultureAndFormats()
    {
        var service = new CsvService(new ExportOptions
        {
            Culture = new CultureInfo("bs-Latn-BA"),
            DateFormat = "dd.MM.yyyy.",
            DateTimeFormat = "dd.MM.yyyy. HH:mm",
            TimeFormat = "HH:mm"
        });

        var csv = Decode(service.Export(FormattedRow.Columns, [new FormattedRow()]));

        Assert.Equal("17.02.2015.,17.02.2015. 13:45,\"1537,5\",02.01.2020.,08:30,Monday,,\r\n", csv.Split("\r\n", 2)[1]);
    }

    [Fact]
    public void Export_WithDisplayFormat_UsesFormatAndNullText()
    {
        var csv = Export(FormattedRow.Columns, [new FormattedRow { Price = 9.5m, Comment = null }]);

        Assert.EndsWith(",9.50 KM,-\r\n", csv);
    }

    private string Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => Decode(CsvService.Export(columns, data));

    private static string Decode(byte[] bytes)
        => Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);

    private class FormattedRow
    {
        public static readonly string[] Columns = ["Date", "DateTime", "Decimal", "DateOnly", "TimeOnly", "Enum", "Price", "Comment"];

        public DateTime Date { get; set; } = new(2015, 2, 17);

        public DateTime DateTime { get; set; } = new(2015, 2, 17, 13, 45, 0);

        public decimal Decimal { get; set; } = 1537.5m;

        public DateOnly DateOnly { get; set; } = new(2020, 1, 2);

        public TimeOnly TimeOnly { get; set; } = new(8, 30);

        public DayOfWeek Enum { get; set; } = DayOfWeek.Monday;

        [DisplayFormat(DataFormatString = "{0:0.00} KM")]
        public decimal? Price { get; set; }

        [DisplayFormat(NullDisplayText = "-")]
        public string? Comment { get; set; } = string.Empty;
    }
}