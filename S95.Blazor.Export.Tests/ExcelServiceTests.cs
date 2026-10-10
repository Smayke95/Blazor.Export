using ClosedXML.Excel;
using S95.Blazor.Export.Services;
using S95.Blazor.Export.Tests.Models;
using Xunit;

namespace S95.Blazor.Export.Tests;

public class ExcelServiceTests
{
    private readonly ExcelService ExcelService = new();

    [Fact]
    public void Export_CreatesSingleWorksheetWithTable()
    {
        using var workbook = Export(["Name", "Note"], [new Person { Name = "A" }]);

        var worksheet = Assert.Single(workbook.Worksheets);
        Assert.Equal("Sheet1", worksheet.Name);
        Assert.Single(worksheet.Tables);
    }

    [Fact]
    public void Export_WritesColumnNamesInFirstRow()
    {
        using var workbook = Export(["Name", "Note"], [new Person { Name = "A" }]);
        var worksheet = workbook.Worksheet(1);

        Assert.Equal("Name", worksheet.Cell(1, 1).GetString());
        Assert.Equal("Note", worksheet.Cell(1, 2).GetString());
    }

    [Fact]
    public void Export_WritesOneRowPerItem()
    {
        using var workbook = Export(["Name", "Note"],
        [
            new Person { Name = "John", Note = "First" },
            new Person { Name = "Jane", Note = "Second" }
        ]);
        var worksheet = workbook.Worksheet(1);

        Assert.Equal(3, worksheet.LastRowUsed()!.RowNumber());
        Assert.Equal("John", worksheet.Cell(2, 1).GetString());
        Assert.Equal("First", worksheet.Cell(2, 2).GetString());
        Assert.Equal("Jane", worksheet.Cell(3, 1).GetString());
        Assert.Equal("Second", worksheet.Cell(3, 2).GetString());
    }

    [Fact]
    public void Export_KeepsNonAsciiCharacters()
    {
        using var workbook = Export(["Name", "Note"], [new Person { Name = "Šešić Đorđe, Čačak Žepče" }]);

        Assert.Equal("Šešić Đorđe, Čačak Žepče", workbook.Worksheet(1).Cell(2, 1).GetString());
    }

    [Fact]
    public void Export_WritesNullAsEmptyCell()
    {
        using var workbook = Export(["Name", "Note"], [new Person { Name = "A", Note = null }]);

        Assert.True(workbook.Worksheet(1).Cell(2, 2).IsEmpty());
    }

    [Theory]
    [InlineData(1, XLDataType.Number)]
    [InlineData(2, XLDataType.Number)]
    [InlineData(3, XLDataType.Number)]
    [InlineData(4, XLDataType.DateTime)]
    [InlineData(5, XLDataType.DateTime)]
    [InlineData(6, XLDataType.TimeSpan)]
    [InlineData(7, XLDataType.TimeSpan)]
    [InlineData(8, XLDataType.Boolean)]
    [InlineData(9, XLDataType.Text)]
    [InlineData(10, XLDataType.Text)]
    [InlineData(11, XLDataType.Text)]
    public void Export_WritesValuesWithMatchingExcelType(int column, XLDataType expectedType)
    {
        using var workbook = Export(TypedRow.Columns, [new TypedRow()]);

        Assert.Equal(expectedType, workbook.Worksheet(1).Cell(2, column).DataType);
    }

    [Fact]
    public void Export_WritesNumbersAndDatesAsValues()
    {
        using var workbook = Export(TypedRow.Columns, [new TypedRow()]);
        var worksheet = workbook.Worksheet(1);

        Assert.Equal(42, worksheet.Cell(2, 1).GetValue<int>());
        Assert.Equal(1537.5m, worksheet.Cell(2, 2).GetValue<decimal>());
        Assert.Equal(new DateTime(2015, 2, 17, 13, 45, 0), worksheet.Cell(2, 4).GetDateTime());
        Assert.Equal(new DateTime(2020, 1, 2), worksheet.Cell(2, 5).GetDateTime());
        Assert.Equal(new TimeSpan(8, 30, 0), worksheet.Cell(2, 7).GetTimeSpan());
    }

    [Fact]
    public void Export_WritesEnumsAsNames()
    {
        using var workbook = Export(TypedRow.Columns, [new TypedRow()]);

        Assert.Equal("Monday", workbook.Worksheet(1).Cell(2, 9).GetString());
    }

    [Fact]
    public void Export_WritesNullNumberAsEmptyCell()
    {
        using var workbook = Export(TypedRow.Columns, [new TypedRow { NullableNumber = null }]);

        Assert.True(workbook.Worksheet(1).Cell(2, 3).IsEmpty());
    }

    [Fact]
    public void Export_WritesNullableNumberWithValueAsNumber()
    {
        using var workbook = Export(TypedRow.Columns, [new TypedRow { NullableNumber = 2.5 }]);
        var cell = workbook.Worksheet(1).Cell(2, 3);

        Assert.Equal(XLDataType.Number, cell.DataType);
        Assert.Equal(2.5, cell.GetDouble());
    }

    [Fact]
    public void Export_WithFewerColumnNames_UsesPropertyNamesForMissingOnes()
    {
        using var workbook = Export(["Name"], [new Person { Name = "A", Note = "B" }]);
        var worksheet = workbook.Worksheet(1);

        Assert.Equal("Note", worksheet.Cell(1, 2).GetString());
        Assert.Equal("B", worksheet.Cell(2, 2).GetString());
    }

    [Fact]
    public void Export_WithMoreColumnNames_IgnoresExtraNames()
    {
        using var workbook = Export(["Name", "Note", "Extra"], [new Person { Name = "A", Note = "B" }]);

        Assert.True(workbook.Worksheet(1).Cell(1, 3).IsEmpty());
    }

    [Fact]
    public void Export_WithDuplicateColumnNames_MakesThemUnique()
    {
        using var workbook = Export(["Name", "name"], [new Person { Name = "A", Note = "B" }]);
        var worksheet = workbook.Worksheet(1);

        Assert.Equal("Name", worksheet.Cell(1, 1).GetString());
        Assert.Equal("name2", worksheet.Cell(1, 2).GetString());
    }

    [Fact]
    public void Export_WithEmptyColumnName_UsesColumnNumber()
    {
        using var workbook = Export(["", "Note"], [new Person { Name = "A", Note = "B" }]);

        Assert.Equal("Column1", workbook.Worksheet(1).Cell(1, 1).GetString());
    }

    private XLWorkbook Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => new(new MemoryStream(ExcelService.Export(columns, data)));

    private class TypedRow
    {
        public static readonly string[] Columns =
            ["Number", "Decimal", "Nullable", "DateTime", "DateOnly", "TimeSpan", "TimeOnly", "Bool", "Enum", "Guid", "Text"];

        public int Number { get; set; } = 42;

        public decimal Decimal { get; set; } = 1537.5m;

        public double? NullableNumber { get; set; } = 1;

        public DateTime DateTime { get; set; } = new(2015, 2, 17, 13, 45, 0);

        public DateOnly DateOnly { get; set; } = new(2020, 1, 2);

        public TimeSpan TimeSpan { get; set; } = TimeSpan.FromHours(1.5);

        public TimeOnly TimeOnly { get; set; } = new(8, 30);

        public bool Bool { get; set; } = true;

        public DayOfWeek Enum { get; set; } = DayOfWeek.Monday;

        public Guid Guid { get; set; } = Guid.Empty;

        public string Text { get; set; } = "text";
    }
}