using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using S95.Blazor.Export.Models;
using S95.Blazor.Export.Services;
using S95.Blazor.Export.Tests.Models;
using Xunit;

namespace S95.Blazor.Export.Tests;

public class WordServiceTests
{
    private readonly WordService WordService = new();

    [Fact]
    public void Export_CreatesValidDocument()
    {
        using var document = Export(["Name", "Note"], [new Person { Name = "A", Note = "B" }]);

        Assert.Empty(new OpenXmlValidator().Validate(document, TestContext.Current.CancellationToken));
    }

    [Fact]
    public void Export_UsesLandscapeOrientation()
    {
        using var document = Export(["Name", "Note"], [new Person { Name = "A" }]);

        var pageSize = document.MainDocumentPart!.Document!.Body!.GetFirstChild<SectionProperties>()!.GetFirstChild<PageSize>()!;

        Assert.Equal(PageOrientationValues.Landscape, pageSize.Orient!.Value);
        Assert.True(pageSize.Width!.Value > pageSize.Height!.Value);
    }

    [Fact]
    public void Export_WritesColumnNamesInFirstRow()
    {
        using var document = Export(["Name", "Note"], [new Person { Name = "A" }]);

        Assert.Equal(["Name", "Note"], GetRows(document)[0]);
    }

    [Fact]
    public void Export_WritesOneRowPerItem()
    {
        using var document = Export(["Name", "Note"],
        [
            new Person { Name = "John", Note = "First" },
            new Person { Name = "Jane", Note = "Second" }
        ]);
        var rows = GetRows(document);

        Assert.Equal(3, rows.Count);
        Assert.Equal(["John", "First"], rows[1]);
        Assert.Equal(["Jane", "Second"], rows[2]);
    }

    [Fact]
    public void Export_WithNoData_WritesOnlyColumnNames()
    {
        using var document = Export(["Name", "Note"], Array.Empty<Person>());

        Assert.Single(GetRows(document));
    }

    [Fact]
    public void Export_KeepsNonAsciiCharactersAndWhitespace()
    {
        using var document = Export(["Name", "Note"], [new Person { Name = "  Šešić Đorđe, Čačak Žepče  " }]);

        Assert.Equal("  Šešić Đorđe, Čačak Žepče  ", GetRows(document)[1][0]);
    }

    [Fact]
    public void Export_WritesNullAsEmptyCell()
    {
        using var document = Export(["Name", "Note"], [new Person { Name = "A", Note = null }]);

        Assert.Equal(["A", ""], GetRows(document)[1]);
    }

    [Fact]
    public void Export_WithFewerColumnNames_UsesPropertyNamesForMissingOnes()
    {
        using var document = Export(["Name"], [new Person { Name = "A", Note = "B" }]);

        Assert.Equal(["Name", "Note"], GetRows(document)[0]);
    }

    [Fact]
    public void Export_WithMoreColumnNames_IgnoresExtraNames()
    {
        using var document = Export(["Name", "Note", "Extra"], [new Person { Name = "A", Note = "B" }]);

        Assert.Equal(["Name", "Note"], GetRows(document)[0]);
    }

    [Fact]
    public void Export_WithOptions_FormatsDates()
    {
        var service = new WordService(new ExportOptions { DateFormat = "dd.MM.yyyy." });

        using var document = WordprocessingDocument.Open(new MemoryStream(service.Export(["Date"], [new DateRow()])), false);

        Assert.Equal(["17.02.2015."], GetRows(document)[1]);
    }

    private WordprocessingDocument Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => WordprocessingDocument.Open(new MemoryStream(WordService.Export(columns, data)), false);

    private static List<string[]> GetRows(WordprocessingDocument document)
        => document
            .MainDocumentPart!
            .Document!
            .Body!
            .Elements<Table>()
            .Single()
            .Elements<TableRow>()
            .Select(row => row.Elements<TableCell>().Select(cell => cell.InnerText).ToArray())
            .ToList();

    private class DateRow
    {
        public DateTime Date { get; set; } = new(2015, 2, 17);
    }
}