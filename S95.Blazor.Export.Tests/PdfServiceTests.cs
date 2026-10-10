using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;
using PdfSharp.Pdf.IO;
using S95.Blazor.Export.Services;
using S95.Blazor.Export.Tests.Models;
using System.Text;
using Xunit;

namespace S95.Blazor.Export.Tests;

public class PdfServiceTests
{
    private const double Margin = 40;

    private readonly PdfService PdfService = new();

    [Fact]
    public void Export_ReturnsPdfDocument()
    {
        var bytes = PdfService.Export(["Name", "Note"], [new Person { Name = "A" }]);

        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
    }

    [Fact]
    public void Export_UsesLandscapeOrientation()
    {
        using var document = Export(["Name", "Note"], [new Person { Name = "A" }]);
        var page = document.Pages[0];

        Assert.True(page.Width.Point > page.Height.Point);
    }

    [Fact]
    public void Export_WithFewRows_CreatesSinglePage()
    {
        using var document = Export(["Name", "Note"], [new Person { Name = "A" }, new Person { Name = "B" }]);

        Assert.Equal(1, document.PageCount);
    }

    [Fact]
    public void Export_WithNoData_CreatesSinglePage()
    {
        using var document = Export(["Name", "Note"], Array.Empty<Person>());

        Assert.Equal(1, document.PageCount);
    }

    [Fact]
    public void Export_WithManyRows_CreatesMultiplePagesInLandscape()
    {
        var people = Enumerable.Range(1, 100).Select(i => new Person { Name = $"Person {i}" });

        using var document = Export(["Name", "Note"], people);

        Assert.True(document.PageCount > 1);
        Assert.All<PdfPage>(document.Pages, page => Assert.True(page.Width.Point > page.Height.Point));
    }

    [Fact]
    public void Export_WithNonAsciiCharacters_CreatesDocument()
    {
        using var document = Export(["Name", "Note"], [new Person { Name = "Šešić Đorđe, Čačak Žepče" }]);

        Assert.Equal(1, document.PageCount);
    }

    [Fact]
    public void Export_WorksWithoutSystemFonts()
    {
        using var document = Export(["Name", "Note"], [new Person { Name = "A" }]);

        Assert.False(GlobalFontSettings.UseWindowsFontsUnderWindows);
        Assert.Equal(1, document.PageCount);
    }

    [Fact]
    public void Export_EmbedsLiberationSansFont()
    {
        using var document = Export(["Name", "Note"], [new Person { Name = "A" }]);

        var fonts = GetFontNames(document.Pages[0]);

        Assert.NotEmpty(fonts);
        Assert.All(fonts, font => Assert.Contains("Liberation Sans", font));
    }

    [Fact]
    public void Export_WithManyColumns_KeepsAllCellsInsidePage()
    {
        using var document = Export(Enumerable.Range(1, 10).Select(i => $"Column {i}"), [new WideRow(), new WideRow()]);
        var page = document.Pages[0];

        var cells = GetCells(page);

        Assert.Equal(30, cells.Count);
        Assert.All(cells, cell =>
        {
            Assert.True(cell.X >= Margin - 0.01);
            Assert.True(cell.X + cell.Width <= page.Width.Point - Margin + 0.01);
        });
    }

    [Fact]
    public void Export_WithManyRows_RepeatsHeaderOnEveryPage()
    {
        var people = Enumerable.Range(1, 100).Select(i => new Person { Name = $"Person {i}" });

        using var document = Export(["Name", "Note"], people);

        Assert.All<PdfPage>(document.Pages, page =>
        {
            var firstCell = GetCells(page)[0];

            Assert.Contains(GetFontNames(page), font => font.Contains("Bold"));
            Assert.Equal(page.Height.Point - Margin, firstCell.Y + firstCell.Height, 0.01);
        });
    }

    [Fact]
    public void Export_WithLongText_MakesRowTaller()
    {
        using var document = Export(["Name", "Note"],
        [
            new Person { Name = "Short" },
            new Person { Name = string.Join(" ", Enumerable.Repeat("long text", 200)) }
        ]);

        var heights = GetCells(document.Pages[0]).Select(cell => cell.Height).ToList();

        Assert.True(heights.Max() > 5 * heights.Min());
    }

    [Fact]
    public void Export_WithMoreColumnNames_DrawsOnlyPropertyColumns()
    {
        using var document = Export(["Name", "Note", "Extra"], [new Person { Name = "A", Note = "B" }]);

        Assert.Equal(4, GetCells(document.Pages[0]).Count);
    }

    [Fact]
    public void Export_WithFewerColumnNames_DrawsAllPropertyColumns()
    {
        using var document = Export(["Name"], [new Person { Name = "A", Note = "B" }]);

        Assert.Equal(4, GetCells(document.Pages[0]).Count);
    }

    private PdfDocument Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => PdfReader.Open(new MemoryStream(PdfService.Export(columns, data)), PdfDocumentOpenMode.Import);

    private static List<string> GetFontNames(PdfPage page)
    {
        var fonts = page.Resources.Elements.GetDictionary("/Font")!;

        return fonts.Elements.Values
            .Select(value => (value as PdfReference)?.Value ?? value)
            .OfType<PdfDictionary>()
            .Select(font => font.Elements.GetName("/BaseFont"))
            .ToList();
    }

    private static List<(double X, double Y, double Width, double Height)> GetCells(PdfPage page)
        => ContentReader.ReadContent(page)
            .OfType<COperator>()
            .Where(op => op.OpCode.Name == "re")
            .Select(op => op.Operands.OfType<CNumber>().Select(ToDouble).ToArray())
            .Select(values => (values[0], values[1], values[2], values[3]))
            .ToList();

    private static double ToDouble(CNumber number)
        => number is CReal real ? real.Value : ((CInteger)number).Value;

    private class WideRow
    {
        public string Column1 { get; set; } = "First column with some text";

        public int Column2 { get; set; } = 2;

        public int Column3 { get; set; } = 3;

        public int Column4 { get; set; } = 4;

        public int Column5 { get; set; } = 5;

        public int Column6 { get; set; } = 6;

        public int Column7 { get; set; } = 7;

        public int Column8 { get; set; } = 8;

        public int Column9 { get; set; } = 9;

        public string Column10 { get; set; } = "Last column";
    }
}