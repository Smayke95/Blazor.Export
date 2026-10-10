using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.DependencyInjection;
using PdfSharp.Fonts;
using S95.Blazor.Export.Extensions;
using S95.Blazor.Export.Interfaces;
using S95.Blazor.Export.Services;
using S95.Blazor.Export.Tests.Models;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Text;
using Xunit;

namespace S95.Blazor.Export.Tests;

public class ExportServiceTests
{
    private readonly IExportService ExportService;

    private readonly Employee[] Employees =
    [
        new Employee { FullName = "John Smith", Age = 30, City = "Sarajevo", Password = "secret", Skills = ["C#"] }
    ];

    public ExportServiceTests()
    {
        var services = new ServiceCollection();
        services.AddExportServices();

        ExportService = services.BuildServiceProvider().GetRequiredService<IExportService>();
    }

    [Fact]
    public void AddExportServices_RegistersExportService()
    {
        Assert.IsType<ExportService>(ExportService);
    }

    [Fact]
    public void AddExportServices_DoesNotChangeGlobalFontSettings()
    {
        Assert.False(GlobalFontSettings.UseWindowsFontsUnderWindows);
        Assert.Null(GlobalFontSettings.FontResolver);
    }

    [Fact]
    public void AddExportServices_WithOptions_AppliesThemToExports()
    {
        var services = new ServiceCollection();
        services.AddExportServices(options =>
        {
            options.Culture = new CultureInfo("bs-Latn-BA");
            options.DateFormat = "dd.MM.yyyy.";
        });
        var exportService = services.BuildServiceProvider().GetRequiredService<IExportService>();

        var lines = ExportCsvLines(exportService.ExportToCsv([new DateRow()]));

        Assert.Equal("17.02.2015.,\"2,5\"", lines[1]);
    }

    [Fact]
    public void ExportToCsv_UsesDisplayNamesAsColumns()
    {
        var lines = ExportCsvLines(ExportService.ExportToCsv(Employees));

        Assert.Equal("Full name,Years,City", lines[0]);
    }

    [Fact]
    public void ExportToCsv_WithDisplayNameAttribute_UsesItAsColumn()
    {
        var lines = ExportCsvLines(ExportService.ExportToCsv([new DisplayNameRow()]));

        Assert.Equal("Product name", lines[0]);
    }

    [Fact]
    public void ExportToCsv_WithDisplayOrderBeforeName_UsesName()
    {
        var lines = ExportCsvLines(ExportService.ExportToCsv([new DisplayOrderRow()]));

        Assert.Equal("Product name", lines[0]);
    }

    [Fact]
    public void ExportToCsv_WithDisplayWithoutName_UsesPropertyName()
    {
        var lines = ExportCsvLines(ExportService.ExportToCsv([new DisplayDescriptionRow()]));

        Assert.Equal("Product", lines[0]);
    }

    [Fact]
    public void ExportToCsv_WithLocalizedDisplay_UsesResourceValue()
    {
        var lines = ExportCsvLines(ExportService.ExportToCsv([new LocalizedRow()]));

        Assert.Equal("Localized product name", lines[0]);
    }

    [Fact]
    public void ExportToCsv_SkipsIgnoredAndCollectionProperties()
    {
        var lines = ExportCsvLines(ExportService.ExportToCsv(Employees));

        Assert.Equal("John Smith,30,Sarajevo", lines[1]);
    }

    [Fact]
    public void ExportToCsv_WithColumns_UsesGivenColumns()
    {
        var lines = ExportCsvLines(ExportService.ExportToCsv(["Name", "Age", "City"], Employees));

        Assert.Equal("Name,Age,City", lines[0]);
    }

    [Fact]
    public void ExportToExcel_ReturnsWorkbookWithDisplayNames()
    {
        using var workbook = new XLWorkbook(new MemoryStream(ExportService.ExportToExcel(Employees)));
        var worksheet = workbook.Worksheet(1);

        Assert.Equal("Full name", worksheet.Cell(1, 1).GetString());
        Assert.Equal("John Smith", worksheet.Cell(2, 1).GetString());
    }

    [Fact]
    public void ExportToPdf_ReturnsPdfDocument()
    {
        var bytes = ExportService.ExportToPdf(Employees);

        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
    }

    [Fact]
    public void ExportToWord_ReturnsDocumentWithDisplayNames()
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(ExportService.ExportToWord(Employees)), false);

        Assert.Contains("Full name", document.MainDocumentPart!.Document!.Body!.InnerText);
        Assert.DoesNotContain("secret", document.MainDocumentPart!.Document!.Body!.InnerText);
    }

    private static string[] ExportCsvLines(byte[] bytes)
        => Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3).Split("\r\n");

    public static class Resources
    {
        public static string ProductName => "Localized product name";
    }

    private class DateRow
    {
        public DateTime Date { get; set; } = new(2015, 2, 17);

        public double Amount { get; set; } = 2.5;
    }

    private class DisplayNameRow
    {
        [DisplayName("Product name")]
        public string Product { get; set; } = "A";
    }

    private class DisplayOrderRow
    {
        [Display(Order = 1, Name = "Product name")]
        public string Product { get; set; } = "A";
    }

    private class DisplayDescriptionRow
    {
        [Display(Description = "Name of the product")]
        public string Product { get; set; } = "A";
    }

    private class LocalizedRow
    {
        [Display(Name = nameof(Resources.ProductName), ResourceType = typeof(Resources))]
        public string Product { get; set; } = "A";
    }
}