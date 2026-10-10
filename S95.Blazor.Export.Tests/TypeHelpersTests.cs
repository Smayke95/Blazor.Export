using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;
using S95.Blazor.Export.Services;
using Xunit;

namespace S95.Blazor.Export.Tests;

public class TypeHelpersTests
{
    [Fact]
    public void GetExportProperties_SkipsIndexers()
    {
        var properties = typeof(IndexerRow).GetExportProperties().Select(p => p.Name);

        Assert.Equal(["Value"], properties);
    }

    [Fact]
    public void GetExportProperties_SkipsPropertiesWithoutPublicGetter()
    {
        var properties = typeof(WriteOnlyRow).GetExportProperties().Select(p => p.Name);

        Assert.Equal(["Value"], properties);
    }

    [Theory]
    [InlineData("csv")]
    [InlineData("excel")]
    [InlineData("pdf")]
    [InlineData("word")]
    public void Export_WithIndexerAndWriteOnlyProperties_DoesNotThrow(string format)
    {
        IBaseExportService service = format switch
        {
            "csv" => new CsvService(),
            "excel" => new ExcelService(),
            "pdf" => new PdfService(),
            _ => new WordService()
        };

        Assert.NotEmpty(service.Export(["Value"], [new IndexerRow()]));
        Assert.NotEmpty(service.Export(["Value"], [new WriteOnlyRow()]));
    }

    private class IndexerRow
    {
        public int Value { get; set; } = 1;

        public int this[int index] => index;
    }

    private class WriteOnlyRow
    {
        public int Value { get; set; } = 1;

        public string Secret { set { } }

        public string Hidden { private get; set; } = "hidden";
    }
}