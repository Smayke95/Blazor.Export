using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;

namespace S95.Blazor.Export.Services;

public class ExportService : IExportService
{
    private readonly IBaseExportService CsvService;
    private readonly IBaseExportService ExcelService;
    private readonly IBaseExportService PdfService;
    private readonly IBaseExportService WordService;

    public ExportService(
        CsvService csvService,
        ExcelService excelService,
        PdfService pdfService,
        WordService wordService)
    {
        CsvService = csvService;
        ExcelService = excelService;
        PdfService = pdfService;
        WordService = wordService;
    }

    public byte[] ExportToCsv<T>(IEnumerable<T> data)
    {
        var columns = typeof(T)
            .GetExportProperties()
            .Select(x => x.CustomAttributes.FirstOrDefault(y => y.AttributeType.Name == "DisplayAttribute")?.NamedArguments.FirstOrDefault().TypedValue.ToString() ?? x.Name);

        return CsvService.Export(columns, data);
    }

    public byte[] ExportToCsv<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => CsvService.Export(columns, data);

    public byte[] ExportToExcel<T>(IEnumerable<T> data)
    {
        var columns = typeof(T)
            .GetExportProperties()
            .Select(x => x.CustomAttributes.FirstOrDefault(y => y.AttributeType.Name == "DisplayAttribute")?.NamedArguments.FirstOrDefault().TypedValue.ToString() ?? x.Name);

        return ExcelService.Export(columns, data);
    }

    public byte[] ExportToExcel<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => ExcelService.Export(columns, data);

    public byte[] ExportToPdf<T>(IEnumerable<T> data)
    {
        var columns = typeof(T)
            .GetExportProperties()
            .Select(x => x.CustomAttributes.FirstOrDefault(y => y.AttributeType.Name == "DisplayAttribute")?.NamedArguments.FirstOrDefault().TypedValue.ToString() ?? x.Name);

        return PdfService.Export(columns, data);
    }

    public byte[] ExportToPdf<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => PdfService.Export(columns, data);

    public byte[] ExportToWord<T>(IEnumerable<T> data)
    {
        var columns = typeof(T)
            .GetExportProperties()
            .Select(x => x.CustomAttributes.FirstOrDefault(y => y.AttributeType.Name == "DisplayAttribute")?.NamedArguments.FirstOrDefault().TypedValue.ToString() ?? x.Name);

        return WordService.Export(columns, data);
    }

    public byte[] ExportToWord<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => WordService.Export(columns, data);
}