using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;

namespace S95.Blazor.Export.Services;

/// <summary>
/// Default <see cref="IExportService"/> implementation that delegates to the format services.
/// </summary>
public class ExportService : IExportService
{
    private readonly IBaseExportService CsvService;
    private readonly IBaseExportService ExcelService;
    private readonly IBaseExportService PdfService;
    private readonly IBaseExportService WordService;

    /// <summary>Creates the service from the format services.</summary>
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

    /// <inheritdoc />
    public byte[] ExportToCsv<T>(IEnumerable<T> data)
        => CsvService.Export(GetColumns<T>(), data);

    /// <inheritdoc />
    public byte[] ExportToCsv<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => CsvService.Export(columns, data);

    /// <inheritdoc />
    public byte[] ExportToExcel<T>(IEnumerable<T> data)
        => ExcelService.Export(GetColumns<T>(), data);

    /// <inheritdoc />
    public byte[] ExportToExcel<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => ExcelService.Export(columns, data);

    /// <inheritdoc />
    public byte[] ExportToPdf<T>(IEnumerable<T> data)
        => PdfService.Export(GetColumns<T>(), data);

    /// <inheritdoc />
    public byte[] ExportToPdf<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => PdfService.Export(columns, data);

    /// <inheritdoc />
    public byte[] ExportToWord<T>(IEnumerable<T> data)
        => WordService.Export(GetColumns<T>(), data);

    /// <inheritdoc />
    public byte[] ExportToWord<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => WordService.Export(columns, data);

    private IEnumerable<string> GetColumns<T>()
    {
        return typeof(T)
            .GetExportProperties()
            .Select(x => x.GetColumnName());
    }
}