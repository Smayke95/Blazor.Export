namespace S95.Blazor.Export.Interfaces;

public interface IExportService
{
    byte[] ExportToCsv<T>(IEnumerable<T> data);
    byte[] ExportToCsv<T>(IEnumerable<string> columns, IEnumerable<T> data);
    byte[] ExportToExcel<T>(IEnumerable<T> data);
    byte[] ExportToExcel<T>(IEnumerable<string> columns, IEnumerable<T> data);
    byte[] ExportToPdf<T>(IEnumerable<T> data);
    byte[] ExportToPdf<T>(IEnumerable<string> columns, IEnumerable<T> data);
    byte[] ExportToWord<T>(IEnumerable<T> data);
    byte[] ExportToWord<T>(IEnumerable<string> columns, IEnumerable<T> data);
}