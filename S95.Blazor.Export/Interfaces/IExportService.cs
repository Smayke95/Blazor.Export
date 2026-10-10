namespace S95.Blazor.Export.Interfaces;

/// <summary>
/// Exports a collection of items to CSV, Excel, PDF or Word. Every public readable property becomes a column,
/// except collections and properties marked with <see cref="System.Runtime.Serialization.IgnoreDataMemberAttribute"/>.
/// </summary>
public interface IExportService
{
    /// <summary>Exports the items to a UTF-8 CSV file. Column names come from <c>[Display]</c>, <c>[DisplayName]</c> or the property name.</summary>
    byte[] ExportToCsv<T>(IEnumerable<T> data);

    /// <summary>Exports the items to a UTF-8 CSV file with the given column names, in property order.</summary>
    byte[] ExportToCsv<T>(IEnumerable<string> columns, IEnumerable<T> data);

    /// <summary>Exports the items to an Excel (.xlsx) file. Column names come from <c>[Display]</c>, <c>[DisplayName]</c> or the property name.</summary>
    byte[] ExportToExcel<T>(IEnumerable<T> data);

    /// <summary>Exports the items to an Excel (.xlsx) file with the given column names, in property order.</summary>
    byte[] ExportToExcel<T>(IEnumerable<string> columns, IEnumerable<T> data);

    /// <summary>Exports the items to a landscape A4 PDF file. Column names come from <c>[Display]</c>, <c>[DisplayName]</c> or the property name.</summary>
    byte[] ExportToPdf<T>(IEnumerable<T> data);

    /// <summary>Exports the items to a landscape A4 PDF file with the given column names, in property order.</summary>
    byte[] ExportToPdf<T>(IEnumerable<string> columns, IEnumerable<T> data);

    /// <summary>Exports the items to a landscape Word (.docx) file. Column names come from <c>[Display]</c>, <c>[DisplayName]</c> or the property name.</summary>
    byte[] ExportToWord<T>(IEnumerable<T> data);

    /// <summary>Exports the items to a landscape Word (.docx) file with the given column names, in property order.</summary>
    byte[] ExportToWord<T>(IEnumerable<string> columns, IEnumerable<T> data);
}