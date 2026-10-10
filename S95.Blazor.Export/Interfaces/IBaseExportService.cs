namespace S95.Blazor.Export.Interfaces;

/// <summary>
/// Exports a collection of items to a single file format.
/// </summary>
public interface IBaseExportService
{
    /// <summary>
    /// Exports the items and returns the file content. Column names are matched to properties by position;
    /// missing names fall back to the property name and extra names are ignored.
    /// </summary>
    byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data);
}