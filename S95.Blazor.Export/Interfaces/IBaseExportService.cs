namespace S95.Blazor.Export.Interfaces;

public interface IBaseExportService
{
    byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data);
}