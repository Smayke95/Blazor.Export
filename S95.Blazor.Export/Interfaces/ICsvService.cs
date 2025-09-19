namespace S95.Blazor.Export.Interfaces;

public interface ICsvService
{
    byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data);
}