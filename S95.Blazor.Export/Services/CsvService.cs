using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;
using System.Text;

namespace S95.Blazor.Export.Services;

public class CsvService : IBaseExportService
{
    public byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        var stringBuilder = new StringBuilder();
        stringBuilder.AppendLine(string.Join(",", columns));

        var properties = typeof(T).GetExportProperties();

        foreach (var item in data)
        {
            var values = properties.Select(p => p.GetValue(item, null)?.ToString()?.Replace(",", ".") ?? "");
            stringBuilder.AppendLine(string.Join(",", values));
        }

        return Encoding.ASCII.GetBytes(stringBuilder.ToString());
    }
}