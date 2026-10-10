using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;
using S95.Blazor.Export.Models;
using System.Text;

namespace S95.Blazor.Export.Services;

/// <summary>
/// Exports items to a UTF-8 CSV file with a BOM, quoting values that contain commas, quotes or line breaks.
/// </summary>
public class CsvService : IBaseExportService
{
    private const string Separator = ",";
    private const string NewLine = "\r\n";
    private static readonly char[] SpecialCharacters = { ',', '"', '\r', '\n' };
    private static readonly UTF8Encoding Encoding = new(encoderShouldEmitUTF8Identifier: true);
    private readonly ExportOptions Options;

    /// <summary>Creates the service with default <see cref="ExportOptions"/>.</summary>
    public CsvService() : this(new ExportOptions()) { }

    /// <summary>Creates the service with the given <see cref="ExportOptions"/>.</summary>
    public CsvService(ExportOptions options)
    {
        Options = options;
    }

    /// <inheritdoc />
    public byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        var exportColumns = typeof(T).GetExportColumns(columns);

        var stringBuilder = new StringBuilder();
        stringBuilder.Append(string.Join(Separator, exportColumns.Select(c => Escape(c.Name)))).Append(NewLine);

        foreach (var item in data)
        {
            var values = exportColumns.Select(c => Escape(c.GetText(item, Options)));
            stringBuilder.Append(string.Join(Separator, values)).Append(NewLine);
        }

        return Encoding
            .GetPreamble()
            .Concat(Encoding.GetBytes(stringBuilder.ToString()))
            .ToArray();
    }

    private static string Escape(string? value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        return value.IndexOfAny(SpecialCharacters) >= 0
            ? $"\"{value.Replace("\"", "\"\"")}\""
            : value;
    }
}
