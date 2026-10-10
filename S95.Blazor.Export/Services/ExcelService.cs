using ClosedXML.Excel;
using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;
using S95.Blazor.Export.Models;
using System.Data;

namespace S95.Blazor.Export.Services;

/// <summary>
/// Exports items to an Excel (.xlsx) table, keeping numbers, dates and booleans as native Excel values.
/// </summary>
public class ExcelService : IBaseExportService
{
    private static readonly HashSet<Type> NativeTypes = new()
    {
        typeof(bool),
        typeof(byte),
        typeof(sbyte),
        typeof(short),
        typeof(ushort),
        typeof(int),
        typeof(uint),
        typeof(long),
        typeof(ulong),
        typeof(float),
        typeof(double),
        typeof(decimal),
        typeof(DateTime),
        typeof(DateTimeOffset),
        typeof(TimeSpan),
        typeof(string)
    };
    private readonly ExportOptions Options;

    /// <summary>Creates the service with default <see cref="ExportOptions"/>.</summary>
    public ExcelService() : this(new ExportOptions()) { }

    /// <summary>Creates the service with the given <see cref="ExportOptions"/>.</summary>
    public ExcelService(ExportOptions options)
    {
        Options = options;
    }

    /// <inheritdoc />
    public byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");

        var exportColumns = typeof(T).GetExportColumns(columns);
        var columnNames = GetUniqueNames(exportColumns.Select(c => c.Name));

        var dataTable = new DataTable();

        for (int i = 0; i < exportColumns.Count; i++)
            dataTable.Columns.Add(columnNames[i], GetColumnType(exportColumns[i].Type));

        foreach (var item in data)
        {
            var row = dataTable.NewRow();

            for (int i = 0; i < exportColumns.Count; i++)
                row[i] = GetCellValue(exportColumns[i], item, dataTable.Columns[i].DataType);

            dataTable.Rows.Add(row);
        }

        worksheet.Cell(1, 1).InsertTable(dataTable);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;

        return stream.ToArray();
    }

    private static Type GetColumnType(Type type)
    {
        var underlyingType = Nullable.GetUnderlyingType(type) ?? type;

        if (underlyingType == typeof(DateOnly))
            return typeof(DateTime);

        if (underlyingType == typeof(TimeOnly))
            return typeof(TimeSpan);

        return NativeTypes.Contains(underlyingType) ? underlyingType : typeof(string);
    }

    private object GetCellValue(ExportColumn column, object? item, Type columnType)
        => column.GetValue(item) switch
        {
            null when columnType == typeof(string) && column.GetText(item, Options) is { Length: > 0 } text => text,
            null => DBNull.Value,
            DateOnly date => date.ToDateTime(TimeOnly.MinValue),
            TimeOnly time => time.ToTimeSpan(),
            _ when columnType == typeof(string) => column.GetText(item, Options),
            var value => value
        };

    private static List<string> GetUniqueNames(IEnumerable<string> names)
    {
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        return names
            .Select((name, i) =>
            {
                var baseName = string.IsNullOrWhiteSpace(name) ? $"Column{i + 1}" : name;
                var uniqueName = baseName;

                for (var number = 2; !usedNames.Add(uniqueName); number++)
                    uniqueName = $"{baseName}{number}";

                return uniqueName;
            })
            .ToList();
    }
}