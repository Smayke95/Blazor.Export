using ClosedXML.Excel;
using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;
using System.Data;

namespace S95.Blazor.Export.Services;

public class ExcelService : IBaseExportService
{
    public byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");

        var properties = typeof(T).GetExportProperties().ToList();
        var items = data.ToList();

        var dataTable = new DataTable();

        foreach (var column in columns)
            dataTable.Columns.Add(column);

        for (int i = 0; i < items.Count; i++)
        {
            dataTable.Rows.Add();

            for (int j = 0; j < properties.Count; j++)
                dataTable.Rows[i][j] = properties[j].GetValue(items[i]);
        }

        worksheet.Cell(1, 1).InsertTable(dataTable);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;

        return stream.ToArray();
    }
}