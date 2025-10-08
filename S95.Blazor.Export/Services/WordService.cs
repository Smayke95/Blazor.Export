using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace S95.Blazor.Export.Services;

public class WordService : IBaseExportService
{
    public byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        using var stream = new MemoryStream();
        using var document = DocX.Create(stream);
        document.PageLayout.Orientation = Orientation.Landscape;

        document.MarginLeft = 30f;
        document.MarginRight = 30f;
        document.MarginTop = 30f;
        document.MarginBottom = 30f;

        var properties = typeof(T).GetExportProperties().ToList();
        var items = data.ToList();

        var table = document.AddTable(data.Count() + 1, properties.Count);

        for (int i = 0; i < columns.Count(); i++)
            table.Rows[0].Cells[i].Paragraphs[0].Append(columns.ToList()[i]);

        for (int i = 1; i <= items.Count; i++)
        {
            for (int j = 0; j < properties.Count; j++)
                table.Rows[i].Cells[j].Paragraphs[0].Append(properties[j].GetValue(items[i - 1])?.ToString() ?? string.Empty);
        }

        document.InsertTable(table);
        document.Save();

        return stream.ToArray();
    }
}