using ClosedXML.Excel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;
using System.Reflection;
using Xceed.Words.NET;
using Document = QuestPDF.Fluent.Document;

namespace S95.Blazor.Export.Services;

public class ExportService(
    ICsvService csvService)
    : IExportService
{
    private readonly ICsvService CsvService = csvService;

    public byte[] ExportToCsv<T>(IEnumerable<T> data)
    {
        var columns = typeof(T)
            .GetExportProperties()
            .Select(x => x.CustomAttributes.FirstOrDefault(y => y.AttributeType.Name == "DisplayAttribute")?.NamedArguments.FirstOrDefault().TypedValue.ToString() ?? x.Name);

        return CsvService.Export(columns, data);
    }

    public byte[] ExportToCsv<T>(IEnumerable<string> columns, IEnumerable<T> data)
        => CsvService.Export(columns, data);

    public byte[] ExportToExcel<T>(IEnumerable<T> data)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");
        worksheet.Cell(1, 1).InsertTable(data);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;

        return stream.ToArray();
    }

    public byte[] ExportToPdf<T>(IEnumerable<T> data)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

        var document = Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Content()
                    .Table(table =>
                    {
                        // Header
                        table.ColumnsDefinition(columns =>
                        {
                            for (int i = 0; i < properties.Length; i++)
                                columns.RelativeColumn();
                        });

                        table.Header(header =>
                        {
                            for (int i = 0; i < properties.Length; i++)
                                header.Cell().Element(CellStyle).Text(properties[i].Name);
                        });

                        // Data
                        foreach (var item in data)
                        {
                            table.Cell().Element(CellStyle).Text(string.Join(" | ", properties.Select(p => p.GetValue(item, null)?.ToString() ?? "")));
                        }
                    });
            });
        });

        return document.GeneratePdf();
    }

    public byte[] ExportToWord<T>(IEnumerable<T> data)
    {
        using var ms = new MemoryStream();
        using (var doc = DocX.Create(ms))
        {
            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            // Header row
            var table = doc.AddTable(data.Count() + 1, properties.Length);
            for (int i = 0; i < properties.Length; i++)
                table.Rows[0].Cells[i].Paragraphs[0].Append(properties[i].Name);

            // Data rows
            int row = 1;
            foreach (var item in data)
            {
                for (int col = 0; col < properties.Length; col++)
                {
                    var value = properties[col].GetValue(item, null)?.ToString() ?? "";
                    table.Rows[row].Cells[col].Paragraphs[0].Append(value);
                }
                row++;
            }

            doc.InsertTable(table);
            doc.Save();
        }
        return ms.ToArray();
    }

    static QuestPDF.Infrastructure.IContainer CellStyle(QuestPDF.Infrastructure.IContainer container)
    {
        return container
            .Padding(5)
            .Border(1)
            .BorderColor(Colors.Grey.Lighten2);
    }

    public byte[] ExportToExcel<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        throw new NotImplementedException();
    }

    public byte[] ExportToPdf<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        throw new NotImplementedException();
    }

    public byte[] ExportToWord<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        throw new NotImplementedException();
    }
}