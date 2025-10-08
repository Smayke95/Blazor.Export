using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;

namespace S95.Blazor.Export.Services;

public class PdfService : IBaseExportService
{
    public byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        var document = new PdfDocument();
        var page = document.AddPage();
        page.Orientation = PageOrientation.Landscape;

        var gfx = XGraphics.FromPdfPage(page);

        var font = new XFont("Arial", 12, XFontStyleEx.Regular);
        var boldFont = new XFont("Arial", 12, XFontStyleEx.Bold);

        double margin = 40;
        double y = margin;
        double x = margin;
        double cellWidth = 150;
        double cellHeight = 25;

        var properties = typeof(T).GetExportProperties().ToList();
        var items = data.ToList();

        for (int i = 0; i < columns.ToList().Count; i++)
        {
            gfx.DrawRectangle(XPens.Black, x + i * cellWidth, y, cellWidth, cellHeight);
            gfx.DrawString(columns.ToList()[i], boldFont, XBrushes.Black, new XRect(x + i * cellWidth, y, cellWidth, cellHeight), XStringFormats.Center);
        }

        y += cellHeight;

        foreach (var item in items)
        {
            for (int j = 0; j < properties.Count; j++)
            {
                var value = properties[j].GetValue(item)?.ToString() ?? string.Empty;
                gfx.DrawRectangle(XPens.Black, x + j * cellWidth, y, cellWidth, cellHeight);
                gfx.DrawString(value, font, XBrushes.Black, new XRect(x + j * cellWidth, y, cellWidth, cellHeight), XStringFormats.Center);
            }

            y += cellHeight;

            if (XUnit.FromPoint(y + cellHeight) > XUnit.FromPoint(page.Height.Value - margin))
            {
                page = document.AddPage();
                page.Orientation = PageOrientation.Landscape;
                gfx = XGraphics.FromPdfPage(page);
                y = margin;
            }
        }

        using var stream = new MemoryStream();
        document.Save(stream, false);
        return stream.ToArray();
    }
}