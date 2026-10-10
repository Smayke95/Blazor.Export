using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using S95.Blazor.Export.Fonts;
using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;
using S95.Blazor.Export.Models;

namespace S95.Blazor.Export.Services;

/// <summary>
/// Exports items to a landscape A4 PDF table using the embedded Liberation Sans font, with wrapped text and a header on every page.
/// </summary>
public class PdfService : IBaseExportService
{
    private const double Margin = 40;
    private const double CellPadding = 4;
    private const double FontSize = 10;
    private readonly ExportOptions Options;

    static PdfService()
    {
        ExportFontResolver.EnsureRegistered();
    }

    /// <summary>Creates the service with default <see cref="ExportOptions"/>.</summary>
    public PdfService() : this(new ExportOptions()) { }

    /// <summary>Creates the service with the given <see cref="ExportOptions"/>.</summary>
    public PdfService(ExportOptions options)
    {
        Options = options;
    }

    /// <inheritdoc />
    public byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        var exportColumns = typeof(T).GetExportColumns(columns);

        var header = exportColumns.Select(c => c.Name).ToList();
        var rows = data.Select(item => exportColumns.Select(c => c.GetText(item, Options)).ToList()).ToList();

        var font = new XFont(ExportFontResolver.FamilyName, FontSize, XFontStyleEx.Regular);
        var boldFont = new XFont(ExportFontResolver.FamilyName, FontSize, XFontStyleEx.Bold);
        var lineHeight = font.GetHeight();

        var document = new PdfDocument();
        var page = AddPage(document);
        var gfx = XGraphics.FromPdfPage(page);

        var tableWidth = page.Width.Point - 2 * Margin;
        var maxRowHeight = page.Height.Point - 2 * Margin;

        var naturalWidths = header
            .Select((name, i) => Math.Max(
                MeasureLines(gfx, name, boldFont),
                rows.Select(row => MeasureLines(gfx, row[i], font)).DefaultIfEmpty(0).Max()) + 2 * CellPadding)
            .ToList();

        var widths = GetColumnWidths(naturalWidths, tableWidth);

        var headerLines = WrapRow(gfx, header, widths, boldFont);
        var headerHeight = GetRowHeight(headerLines, lineHeight, maxRowHeight);

        var y = Margin;

        if (header.Count > 0)
            y = DrawRow(gfx, headerLines, widths, y, headerHeight, lineHeight, boldFont, XBrushes.LightGray);

        foreach (var row in rows)
        {
            var lines = WrapRow(gfx, row, widths, font);
            var rowHeight = GetRowHeight(lines, lineHeight, maxRowHeight - headerHeight);

            if (y + rowHeight > page.Height.Point - Margin)
            {
                gfx.Dispose();
                page = AddPage(document);
                gfx = XGraphics.FromPdfPage(page);

                y = DrawRow(gfx, headerLines, widths, Margin, headerHeight, lineHeight, boldFont, XBrushes.LightGray);
            }

            y = DrawRow(gfx, lines, widths, y, rowHeight, lineHeight, font, null);
        }

        gfx.Dispose();

        using var stream = new MemoryStream();
        document.Save(stream, false);
        return stream.ToArray();
    }

    internal static double[] GetColumnWidths(IReadOnlyList<double> naturalWidths, double availableWidth)
    {
        var widths = new double[naturalWidths.Count];

        if (widths.Length == 0)
            return widths;

        var totalWidth = naturalWidths.Sum();

        if (totalWidth <= availableWidth)
        {
            for (var i = 0; i < widths.Length; i++)
                widths[i] = totalWidth > 0 ? naturalWidths[i] * availableWidth / totalWidth : availableWidth / widths.Length;

            return widths;
        }

        var remaining = Enumerable.Range(0, widths.Length).ToList();
        var remainingWidth = availableWidth;

        while (remaining.Count > 0)
        {
            var share = remainingWidth / remaining.Count;
            var fitting = remaining.Where(i => naturalWidths[i] <= share).ToList();

            if (fitting.Count == 0)
            {
                foreach (var i in remaining)
                    widths[i] = share;

                break;
            }

            foreach (var i in fitting)
            {
                widths[i] = naturalWidths[i];
                remainingWidth -= naturalWidths[i];
                remaining.Remove(i);
            }
        }

        return widths;
    }

    internal static List<string> WrapText(string text, double maxWidth, Func<string, double> measure)
    {
        var lines = new List<string>();

        foreach (var paragraph in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
        {
            var line = string.Empty;

            foreach (var word in paragraph.Split(' '))
            {
                var candidate = line.Length == 0 ? word : $"{line} {word}";

                if (measure(candidate) <= maxWidth)
                {
                    line = candidate;
                    continue;
                }

                if (line.Length > 0)
                    lines.Add(line);

                line = word;

                while (line.Length > 1 && measure(line) > maxWidth)
                {
                    var length = line.Length - 1;

                    while (length > 1 && measure(line[..length]) > maxWidth)
                        length--;

                    lines.Add(line[..length]);
                    line = line[length..];
                }
            }

            lines.Add(line);
        }

        return lines;
    }

    private static PdfPage AddPage(PdfDocument document)
    {
        var page = document.AddPage();
        page.Size = PageSize.A4;
        page.Orientation = PageOrientation.Landscape;
        return page;
    }

    private static double MeasureLines(XGraphics gfx, string text, XFont font)
        => text
            .Replace("\r\n", "\n")
            .Split('\n')
            .Max(line => gfx.MeasureString(line, font).Width);

    private static List<List<string>> WrapRow(XGraphics gfx, IReadOnlyList<string> cells, double[] widths, XFont font)
        => cells
            .Select((text, i) => WrapText(text, widths[i] - 2 * CellPadding, value => gfx.MeasureString(value, font).Width))
            .ToList();

    private static double GetRowHeight(List<List<string>> lines, double lineHeight, double maxHeight)
    {
        var height = lines.Select(cell => cell.Count).DefaultIfEmpty(1).Max() * lineHeight + 2 * CellPadding;
        return Math.Min(height, maxHeight);
    }

    private static double DrawRow(XGraphics gfx, List<List<string>> lines, double[] widths, double y, double height, double lineHeight, XFont font, XBrush? background)
    {
        var x = Margin;
        var maxLines = (int)Math.Floor((height - 2 * CellPadding) / lineHeight);

        for (var i = 0; i < widths.Length; i++)
        {
            var rectangle = new XRect(x, y, widths[i], height);

            if (background is not null)
                gfx.DrawRectangle(XPens.Black, background, rectangle);
            else
                gfx.DrawRectangle(XPens.Black, rectangle);

            var lineY = y + CellPadding;

            foreach (var line in lines[i].Take(maxLines))
            {
                gfx.DrawString(line, font, XBrushes.Black, new XRect(x + CellPadding, lineY, widths[i] - 2 * CellPadding, lineHeight), XStringFormats.TopLeft);
                lineY += lineHeight;
            }

            x += widths[i];
        }

        return y + height;
    }
}