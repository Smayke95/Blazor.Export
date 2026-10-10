using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;
using S95.Blazor.Export.Models;

namespace S95.Blazor.Export.Services;

/// <summary>
/// Exports items to a landscape Word (.docx) table.
/// </summary>
public class WordService : IBaseExportService
{
    private const int PageWidth = 16838;
    private const int PageHeight = 11906;
    private const int PageMargin = 450;
    private readonly ExportOptions Options;

    /// <summary>Creates the service with default <see cref="ExportOptions"/>.</summary>
    public WordService() : this(new ExportOptions()) { }

    /// <summary>Creates the service with the given <see cref="ExportOptions"/>.</summary>
    public WordService(ExportOptions options)
    {
        Options = options;
    }

    /// <inheritdoc />
    public byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        using var stream = new MemoryStream();

        var exportColumns = typeof(T).GetExportColumns(columns);
        var items = data.ToList();

        var columnWidth = (PageWidth - 2 * PageMargin) / Math.Max(exportColumns.Count, 1);

        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();

            var table = new Table(
                new TableProperties(
                    new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new RightBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }),
                    new TableLayout { Type = TableLayoutValues.Fixed }),
                new TableGrid(exportColumns.Select(_ => new GridColumn { Width = columnWidth.ToString() }))
            );

            table.Append(CreateRow(exportColumns.Select(c => c.Name), columnWidth));

            foreach (var item in items)
                table.Append(CreateRow(exportColumns.Select(c => c.GetText(item, Options)), columnWidth));

            mainPart.Document = new Document(new Body(
                table,
                new SectionProperties(
                    new PageSize { Width = PageWidth, Height = PageHeight, Orient = PageOrientationValues.Landscape },
                    new PageMargin { Left = PageMargin, Right = PageMargin, Top = PageMargin, Bottom = PageMargin }))
                );
        }

        return stream.ToArray();
    }

    private static TableRow CreateRow(IEnumerable<string> values, int columnWidth) =>
        new(values.Select(value => new TableCell(
            new TableCellProperties(new TableCellWidth { Width = columnWidth.ToString(), Type = TableWidthUnitValues.Dxa }),
            new Paragraph(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve })))));
}