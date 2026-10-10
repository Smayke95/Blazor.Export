namespace S95.Blazor.Export.Models;

/// <summary>
/// File formats supported by <see cref="Interfaces.IExportService"/>.
/// </summary>
public enum ExportFormat
{
    /// <summary>Comma separated values (.csv).</summary>
    Csv,

    /// <summary>Excel workbook (.xlsx).</summary>
    Excel,

    /// <summary>PDF document (.pdf).</summary>
    Pdf,

    /// <summary>Word document (.docx).</summary>
    Word
}