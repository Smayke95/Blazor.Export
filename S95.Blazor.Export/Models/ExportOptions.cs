using System.Globalization;

namespace S95.Blazor.Export.Models;

/// <summary>
/// Formatting of values in CSV, PDF and Word exports. Excel keeps native numbers and dates.
/// A <see cref="System.ComponentModel.DataAnnotations.DisplayFormatAttribute"/> on a property takes precedence.
/// </summary>
public class ExportOptions
{
    /// <summary>Culture used for numbers and dates. Defaults to <see cref="CultureInfo.InvariantCulture"/>.</summary>
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>Format for dates without a time part. Defaults to <c>yyyy-MM-dd</c>.</summary>
    public string DateFormat { get; set; } = "yyyy-MM-dd";

    /// <summary>Format for dates with a time part. Defaults to <c>yyyy-MM-dd HH:mm:ss</c>.</summary>
    public string DateTimeFormat { get; set; } = "yyyy-MM-dd HH:mm:ss";

    /// <summary>Format for <see cref="TimeOnly"/> values. Defaults to <c>HH:mm:ss</c>.</summary>
    public string TimeFormat { get; set; } = "HH:mm:ss";
}