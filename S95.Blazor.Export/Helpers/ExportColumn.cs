using S95.Blazor.Export.Models;
using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace S95.Blazor.Export.Helpers;

internal class ExportColumn
{
    private readonly PropertyInfo Property;
    private readonly DisplayFormatAttribute? DisplayFormat;

    public ExportColumn(string name, PropertyInfo property)
    {
        Name = name;
        Property = property;
        DisplayFormat = property.GetCustomAttribute<DisplayFormatAttribute>();
    }

    public string Name { get; }

    public Type Type => Property.PropertyType;

    public object? GetValue(object? item) => Property.GetValue(item);

    public string GetText(object? item, ExportOptions options)
    {
        var value = GetValue(item);

        if (value is null)
            return DisplayFormat?.NullDisplayText ?? string.Empty;

        if (DisplayFormat?.DataFormatString is { Length: > 0 } format)
            return string.Format(options.Culture, format, value);

        return value switch
        {
            DateTime date when date.TimeOfDay == TimeSpan.Zero => date.ToString(options.DateFormat, options.Culture),
            DateTime date => date.ToString(options.DateTimeFormat, options.Culture),
            DateTimeOffset date => date.ToString(options.DateTimeFormat, options.Culture),
            DateOnly date => date.ToString(options.DateFormat, options.Culture),
            TimeOnly time => time.ToString(options.TimeFormat, options.Culture),
            IFormattable formattable => formattable.ToString(null, options.Culture),
            _ => value.ToString() ?? string.Empty
        };
    }
}