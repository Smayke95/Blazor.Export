using System.Collections;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using System.Runtime.Serialization;

namespace S95.Blazor.Export.Helpers;

/// <summary>
/// Reflection helpers used by the exports.
/// </summary>
public static class TypeHelpers
{
    /// <summary>
    /// Returns the properties that become columns: public, readable and non-indexed properties,
    /// except collections and properties marked with <see cref="IgnoreDataMemberAttribute"/>.
    /// </summary>
    public static IEnumerable<PropertyInfo> GetExportProperties(this Type type)
        => type
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(x =>
                x.GetMethod is { IsPublic: true } &&
                x.GetIndexParameters().Length == 0 &&
                !x.IsDefined(typeof(IgnoreDataMemberAttribute), inherit: true) &&
                !(typeof(IEnumerable).IsAssignableFrom(x.PropertyType) && x.PropertyType != typeof(string))
            );

    internal static string GetColumnName(this PropertyInfo property)
        => property.GetCustomAttribute<DisplayAttribute>()?.GetName() ??
           property.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ??
           property.Name;

    internal static List<ExportColumn> GetExportColumns(this Type type, IEnumerable<string> columns)
    {
        var names = columns.ToList();

        return type
            .GetExportProperties()
            .Select((property, i) => new ExportColumn(i < names.Count ? names[i] : property.GetColumnName(), property))
            .ToList();
    }
}