using System.Collections;
using System.Reflection;
using System.Runtime.Serialization;

namespace S95.Blazor.Export.Helpers;

public static class TypeHelpers
{
    public static IEnumerable<PropertyInfo> GetExportProperties(this Type type)
        => type
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(x =>
                !x.IsDefined(typeof(IgnoreDataMemberAttribute), inherit: true) &&
                !(typeof(IEnumerable).IsAssignableFrom(x.PropertyType) && x.PropertyType != typeof(string))
            );
}