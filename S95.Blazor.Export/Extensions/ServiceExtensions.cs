using Microsoft.Extensions.DependencyInjection;
using PdfSharp.Fonts;
using System.Reflection;

namespace S95.Blazor.Export.Extensions;

public static class ServiceExtensions
{
    public static void AddExportServices(this IServiceCollection serviceCollection)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;

        var services = Assembly
            .GetExecutingAssembly()
            .GetTypes()
            .Where(x =>
                x.IsClass &&
                !x.IsAbstract &&
                x.Name.EndsWith("Service")
             );

        foreach (var service in services)
        {
            serviceCollection.AddScoped(service);

            var interfaces = service
                .GetInterfaces()
                .Where(x => !x.IsGenericType);

            foreach (var inter in interfaces)
                serviceCollection.AddScoped(inter, service);
        }
    }
}