using Microsoft.Extensions.DependencyInjection;
using S95.Blazor.Export.Models;
using System.Reflection;

namespace S95.Blazor.Export.Extensions;

/// <summary>
/// Registers the export services in the dependency injection container.
/// </summary>
public static class ServiceExtensions
{
    /// <summary>Registers <see cref="Interfaces.IExportService"/> and the format services with default <see cref="ExportOptions"/>.</summary>
    public static void AddExportServices(this IServiceCollection serviceCollection)
        => serviceCollection.AddExportServices(_ => { });

    /// <summary>Registers <see cref="Interfaces.IExportService"/> and the format services with configured <see cref="ExportOptions"/>.</summary>
    public static void AddExportServices(this IServiceCollection serviceCollection, Action<ExportOptions> configure)
    {
        var options = new ExportOptions();
        configure(options);
        serviceCollection.AddSingleton(options);

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