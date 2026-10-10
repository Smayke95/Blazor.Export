using S95.Blazor.Export.Services;
using Xunit;

namespace S95.Blazor.Export.Tests;

public class AssemblyTests
{
    [Fact]
    public void Assembly_EmbedsOnlyFonts()
    {
        var resources = typeof(ExportService).Assembly.GetManifestResourceNames().Order();

        Assert.Equal(["S95.Blazor.Export.Fonts.LiberationSans-Bold.ttf", "S95.Blazor.Export.Fonts.LiberationSans-Regular.ttf"], resources);
    }
}