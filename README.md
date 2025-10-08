# Blazor.Export
Blazor.Export is a lightweight and flexible NuGet package that enables exporting table data from Blazor applications into multiple file formats.

Supported export formats:
- Word (.docx)
- Excel (.xlsx)
- PDF (.pdf)
- CSV (.csv)

The package is designed to work seamlessly with both Blazor Server and Blazor WebAssembly projects.
It allows developers to easily convert and export data from any table component with minimal configuration and high performance.

Blazor.Export helps you integrate professional document export functionality into your Blazor apps with just a few lines of code.


## How to use

1. Install the Blazor.Export NuGet package in your Blazor project.
2. In Program.cs add: ```builder.Services.AddExportServices();```
3. Add ExportController:
```
public class ExportController(IExportService exportService) : ControllerBase
{
    private readonly IExportService ExportService = exportService;

    [HttpPost("csv")]
    public FileContentResult ExportToCsv([FromBody] List<HouseCompare> data)
    {
        var bytes = ExportService.ExportToCsv<HouseCompare>(data);
        return File(bytes, "text/csv", "export.csv");
    }

    [HttpPost("excel")]
    public FileContentResult ExportToExcel([FromBody] List<HouseCompare> data)
    {
        var bytes = ExportService.ExportToExcel<HouseCompare>(data);
        return File(bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "export.xlsx");
    }

    [HttpPost("pdf")]
    public FileContentResult ExportToPdf([FromBody] List<HouseCompare> data)
    {
        var bytes = ExportService.ExportToPdf<HouseCompare>(data);
        return File(bytes, "application/pdf", "export.pdf");
    }

    [HttpPost("word")]
    public FileContentResult ExportToWord([FromBody] List<HouseCompare> data)
    {
        var bytes = ExportService.ExportToWord<HouseCompare>(data);
        return File(bytes,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "export.docx");
    }
}
```
4. Include javascript file
5. Add button which calls http method
6. Start using export functionality in your components.


## Licence

Project is licensed under the [MIT License](https://github.com/Smayke95/Blazor.Export/blob/master/LICENSE)