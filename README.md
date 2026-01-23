# Blazor.Export
Blazor.Export is a lightweight and flexible NuGet package that enables exporting table data from Blazor applications into various formats.

The package is designed to work seamlessly with both Blazor Server and Blazor WebAssembly projects.
It allows developers to easily convert and export data from any table component with minimal configuration and high performance.
Blazor.Export helps you integrate professional document export functionality into your Blazor apps with just a few lines of code.

## Features

- **Export to:** CSV (.csv), Excel (.xlsx), PDF (.pdf) & Word (.docx).
- **Blazor Server and WebAssembly support**
- **Localized columns:** Customize column names as you wish.
- **Minimal dependencies**

## How to use

1. Install the Blazor.Export NuGet package in your Blazor project.
2. In Program.cs (Server) add: ```builder.Services.AddExportServices();```
3. Add this JavaScript code:
```
function saveAsFile(filename, bytesBase64, contentType) {
    const link = document.createElement('a');
    link.download = filename;
    link.href = `data:${contentType};base64,${bytesBase64}`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
```
4. Inject IExportService into your component and call the desired export method.
```
private async Task ExportAsync(ExportFormat format)
{
    var data = new List<T> { };

    var bytes = format switch
    {
        ExportFormat.Csv => ExportService.ExportToCsv(data),
        ExportFormat.Excel => ExportService.ExportToExcel(data),
        ExportFormat.Pdf => ExportService.ExportToPdf(data),
        ExportFormat.Word => ExportService.ExportToWord(data),
        _ => ExportService.ExportToCsv(data)
    };

    var (fileName, contentType) = format switch
    {
        ExportFormat.Csv => ($"File_{DateTime.Now:yyyy-MM-dd_HH-mm}.csv", "text/csv"),
        ExportFormat.Excel => ($"File_{DateTime.Now:yyyy-MM-dd_HH-mm}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        ExportFormat.Pdf => ($"File_{DateTime.Now:yyyy-MM-dd_HH-mm}.pdf", "application/pdf"),
        ExportFormat.Word => ($"File_{DateTime.Now:yyyy-MM-dd_HH-mm}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        _ => ($"File_{DateTime.Now:yyyy-MM-dd_HH-mm}.csv", "text/csv")
    };

    await JSRuntime.InvokeVoidAsync("saveAsFile", fileName, Convert.ToBase64String(bytes), contentType);
}
```

3. If you using WebAssembly then add ExportController on Server:
```
public class ExportController(IExportService exportService) : ControllerBase
{
    private readonly IExportService ExportService = exportService;

    [HttpPost("csv")]
    public FileContentResult ExportToCsv([FromBody] List<ExportData> data)
    {
        var bytes = ExportService.ExportToCsv<ExportData>(data);
        return File(bytes, "text/csv", "export.csv");
    }

    [HttpPost("excel")]
    public FileContentResult ExportToExcel([FromBody] List<ExportData> data)
    {
        var bytes = ExportService.ExportToExcel<ExportData>(data);
        return File(bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "export.xlsx");
    }

    [HttpPost("pdf")]
    public FileContentResult ExportToPdf([FromBody] List<ExportData> data)
    {
        var bytes = ExportService.ExportToPdf<ExportData>(data);
        return File(bytes, "application/pdf", "export.pdf");
    }

    [HttpPost("word")]
    public FileContentResult ExportToWord([FromBody] List<ExportData> data)
    {
        var bytes = ExportService.ExportToWord<ExportData>(data);
        return File(bytes,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "export.docx");
    }
}
```

## Licence

Project is licensed under the [MIT License](https://github.com/Smayke95/Blazor.Export/blob/master/LICENSE)