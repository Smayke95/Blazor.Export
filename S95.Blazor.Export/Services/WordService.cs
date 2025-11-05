using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using S95.Blazor.Export.Helpers;
using S95.Blazor.Export.Interfaces;

namespace S95.Blazor.Export.Services;

public class WordService : IBaseExportService
{
    public byte[] Export<T>(IEnumerable<string> columns, IEnumerable<T> data)
    {
        using var stream = new MemoryStream();
        var document = new XWPFDocument();

        var section = document.Document.body.AddNewSectPr();

        var pageSize = section.AddNewPgSz();
        pageSize.orient = ST_PageOrientation.landscape;
        pageSize.w = 16838;
        pageSize.h = 11906;

        var pageMargin = section.AddPageMar();
        pageMargin.left = 450;
        pageMargin.right = 450;
        pageMargin.top = 450;
        pageMargin.bottom = 450;

        var properties = typeof(T).GetExportProperties().ToList();
        var items = data.ToList();

        var table = document.CreateTable(items.Count + 1, properties.Count);

        var layout = table.GetCTTbl().tblPr.AddNewTblLayout();
        layout.type = ST_TblLayoutType.@fixed;
        table.Width = 5000;

        for (var i = 0; i < columns.Count(); i++)
        {
            table.SetColumnWidth(i, (ulong)(5000 / columns.Count()));
            table.GetRow(0).GetCell(i).SetText(columns.ElementAt(i));
        }

        for (var i = 0; i < items.Count; i++)
            for (var y = 0; y < properties.Count; y++)
                table.GetRow(i + 1).GetCell(y).SetText(properties[y].GetValue(items[i])?.ToString() ?? string.Empty);

        document.Write(stream);
        return stream.ToArray();
    }
}