using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Excel.Drawing;

namespace VerifyTests;

public static partial class VerifyGrapeCity
{
    static ConversionResult ConvertExcel(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        var book = new Workbook();
        book.Open(stream);
        return ConvertExcel(book, settings);
    }

    static ConversionResult ConvertExcel(Workbook book, IReadOnlyDictionary<string, object> settings)
    {
        var info = GetInfo(book);
        return new(info, GetExcelStreams(book).ToList());
    }

    static object GetInfo(Workbook book)
    {
        return new
        {
            book.BuiltInDocumentProperties,
            book.CustomDocumentProperties,
            book.AllowDynamicArray,
            book.AutoParse,
            book.AutoRoundValue,
            book.DefaultTableStyle,
            book.EnableCalculation,
            book.Name,
            book.ProtectWindows
        };
    }

    static IEnumerable<Target> GetExcelStreams(Workbook book)
    {
        foreach (var sheet in book.Worksheets)
        {
            var setup = sheet.PageSetup;
            setup.PrintGridlines = true;
            setup.LeftMargin = 0;
            setup.TopMargin = 0;
            setup.RightMargin = 0;
            setup.BottomMargin = 0;
            var stream = new MemoryStream();
            sheet.ToImage(stream, ImageType.PNG);
            yield return new("png", stream);
        }
    }
}