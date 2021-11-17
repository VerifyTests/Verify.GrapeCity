using GrapeCity.Documents.Pdf;

namespace VerifyTests;

public static partial class VerifyGrapeCity
{
    static ConversionResult ConvertPdf(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        var document = new GcPdfDocument();
        document.Load(stream);
        return ConvertPdf(document, settings);
    }

    static ConversionResult ConvertPdf(GcPdfDocument document, IReadOnlyDictionary<string, object> settings)
    {
        var info = document.DocumentInfo;
        return new(
            new
            {
                info.Title,
                info.Author,
                CreationDate = info.CreationDate?.DateTimeOffset,
                ModifyDate = info.ModifyDate?.DateTimeOffset,
                info.Subject,
                info.Keywords,
                info.Creator,
                info.Producer,
            },
            GetPdfStreams(document, settings).ToList());
    }

    static IEnumerable<Target> GetPdfStreams(GcPdfDocument document, IReadOnlyDictionary<string, object> settings)
    {
        var pagesToInclude = settings.GetPagesToInclude(document.Pages.Count);
        for (var index = 0; index < pagesToInclude; index++)
        {
            var page = document.Pages[index];
            var stream = new MemoryStream();
            page.SaveAsPng(stream, settings.GetPdfSaveAsImageOptions(page));
            yield return new("png", stream);
        }
    }
}