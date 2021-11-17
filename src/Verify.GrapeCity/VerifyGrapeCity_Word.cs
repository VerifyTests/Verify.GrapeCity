
using GrapeCity.Documents.Word;
using GrapeCity.Documents.Word.Layout;

namespace VerifyTests;

public static partial class VerifyGrapeCity
{
    static ConversionResult ConvertWord(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        var document = new GcWordDocument();
        document.Load(stream);
        return ConvertWord(document, settings);
    }

    static ConversionResult ConvertWord(GcWordDocument document, IReadOnlyDictionary<string, object> settings)
    {
        return new(GetInfo(document), GetWordStreams(document, settings).ToList());
    }

    static object GetInfo(GcWordDocument document)
    {
        var properties = document.Settings.BuiltinProperties;
        return new
        {
            properties.Author,
            properties.Subject,
            properties.Title,
            properties.Pages,
            properties.ApplicationName,
            properties.AplicationVersion,
            properties.Category,
            properties.Comments,
            properties.Company,
            properties.CreatedTime,
            properties.Keywords,
            properties.Language,
            properties.Security,
            properties.RevisionNumber,
            properties.Version,
            DefaultFont = document.Styles.DefaultFont.Name,
        };
    }

    static IEnumerable<Target> GetWordStreams(GcWordDocument document, IReadOnlyDictionary<string, object> settings)
    {
        using var layout = new GcWordLayout(document);
        var pagesToInclude = settings.GetPagesToInclude(layout.Pages.Count);

        for (var pageIndex = 0; pageIndex < pagesToInclude; pageIndex++)
        {
            var page = layout.Pages[pageIndex];
            var stream = new MemoryStream();
            page.SaveAsPng(stream, settings.GetImageOutputSettings(page));
            yield return new("png", stream);
        }
    }
}