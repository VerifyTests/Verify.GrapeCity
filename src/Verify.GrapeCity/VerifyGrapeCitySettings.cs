using GrapeCity.Documents.Pdf;
using GrapeCity.Documents.Word.Layout;
using PdfPage = GrapeCity.Documents.Pdf.Page;
using WordPage = GrapeCity.Documents.Word.Layout.Page;
namespace VerifyTests;

public static class VerifyGrapeCitySettings
{
    public static void PagesToInclude(this VerifySettings settings, int count)
    {
        settings.Context["VerifyGrapeCityPagesToInclude"] = count;
    }

    public static SettingsTask PagesToInclude(this SettingsTask settings, int count)
    {
        settings.CurrentSettings.PagesToInclude(count);
        return settings;
    }

    internal static int GetPagesToInclude(this IReadOnlyDictionary<string, object> settings, int count)
    {
        if (!settings.TryGetValue("VerifyGrapeCityPagesToInclude", out var value))
        {
            return count;
        }

        return Math.Min(count, (int)value);
    }

    static ImageOutputSettings imageOutputSettings = new();

    public static void WordImageOutputSettings(this VerifySettings settings, Func<WordPage, ImageOutputSettings> func)
    {
        settings.Context["VerifyGrapeCityImageOutputSettings"] = func;
    }

    public static SettingsTask WordImageOutputSettings(this SettingsTask settings, Func<WordPage, ImageOutputSettings> func)
    {
        settings.CurrentSettings.WordImageOutputSettings(func);
        return settings;
    }

    internal static ImageOutputSettings GetImageOutputSettings(this IReadOnlyDictionary<string, object> settings, WordPage page)
    {
        if (!settings.TryGetValue("VerifyGrapeCityImageOutputSettings", out var value))
        {
            return imageOutputSettings;
        }

        var func = (Func<WordPage, ImageOutputSettings>)value;
        return func(page);
    }

    static SaveAsImageOptions saveAsImageOptions = new();

    public static void PdfSaveAsImageOptions(this VerifySettings settings, Func<PdfPage, SaveAsImageOptions> func)
    {
        settings.Context["VerifyGrapeCityPdfSaveAsImageOptions"] = func;
    }

    public static SettingsTask PdfSaveAsImageOptions(this SettingsTask settings, Func<PdfPage, SaveAsImageOptions> func)
    {
        settings.CurrentSettings.PdfSaveAsImageOptions(func);
        return settings;
    }

    internal static SaveAsImageOptions GetPdfSaveAsImageOptions(this IReadOnlyDictionary<string, object> settings, PdfPage page)
    {
        if (!settings.TryGetValue("VerifyGrapeCityPdfSaveAsImageOptions", out var value))
        {
            return saveAsImageOptions;
        }

        var func = (Func<PdfPage, SaveAsImageOptions>)value;
        return func(page);
    }
}