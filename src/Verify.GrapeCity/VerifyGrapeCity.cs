using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Pdf;
using GrapeCity.Documents.Word;

namespace VerifyTests
{
    public static partial class VerifyGrapeCity
    {
        public static void Initialize()
        {
            VerifierSettings.RegisterFileConverter("xlsx", ConvertExcel);
            VerifierSettings.RegisterFileConverter("xls", ConvertExcel);
            VerifierSettings.RegisterFileConverter<Workbook>(ConvertExcel);

            VerifierSettings.RegisterFileConverter("pdf", ConvertPdf);
            VerifierSettings.RegisterFileConverter<GcPdfDocument>(ConvertPdf);

            VerifierSettings.RegisterFileConverter("docx", ConvertWord);
            VerifierSettings.RegisterFileConverter("doc", ConvertWord);
            VerifierSettings.RegisterFileConverter<GcWordDocument>(ConvertWord);
        }
    }
}