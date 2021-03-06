using VerifyTests;
using VerifyNUnit;
using NUnit.Framework;

[TestFixture]
public class Samples
{
    #region VerifyPdf

    [Test]
    public Task VerifyPdf()
    {
        return Verifier.VerifyFile("sample.pdf");
    }

    #endregion

    [Test]
    public Task VerifyPdfResolution()
    {
        return Verifier.VerifyFile("sample.pdf")
            .PdfSaveAsImageOptions(page => new(){Resolution = 100});
    }

    #region VerifyPdfStream

    [Test]
    public Task VerifyPdfStream()
    {
        return Verifier.Verify(File.OpenRead("sample.pdf"))
            .UseExtension("pdf");
    }

    #endregion

    #region VerifyExcel

    [Test]
    public Task VerifyExcel()
    {
        return Verifier.VerifyFile("sample.xlsx");
    }

    #endregion

    #region VerifyExcelStream

    [Test]
    public Task VerifyExcelStream()
    {
        return Verifier.Verify(File.OpenRead("sample.xlsx"))
            .UseExtension("xlsx");
    }

    #endregion

    #region VerifyWord

    [Test]
    public Task VerifyWord()
    {
        return Verifier.VerifyFile("sample.docx");
    }

    #endregion

    #region VerifyWordStream

    [Test]
    public Task VerifyWordStream()
    {
        return Verifier.Verify(File.OpenRead("sample.docx"))
            .UseExtension("docx");
    }

    #endregion
}