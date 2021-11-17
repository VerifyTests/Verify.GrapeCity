# <img src="/src/icon.png" height="30px"> Verify.GrapeCity

[![Build status](https://ci.appveyor.com/api/projects/status/7k8hh0guut2ioak2?svg=true)](https://ci.appveyor.com/project/SimonCropp/Verify-GrapeCity)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.GrapeCity.svg)](https://www.nuget.org/packages/Verify.GrapeCity/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of documents via [GrapeCity Documents](https://www.grapecity.com/#dx).

Converts documents (pdf, docx, and xslx) to png for verification.

An [GrapeCity License](https://www.grapecity.com/documents-api/pricing) is required to use this tool.

<a href='https://dotnetfoundation.org' alt='Part of the .NET Foundation'><img src='https://raw.githubusercontent.com/VerifyTests/Verify/master/docs/dotNetFoundation.svg' height='30px'></a><br>
Part of the <a href='https://dotnetfoundation.org' alt=''>.NET Foundation</a>


## NuGet package

https://nuget.org/packages/Verify.GrapeCity/


## Usage


### Enable Verify.GrapeCity

<!-- snippet: ModuleInitializer.cs -->
<a id='snippet-ModuleInitializer.cs'></a>
```cs
using VerifyTests;

public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyGrapeCity.Initialize();
    }
}
```
<sup><a href='/src/Tests/ModuleInitializer.cs#L1-L10' title='Snippet source file'>snippet source</a> | <a href='#snippet-ModuleInitializer.cs' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### PDF


#### Verify a file

<!-- snippet: VerifyPdf -->
<a id='snippet-verifypdf'></a>
```cs
[Test]
public Task VerifyPdf()
{
    return Verifier.VerifyFile("sample.pdf");
}
```
<sup><a href='/src/Tests/Samples.cs#L8-L16' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifypdf' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyPdfStream -->
<a id='snippet-verifypdfstream'></a>
```cs
[Test]
public Task VerifyPdfStream()
{
    return Verifier.Verify(File.OpenRead("sample.pdf"))
        .UseExtension("pdf");
}
```
<sup><a href='/src/Tests/Samples.cs#L25-L34' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifypdfstream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyPdf.00.verified.txt -->
<a id='snippet-Samples.VerifyPdf.00.verified.txt'></a>
```txt
{
  Title: ,
  Author: ,
  CreationDate: DateTimeOffset_1,
  ModifyDate: DateTimeOffset_2,
  Subject: ,
  Keywords: ,
  Creator: RAD PDF,
  Producer: RAD PDF 3.9.0.0 - http://www.radpdf.com
}
```
<sup><a href='/src/Tests/Samples.VerifyPdf.00.verified.txt#L1-L10' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyPdf.00.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyPdf.01.verified.png](/src/Tests/Samples.VerifyPdf.01.verified.png):

<img src="/src/Tests/Samples.VerifyPdf.01.verified.png" width="200px">


### Excel


#### Verify a file

<!-- snippet: VerifyExcel -->
<a id='snippet-verifyexcel'></a>
```cs
[Test]
public Task VerifyExcel()
{
    return Verifier.VerifyFile("sample.xlsx");
}
```
<sup><a href='/src/Tests/Samples.cs#L36-L44' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifyexcel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyExcelStream -->
<a id='snippet-verifyexcelstream'></a>
```cs
[Test]
public Task VerifyExcelStream()
{
    return Verifier.Verify(File.OpenRead("sample.xlsx"))
        .UseExtension("xlsx");
}
```
<sup><a href='/src/Tests/Samples.cs#L46-L55' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifyexcelstream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyExcel.00.verified.txt -->
<a id='snippet-Samples.VerifyExcel.00.verified.txt'></a>
```txt
{
  BuiltInDocumentProperties: {},
  CustomDocumentProperties: {},
  DefaultTableStyle: TableStyleMedium2,
  EnableCalculation: true
}
```
<sup><a href='/src/Tests/Samples.VerifyExcel.00.verified.txt#L1-L6' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyExcel.00.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyExcel.01.verified.png](/src/Tests/Samples.VerifyExcel.01.verified.png):

<img src="/src/Tests/Samples.VerifyExcel.01.verified.png" width="200px">


### Word


#### Verify a file

<!-- snippet: VerifyWord -->
<a id='snippet-verifyword'></a>
```cs
[Test]
public Task VerifyWord()
{
    return Verifier.VerifyFile("sample.docx");
}
```
<sup><a href='/src/Tests/Samples.cs#L57-L65' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifyword' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyWordStream -->
<a id='snippet-verifywordstream'></a>
```cs
[Test]
public Task VerifyWordStream()
{
    return Verifier.Verify(File.OpenRead("sample.docx"))
        .UseExtension("docx");
}
```
<sup><a href='/src/Tests/Samples.cs#L67-L76' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifywordstream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyWord.00.verified.txt -->
<a id='snippet-Samples.VerifyWord.00.verified.txt'></a>
```txt
{
  Pages: 1,
  ApplicationName: Microsoft Office Word,
  AplicationVersion: 16.0,
  Company: ,
  CreatedTime: DateTime_1,
  Language: en-US,
  Security: None,
  RevisionNumber: 3,
  DefaultFont: Liberation Serif
}
```
<sup><a href='/src/Tests/Samples.VerifyWord.00.verified.txt#L1-L11' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyWord.00.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyWord.01.verified.png](/src/Tests/Samples.VerifyWord.01.verified.png):

<img src="/src/Tests/Samples.VerifyWord.01.verified.png" width="200px">



## File Samples

http://file-examples.com/



## Icon

[Swirl](https://thenounproject.com/term/swirl/1568686/) designed by [creativepriyanka](https://thenounproject.com/creativepriyanka) from [The Noun Project](https://thenounproject.com/).
