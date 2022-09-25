// See https://aka.ms/new-console-template for more information
using iText.Kernel.Pdf;
using System.IO;
using iText.Layout;
using Microsoft.Office.Interop.Word;//C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429cSS    Add as COM
using iText.Kernel.Utils;

Console.WriteLine("Hello, World!");


var appWord = new Application();
if (appWord.Documents != null)
{
    //yourDoc is your word document
    var path1 = AppDomain.CurrentDomain.BaseDirectory;
    var path2 = Environment.CurrentDirectory + "\\1.docx";
    var wordDocument = appWord.Documents.Open(path2);
    //wordDocument.SelectContentControlsByTitle("FIRLD")[1].Range.Text = "VALIETOBIND";//IF BIND DYNAMIC CONTENT
    string pdfDocName = path1 + "\\pdfDocument.pdf";
    if (wordDocument != null)
    {
        wordDocument.ExportAsFixedFormat(pdfDocName,
        WdExportFormat.wdExportFormatPDF);
        wordDocument.Close();
    }
    appWord.Quit();
    ManipulatePdf(pdfDocName);
}


 void ManipulatePdf(String dest)
{
    var path2 = Environment.CurrentDirectory ;

    PdfDocument pdfDoc = new PdfDocument(new PdfWriter(path2+"\\2.pdf"));
    PdfDocument cover0 = new PdfDocument(new PdfReader(dest));

    PdfDocument cover = new PdfDocument(new PdfReader(path2+ "\\a1.pdf"));
    PdfDocument resource = new PdfDocument(new PdfReader(path2 + "\\a0.pdf"));
    
    PdfMerger merger = new PdfMerger(pdfDoc);
    merger.Merge(cover0, 1, cover0.GetNumberOfPages());
    merger.Merge(cover, 1, cover.GetNumberOfPages());
    merger.Merge(resource, 1, resource.GetNumberOfPages());

    // Source documents can be closed implicitly after merging,
    // by passing true to the setCloseSourceDocuments(boolean) method
    cover.Close();
    resource.Close();

    // The resultant pdf doc will be closed implicitly.
    merger.Close();
}