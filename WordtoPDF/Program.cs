// See https://aka.ms/new-console-template for more information
using iText.Kernel.Pdf;
using System.IO;
using iText.Layout;
using Microsoft.Office.Interop.Word;//C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429cSS    Add as COM
using iText.Kernel.Utils;
using iText.Kernel.Geom;

Console.WriteLine("Hello, World!");


var appWord = new Application();
if (appWord.Documents != null)
{
    //yourDoc is your word document
    //var path1 = AppDomain.CurrentDomain.BaseDirectory;
    var path2 = Environment.CurrentDirectory ;
    var wordDocument = appWord.Documents.Open(path2 + "\\1.doc");
    //wordDocument.SelectContentControlsByTitle("FIRLD")[1].Range.Text = "VALIETOBIND";//IF BIND DYNAMIC CONTENT
    string pdfDocName = path2 + "\\pdfDocument.pdf";
    if (wordDocument != null)
    {
        wordDocument.ExportAsFixedFormat(pdfDocName,
        WdExportFormat.wdExportFormatPDF);
        wordDocument.Close();
    }
    appWord.Quit();
    List<string> pathsObj = new List<string>();
    pathsObj.Add(pdfDocName);
    pathsObj.Add(path2 + "\\a1.pdf");
    pathsObj.Add(path2 + "\\a0.pdf");
    ManipulatePdf(pathsObj);
}


 void ManipulatePdf(List<string> pathsObj)
{
    var path2 = Environment.CurrentDirectory ;
    PdfDocument pdfDoc = new PdfDocument(new PdfWriter(path2+"\\2.pdf"));//ceates new file
    PdfMerger merger = new PdfMerger(pdfDoc);
    foreach (var path in pathsObj)
    {
        PdfDocument cover = new PdfDocument(new PdfReader(path));
        merger.Merge(cover, 1, cover.GetNumberOfPages());
        // Source documents can be closed implicitly after merging,
        // by passing true to the setCloseSourceDocuments(boolean) method
        cover.Close();
    }
    // The resultant pdf doc will be closed implicitly.
    merger.Close();
}