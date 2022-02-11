using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;


namespace PDFWriter
{
    class Program
    {
        static void Main(string[] args)
        {
              
            //Microsoft.Office.Interop.Word.Document wordDocument;

            var appWord = new Application();
            if (appWord.Documents != null)
            {
                //yourDoc is your word document
                var path1 = AppDomain.CurrentDomain.BaseDirectory;
                var path2 = Environment.CurrentDirectory+ "\\1.docx";
                var wordDocument = appWord.Documents.Open(path2);
                string pdfDocName = path1+"\\pdfDocument.pdf";
                if (wordDocument != null)
                {
                    wordDocument.ExportAsFixedFormat(pdfDocName,
                    WdExportFormat.wdExportFormatPDF);
                    wordDocument.Close();
                }
                appWord.Quit();
            }
            //Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            //wordDocument = appWord.Documents.Open(@"\\1.docx");
            //wordDocument.ExportAsFixedFormat(@"\\DocTo.pdf", WdExportFormat.wdExportFormatPDF);
        }
    }
}
