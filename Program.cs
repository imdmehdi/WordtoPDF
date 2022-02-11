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
                var wordDocument = appWord.Documents.Open("\\1.docx");
                string pdfDocName = "\\pdfDocument.pdf";
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
