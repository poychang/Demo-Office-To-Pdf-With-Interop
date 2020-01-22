using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace DemoOfficeToPdf
{
    public enum SourceType
    {
        Excel,
        PowerPoint,
        Word
    }

    public class PdfConverter : IDisposable
    {
        private bool disposed = false;

        /// <summary>
        ///  轉成 PDF
        /// </summary>
        /// <param name="type">來源檔案類型</param>
        /// <param name="source">檔案來源位置</param>
        /// <param name="destination">儲存目的位置</param>
        public void SaveToPdf(SourceType type, string source, string destination)
        {
            switch (type)
            {
                case SourceType.Excel:
                    ExcelToPdf(source, destination);
                    break;
                case SourceType.PowerPoint:
                    PowerPointToPdf(source, destination);
                    break;
                case SourceType.Word:
                    WordToPdf(source, destination);
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 將 Excel 轉成 PDF
        /// </summary>
        /// <param name="source">檔案來源位置</param>
        /// <param name="destination">儲存目的位置</param>
        public void ExcelToPdf(string source, string destination)
        {
            var excelApp = new Excel.Application();
            var workbooks = excelApp.Workbooks;
            var workbook = workbooks.Open(source);
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, destination);
            workbook.Close();
            excelApp.Quit();
            GC.Collect();
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(workbooks);
            Marshal.FinalReleaseComObject(excelApp);
        }

        /// <summary>
        /// 將 PowerPoint 轉成 PDF
        /// </summary>
        /// <param name="source">檔案來源位置</param>
        /// <param name="destination">儲存目的位置</param>
        public void PowerPointToPdf(string source, string destination)
        {
            var powerPointApp = new PowerPoint.Application();
            var presentations = powerPointApp.Presentations;
            var presentation = presentations.Open(source);
            presentation.ExportAsFixedFormat(destination, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
            presentation.Close();
            powerPointApp.Quit();
            GC.Collect();
            Marshal.FinalReleaseComObject(presentation);
            Marshal.FinalReleaseComObject(presentations);
            Marshal.FinalReleaseComObject(powerPointApp);
        }

        /// <summary>
        /// 將 Word 轉成 PDF
        /// </summary>
        /// <param name="source">檔案來源位置</param>
        /// <param name="destination">儲存目的位置</param>
        public void WordToPdf(string source, string destination)
        {
            var wordApp = new Word.Application();
            var documents = wordApp.Documents;
            var document = documents.Open(source);
            document.ExportAsFixedFormat(destination, Word.WdExportFormat.wdExportFormatPDF);
            document.Close();
            wordApp.Quit();
            GC.Collect();
            Marshal.FinalReleaseComObject(document);
            Marshal.FinalReleaseComObject(documents);
            Marshal.FinalReleaseComObject(wordApp);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed) return;

            if (disposing)
            {
                // Free any other managed objects here.
                //
            }

            // Free any unmanaged objects here.
            //
            disposed = true;
        }

        ~PdfConverter()
        {
            Dispose(false);
        }
    }
}
