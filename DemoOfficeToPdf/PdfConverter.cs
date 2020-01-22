using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace DemoOfficeToPdf
{
    /// <summary>來源檔案類型</summary>
    public enum SourceType
    {
        Excel,
        PowerPoint,
        Word
    }

    /// <summary>轉存相關資訊</summary>
    public class SaveParam
    {
        /// <summary>檔案來源位置</summary>
        public string Source { get; set; }
        /// <summary>儲存目的位置</summary>
        public string Destination { get; set; }
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
        public void SaveAsPdf(SourceType type, string source, string destination)
        {
            switch (type)
            {
                case SourceType.Excel:
                    ExcelAsPdf(source, destination);
                    break;
                case SourceType.PowerPoint:
                    PowerPointAsPdf(source, destination);
                    break;
                case SourceType.Word:
                    WordAsPdf(source, destination);
                    break;
                default:
                    break;
            }
        }
        public void SaveAsPdf(SourceType type, SaveParam param) => SaveAsPdf(type, param.Source, param.Destination);

        /// <summary>
        /// 將 Excel 轉成 PDF
        /// </summary>
        /// <param name="source">檔案來源位置</param>
        /// <param name="destination">儲存目的位置</param>
        public void ExcelAsPdf(string source, string destination)
        {
            var excelApp = new Excel.Application();
            var workbooks = excelApp.Workbooks;
            var workbook = workbooks.Open(source);
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, destination);
            workbook.Close();
            excelApp.Quit();
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(workbooks);
            Marshal.FinalReleaseComObject(excelApp);
        }
        public void ExcelAsPdf(SaveParam param) => ExcelAsPdf(param.Source, param.Destination);

        /// <summary>
        /// 將 PowerPoint 轉成 PDF
        /// </summary>
        /// <param name="source">檔案來源位置</param>
        /// <param name="destination">儲存目的位置</param>
        public void PowerPointAsPdf(string source, string destination)
        {
            var powerPointApp = new PowerPoint.Application();
            var presentations = powerPointApp.Presentations;
            var presentation = presentations.Open(source);
            presentation.ExportAsFixedFormat(destination, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
            presentation.Close();
            powerPointApp.Quit();
            Marshal.FinalReleaseComObject(presentation);
            Marshal.FinalReleaseComObject(presentations);
            Marshal.FinalReleaseComObject(powerPointApp);
        }
        public void PowerPointAsPdf(SaveParam param) => PowerPointAsPdf(param.Source, param.Destination);

        /// <summary>
        /// 將 Word 轉成 PDF
        /// </summary>
        /// <param name="source">檔案來源位置</param>
        /// <param name="destination">儲存目的位置</param>
        public void WordAsPdf(string source, string destination)
        {
            var wordApp = new Word.Application();
            var documents = wordApp.Documents;
            var document = documents.Open(source);
            document.ExportAsFixedFormat(destination, Word.WdExportFormat.wdExportFormatPDF);
            document.Close();
            wordApp.Quit();
            Marshal.FinalReleaseComObject(document);
            Marshal.FinalReleaseComObject(documents);
            Marshal.FinalReleaseComObject(wordApp);
        }
        public void WordAsPdf(SaveParam param) => WordAsPdf(param.Source, param.Destination);

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
