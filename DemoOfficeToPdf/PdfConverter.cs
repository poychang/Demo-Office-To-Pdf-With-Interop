using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace DemoOfficeToPdf
{
    public class PdfConverter : IDisposable
    {
        bool disposed = false;
        private Excel.Application ExcelApp { get; set; }
        private PowerPoint.Application PowerPointApp { get; set; }
        private Word.Application WordApp { get; set; }

        public PdfConverter()
        {
            ExcelApp = new Excel.Application();
            PowerPointApp = new PowerPoint.Application();
            WordApp = new Word.Application();
        }

        /// <summary>
        /// 將 Excel 轉成 PDF
        /// </summary>
        /// <param name="fileSource">Excel 檔案位置</param>
        /// <param name="saveDestination">PDF 儲存位置</param>
        public void ExcelToPdf(string fileSource, string saveDestination)
        {
            var workbook = ExcelApp.Workbooks.Open(fileSource);
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, saveDestination);
            workbook.Close();
        }

        /// <summary>
        /// 將 PowerPoint 轉成 PDF
        /// </summary>
        /// <param name="fileSource">PowerPoint 檔案位置</param>
        /// <param name="saveDestination">PDF 儲存位置</param>
        public void PowerPointToPdf(string fileSource, string saveDestination)
        {
            var presentations = PowerPointApp.Presentations.Open(fileSource);
            presentations.ExportAsFixedFormat(saveDestination, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
            presentations.Close();
        }

        /// <summary>
        /// 將 Word 轉成 PDF
        /// </summary>
        /// <param name="fileSource">Word 檔案位置</param>
        /// <param name="saveDestination">PDF 儲存位置</param>
        public void WordToPdf(string fileSource, string saveDestination)
        {
            var document = WordApp.Documents.Open(fileSource);
            document.ExportAsFixedFormat(saveDestination, Word.WdExportFormat.wdExportFormatPDF);
            document.Close();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed) return;

            // Free any unmanaged objects here.
            try
            {
                if (disposing)
                {
                    // Free any other managed objects here.
                    ExcelApp.Quit();
                    PowerPointApp.Quit();
                    WordApp.Quit();
                }
                if (ExcelApp?.Sheets != null)
                {
                    Marshal.FinalReleaseComObject(ExcelApp.Sheets);
                }
                if (ExcelApp?.Workbooks != null)
                {
                    ExcelApp.Workbooks.Close();
                    Marshal.FinalReleaseComObject(ExcelApp.Workbooks);
                }
                Marshal.FinalReleaseComObject(ExcelApp);
                Marshal.FinalReleaseComObject(PowerPointApp);
                if (PowerPointApp.Presentations != null)
                {
                    Marshal.FinalReleaseComObject(PowerPointApp.Presentations);
                }
                Marshal.FinalReleaseComObject(WordApp);

                disposed = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                throw;
            }
        }

        ~PdfConverter()
        {
            Dispose(false);
        }
    }
}
