using System;
using System.IO;
using System.Linq;

namespace DemoOfficeToPdf
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            using (var pdfConverter = new PdfConverter())
            {
                var excelSource = Path.Combine(currentDirectory, "Files/workbook.xlsx");
                var excelPdfDestination = excelSource.Replace(excelSource.Split('.').Last(), "pdf");
                pdfConverter.SaveAsPdf(SourceType.Excel, excelSource, excelPdfDestination);
                //pdfConverter.ExcelAsPdf(excelSource, excelPdfDestination);

                var powerpointSource = Path.Combine(currentDirectory, "Files/presentation.pptx");
                var powerpointPdfDestination = powerpointSource.Replace(powerpointSource.Split('.').Last(), "pdf");
                pdfConverter.SaveAsPdf(SourceType.PowerPoint, powerpointSource, powerpointPdfDestination);
                //pdfConverter.PowerPointAsPdf(powerpointSource, powerpointPdfDestination);

                var wordSource = Path.Combine(currentDirectory, "Files/document.docx");
                var wordPdfDestination = wordSource.Replace(wordSource.Split('.').Last(), "pdf");
                pdfConverter.SaveAsPdf(SourceType.Word, wordSource, wordPdfDestination);
                //pdfConverter.WordAsPdf(wordSource, wordPdfDestination);
            }

            Console.WriteLine("Done!");
        }
    }
}
