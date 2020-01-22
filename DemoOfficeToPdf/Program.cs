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
                pdfConverter.SaveToPdf(SourceType.Excel, excelSource, excelPdfDestination);
                //pdfConverter.ExcelToPdf(excelSource, excelPdfDestination);

                var powerpointSource = Path.Combine(currentDirectory, "Files/presentation.pptx");
                var powerpointPdfDestination = powerpointSource.Replace(powerpointSource.Split('.').Last(), "pdf");
                pdfConverter.SaveToPdf(SourceType.PowerPoint, powerpointSource, powerpointPdfDestination);
                //pdfConverter.PowerPointToPdf(powerpointSource, powerpointPdfDestination);

                var wordSource = Path.Combine(currentDirectory, "Files/document.docx");
                var wordPdfDestination = wordSource.Replace(wordSource.Split('.').Last(), "pdf");
                pdfConverter.SaveToPdf(SourceType.Word, wordSource, wordPdfDestination);
                //pdfConverter.WordToPdf(wordSource, wordPdfDestination);
            }

            Console.WriteLine("Done!");
        }
    }
}
