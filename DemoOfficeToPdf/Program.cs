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
                var excelFilepath = Path.Combine(currentDirectory, "Files/workbook.xlsx");
                pdfConverter.ExcelToPdf(excelFilepath, excelFilepath.Replace(excelFilepath.Split('.').Last(), "pdf"));

                var powerpointFilepath = Path.Combine(currentDirectory, "Files/presentation.pptx");
                pdfConverter.PowerPointToPdf(powerpointFilepath, powerpointFilepath.Replace(powerpointFilepath.Split('.').Last(), "pdf"));

                var wordFilepath = Path.Combine(currentDirectory, "Files/document.docx");
                pdfConverter.WordToPdf(wordFilepath, wordFilepath.Replace(wordFilepath.Split('.').Last(), "pdf"));
            }

            Console.WriteLine("Done!");
        }
    }
}
