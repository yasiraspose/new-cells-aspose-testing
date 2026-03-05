using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class ExcelToPdfConverter
    {
        /// <summary>
        /// Converts an XLSX workbook to PDF.
        /// </summary>
        /// <param name="inputPath">Full path of the source XLSX file.</param>
        /// <param name="outputPath">Full path where the PDF will be saved.</param>
        public static void ConvertXlsxToPdf(string inputPath, string outputPath)
        {
            // Load the existing Excel workbook from the specified file.
            // This uses the Workbook(string) constructor rule.
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as PDF.
            // This uses the Workbook.Save(string, SaveFormat) rule.
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Conversion completed: \"{inputPath}\" -> \"{outputPath}\"");
        }

        // Example usage
        public static void Main()
        {
            // Define source XLSX and destination PDF file paths.
            string sourceFile = "source.xlsx";
            string destFile = "output.pdf";

            // Perform the conversion.
            ConvertXlsxToPdf(sourceFile, destFile);
        }
    }
}