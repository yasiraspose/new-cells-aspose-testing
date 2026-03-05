using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Desired output PDF file path
            string destPath = "output.pdf";

            // Convert the Excel workbook to PDF using the provided ConversionUtility method
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine("Excel file has been successfully converted to PDF.");
        }
    }
}