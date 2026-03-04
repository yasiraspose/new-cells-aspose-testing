using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class ExcelToPdfConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string destPath = "output.pdf";

            ConversionUtility.Convert(sourcePath, destPath);
            Console.WriteLine("Excel file has been successfully converted to PDF.");
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}