using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Desired output XPS file path
            string destPath = "output.xps";

            // Load the workbook and save it as XPS
            var workbook = new Workbook(sourcePath);
            workbook.Save(destPath, SaveFormat.Xps);

            Console.WriteLine($"Conversion completed: '{sourcePath}' → '{destPath}'");
        }
    }
}