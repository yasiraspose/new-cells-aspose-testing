using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing XLSX workbook
            string inputPath = "input.xlsx";

            // Load the workbook (create rule)
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook locale to German (for demonstration)
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and cell A1
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Ensure the cell has a formula in standard (English) notation
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized forms
            Console.WriteLine("Standard Formula : " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Set the formula using the localized (German) notation
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Display the formulas again after setting the localized version
            Console.WriteLine("\nAfter setting FormulaLocal:");
            Console.WriteLine("Standard Formula : " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Optionally calculate the workbook to update the result
            workbook.CalculateFormula();

            // Show the calculated value
            Console.WriteLine("\nCalculated Value: " + cell.Value);

            // Save the modified workbook (save rule)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}