using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Load the workbook with formula parsing enabled
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = true; // ensure formulas are parsed when the file is opened
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Set the workbook region to German to obtain German localized formulas
        workbook.Settings.Region = CountryCode.Germany;

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Iterate through all cells that contain formulas
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];

                // Process only cells with a formula
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    Console.WriteLine($"Cell {cell.Name}:");
                    Console.WriteLine($"  Standard Formula : {cell.Formula}");
                    Console.WriteLine($"  Localized Formula: {cell.FormulaLocal}");
                    // Demonstrate GetFormula with explicit locale flag
                    Console.WriteLine($"  GetFormula (local) : {cell.GetFormula(false, true)}");
                }
            }
        }

        // Save the workbook (optional, e.g., after modifications)
        workbook.Save("output.xlsx");
    }
}