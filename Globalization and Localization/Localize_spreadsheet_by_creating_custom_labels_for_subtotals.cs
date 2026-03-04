using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsLocalizationDemo
{
    // Custom class to set localized subtotal labels
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create an instance of SettableGlobalizationSettings
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Define a custom label for the Subtotal (Sum) function
            // This label will be used when the Subtotal feature inserts a total row
            globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Subtotal");

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // Define the range on which to apply Subtotal (e.g., A1:B5)
            // Adjust the range as needed for your data
            CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0-4, columns 0-1

            // Apply Subtotal: group by column 0 (first column), use Sum on column 1 (second column)
            // Parameters: area, column index to group, function, array of columns to subtotal,
            // replace existing subtotals, page break between groups, summary below data
            workbook.Worksheets[0].Cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}