using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings to localize subtotal and grand total labels
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override subtotal label text
        public override string GetSubTotalName(PivotFieldSubtotalType subTotalType)
        {
            // Example: prepend "Loc-" to the default subtotal type name
            return "Loc-" + subTotalType.ToString();
        }

        // Override grand total label text for a specific consolidation function
        public override string GetGrandTotalName(ConsolidationFunction functionType)
        {
            // Example: provide a custom label for Sum grand total
            if (functionType == ConsolidationFunction.Sum)
                return "Loc-Total Sum";
            // Fallback to default implementation for other functions
            return base.GetGrandTotalName(functionType);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Apply custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Define the range for which subtotals will be calculated (e.g., A1:B10)
            CellArea area = CellArea.CreateCellArea(0, 0, 9, 1); // rows 0-9, columns 0-1

            // Apply subtotal:
            //   - group by column 0 (first column)
            //   - use Sum function
            //   - include subtotals, grand totals, and replace existing data
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, true, true);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}