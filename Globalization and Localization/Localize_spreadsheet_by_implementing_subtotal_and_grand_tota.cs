using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create a SettableGlobalizationSettings instance to customize built‑in texts
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

        // Localize the "Total" label for the SUM function (e.g., Spanish)
        globalization.SetTotalName(ConsolidationFunction.Sum, "Suma Total");

        // Localize the "Grand Total" label for the SUM function
        globalization.SetGrandTotalName(ConsolidationFunction.Sum, "Suma Total General");

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // OPTIONAL: Demonstrate the effect by applying a Subtotal operation
        // (Assumes data exists in the first worksheet from A1 to B5)
        CellArea dataRange = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0‑4, columns 0‑1
        workbook.Worksheets[0].Cells.Subtotal(
            dataRange,               // range to subtotal
            0,                       // column index to group by (A)
            ConsolidationFunction.Sum, // function for subtotal
            new int[] { 0 },         // columns to subtotal (A)
            true,                    // replace data
            false,                   // page break after each group
            true);                   // include grand total

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}