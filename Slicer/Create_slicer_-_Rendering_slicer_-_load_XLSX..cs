using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing workbook that contains a pivot table
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one pivot table
        if (worksheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Retrieve the first pivot table
        PivotTable pivotTable = worksheet.PivotTables[0];

        // Determine the name of the first base field to use for the slicer
        string baseFieldName = pivotTable.BaseFields[0].Name;

        // Add a slicer anchored at cell D1 for the selected base field
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "D1", baseFieldName);
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Optional: customize slicer appearance
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer.Caption = "Sample Slicer";

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}