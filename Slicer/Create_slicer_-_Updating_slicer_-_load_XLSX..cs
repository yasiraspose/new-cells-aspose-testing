using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Assume the first worksheet contains the data and (optionally) a pivot table
        Worksheet dataSheet = workbook.Worksheets[0];

        // Try to get an existing pivot table; if none exists, create a simple one
        PivotTable pivotTable;
        if (dataSheet.PivotTables.Count > 0)
        {
            pivotTable = dataSheet.PivotTables[0];
        }
        else
        {
            // Create a basic pivot table using a sample range (adjust as needed)
            int pivotIdx = dataSheet.PivotTables.Add("A1:B10", "D1", "PivotTable1");
            pivotTable = dataSheet.PivotTables[pivotIdx];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // First column as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Second column as data field
        }

        // Add a new worksheet that will host the slicer
        Worksheet slicerSheet = workbook.Worksheets.Add("SlicerSheet");

        // Add a slicer linked to the pivot table.
        // Using overload Add(PivotTable, string destCellName, int baseFieldIndex)
        int slicerIndex = slicerSheet.Slicers.Add(pivotTable, "A1", 0);
        Slicer slicer = slicerSheet.Slicers[slicerIndex];

        // Modify underlying data to demonstrate that the slicer needs to be refreshed
        dataSheet.Cells["A2"].PutValue("UpdatedValue");

        // Refresh the slicer (also refreshes the associated pivot table)
        slicer.Refresh();

        // Save the workbook with the new slicer
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}