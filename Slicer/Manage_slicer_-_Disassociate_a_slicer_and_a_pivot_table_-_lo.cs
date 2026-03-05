using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerDisassociateDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (XLSX)
            Workbook workbook = new Workbook("input.xlsx");

            // Assume the slicer and its related pivot table are on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the slicer collection from the worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Ensure there is at least one slicer
            if (slicers.Count > 0)
            {
                // Get the first slicer (adjust index as needed)
                Slicer slicer = slicers[0];

                // Get the pivot table to which the slicer is currently connected.
                // Here we simply take the first pivot table on the same sheet.
                // In a real scenario, you might locate the specific pivot table by name.
                if (sheet.PivotTables.Count > 0)
                {
                    PivotTable pivot = sheet.PivotTables[0];

                    // Disassociate the slicer from the pivot table
                    slicer.RemovePivotConnection(pivot);
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}