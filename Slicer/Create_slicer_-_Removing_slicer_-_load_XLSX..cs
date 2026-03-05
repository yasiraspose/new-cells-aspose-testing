using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Assume the first worksheet contains a pivot table
            Worksheet worksheet = workbook.Worksheets[0];
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivot = worksheet.PivotTables[0];

            // Access the slicer collection of the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Add a slicer for the pivot table.
            // Destination cell "E2" is the upper‑left corner of the slicer range.
            // "Fruit" is the name of the base field in the pivot table.
            int slicerIndex = slicers.Add(pivot, "E2", "Fruit");

            // Optional: retrieve the newly added slicer object
            Slicer slicer = slicers[slicerIndex];

            // Remove the slicer we just added using its index
            slicers.RemoveAt(slicerIndex);
            // Alternatively, you could remove by object:
            // slicers.Remove(slicer);

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}