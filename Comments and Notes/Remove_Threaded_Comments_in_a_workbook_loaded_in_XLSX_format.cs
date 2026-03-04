using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentRemoval
{
    class Program
    {
        static void Main()
        {
            // Load the workbook from an existing XLSX file
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Determine the used range of the worksheet
                Cells cells = worksheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Loop through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        // Retrieve the threaded comments collection for the current cell
                        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(row, col);

                        // If there are any threaded comments, clear them
                        if (threadedComments != null && threadedComments.Count > 0)
                        {
                            threadedComments.Clear();
                        }
                    }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}