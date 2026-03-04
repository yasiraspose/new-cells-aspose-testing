using System;
using Aspose.Cells;

namespace RemoveThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Determine the used range of the worksheet
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxCol = worksheet.Cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Retrieve the threaded comments collection for the current cell
                        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(row, col);

                        // If there are any threaded comments, remove them all
                        if (threadedComments != null && threadedComments.Count > 0)
                        {
                            // Remove comments starting from the last index to avoid shifting issues
                            for (int i = threadedComments.Count - 1; i >= 0; i--)
                            {
                                threadedComments.RemoveAt(i);
                            }
                        }
                    }
                }
            }

            // Save the workbook after removing all threaded comments
            workbook.Save("output.xlsx");
        }
    }
}