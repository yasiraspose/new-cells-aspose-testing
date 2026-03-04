using System;
using Aspose.Cells;

class RemoveThreadedCommentsDemo
{
    static void Main()
    {
        // Load the workbook from an existing XLSX file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range of the worksheet
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Scan each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    // Retrieve threaded comments for the current cell
                    ThreadedCommentCollection threadedComments = sheet.Comments.GetThreadedComments(row, col);

                    // If threaded comments exist, clear them
                    if (threadedComments != null && threadedComments.Count > 0)
                    {
                        threadedComments.Clear();
                    }
                }
            }
        }

        // Save the workbook after removing all threaded comments
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}