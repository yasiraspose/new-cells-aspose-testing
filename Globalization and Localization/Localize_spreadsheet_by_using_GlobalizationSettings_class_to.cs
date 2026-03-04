using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsGlobalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // ------------------------------------------------------------
            // 1. Create and configure SettableGlobalizationSettings
            // ------------------------------------------------------------
            SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();

            // Customize the total label for the SUM function (used in subtotals)
            globalSettings.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

            // ------------------------------------------------------------
            // 2. Create and configure SettableChartGlobalizationSettings
            // ------------------------------------------------------------
            SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();

            // Customize various chart texts
            chartSettings.SetSeriesName("Custom Series");
            chartSettings.SetChartTitleName("Custom Pie Chart");
            chartSettings.SetLegendTotalName("Custom Legend Total");
            chartSettings.SetOtherName("Other Category");

            // Attach the chart globalization settings to the workbook settings
            globalSettings.ChartSettings = chartSettings;

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalSettings;

            // ------------------------------------------------------------
            // 3. Demonstrate subtotal with the customized total name
            // ------------------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            Cells cells = dataSheet.Cells;

            // Define the range for subtotal (assumes data in A1:B5)
            CellArea area = CellArea.CreateCellArea(0, 0, 4, 1);
            // Subtotal on the first column (index 0) using SUM, grouping by column 0
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // ------------------------------------------------------------
            // 4. Locate a pie chart and verify the custom labels
            // ------------------------------------------------------------
            // Assume the first chart on the first worksheet is a pie chart
            if (dataSheet.Charts.Count > 0)
            {
                Chart chart = dataSheet.Charts[0];

                // Ensure the chart type is Pie; if not, change it for demonstration
                if (chart.Type != ChartType.Pie)
                {
                    chart.Type = ChartType.Pie;
                }

                // The chart title will reflect the custom title set via globalization settings
                // (no need to set it manually; the globalization settings affect it automatically)

                // Optionally, set category and series data if the chart is empty
                if (chart.NSeries.Count == 0)
                {
                    // Example data range (adjust as needed)
                    chart.NSeries.Add("B2:B5", true);
                    chart.NSeries.CategoryData = "A2:A5";
                }
            }

            // ------------------------------------------------------------
            // 5. Save the modified workbook
            // ------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}