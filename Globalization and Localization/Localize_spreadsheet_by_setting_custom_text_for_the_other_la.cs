using System;
using Aspose.Cells;

class LocalizePieChartOtherLabel
{
    static void Main()
    {
        // Path to the existing XLSX workbook
        string inputPath = "input.xlsx";

        // Load the workbook (XLSX format)
        Workbook workbook = new Workbook(inputPath);

        // Create chart globalization settings and set custom text for the "Other" label
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetOtherName("Miscellaneous Items");

        // Create overall globalization settings and assign the chart settings
        SettableGlobalizationSettings globalizationSettings = new SettableGlobalizationSettings
        {
            ChartSettings = chartSettings
        };

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalizationSettings;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}