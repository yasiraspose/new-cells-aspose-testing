using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook with a specific culture (e.g., German)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CultureInfo = new CultureInfo("de-DE"); // German uses comma as decimal separator
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Create a SettableGlobalizationSettings instance to customize localization
            SettableGlobalizationSettings locSettings = new SettableGlobalizationSettings();

            // Example: change list separator from comma to semicolon
            locSettings.SetListSeparator(';');

            // Example: customize boolean display strings
            locSettings.SetBooleanValueString(true, "WAHR");   // German for TRUE
            locSettings.SetBooleanValueString(false, "FALSCH"); // German for FALSE

            // Example: map standard function names to German equivalents
            locSettings.SetLocalFunctionName("SUM", "SUMME", true);
            locSettings.SetLocalFunctionName("AVERAGE", "MITTELWERT", true);
            locSettings.SetLocalFunctionName("MAX", "MAXIMUM", true);

            // Example: customize table row type names
            locSettings.SetTableRowTypeOfHeaders("KOPFZEILEN");
            locSettings.SetTableRowTypeOfData("DATEN");
            locSettings.SetTableRowTypeOfTotals("SUMMEN");

            // Apply the localization settings to the workbook
            workbook.Settings.GlobalizationSettings = locSettings;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample numeric data
            for (int i = 1; i <= 5; i++)
            {
                cells[$"B{i}"].PutValue(i * 10); // B1..B5 = 10,20,30,40,50
            }

            // Use localized function names in formulas
            cells["A1"].Formula = "=SUMME(B1:B5)";        // Localized SUM
            cells["A2"].Formula = "=MITTELWERT(B1:B5)"; // Localized AVERAGE
            cells["A3"].Formula = "=MAXIMUM(B1:B5)";    // Localized MAX

            // Use localized boolean values
            cells["C1"].PutValue(true);   // Will display "WAHR"
            cells["C2"].PutValue(false);  // Will display "FALSCH"

            // Calculate all formulas so that results are stored
            workbook.CalculateFormula();

            // Demonstrate retrieving a localized function name via the settings object
            string localizedSum = locSettings.GetLocalFunctionName("SUM");
            Console.WriteLine($"Localized name for 'SUM' is: {localizedSum}");

            // Save the localized workbook
            workbook.Save("localized_output.xlsx");
        }
    }
}