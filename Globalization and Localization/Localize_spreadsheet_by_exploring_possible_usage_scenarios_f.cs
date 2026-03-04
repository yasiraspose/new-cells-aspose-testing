using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings that localize boolean and error strings
    public class CustomGlobalizationSettings : SettableGlobalizationSettings
    {
        // Override to provide localized boolean display strings
        public override string GetBooleanValueString(bool value)
        {
            // Example: Russian localization
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Override to provide localized error display strings
        public override string GetErrorValueString(string error)
        {
            // Map standard English error texts to Russian equivalents
            return error switch
            {
                "#NAME?" => "#ИМЯ?",
                "#DIV/0!" => "#ДЕЛ/0!",
                "#REF!" => "#ССЫЛКА!",
                "#VALUE!" => "#ЗНАЧ!",
                "#N/A" => "#Н/Д",
                "#NUM!" => "#ЧИСЛО!",
                "#NULL!" => "#ПУСТО!",
                _ => base.GetErrorValueString(error)
            };
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (replace with actual path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for demonstration
            cells[0, 0].PutValue(true);   // Boolean true
            cells[0, 1].PutValue(false);  // Boolean false

            // Insert various error values
            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };
            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Display the localized string representations in the console
            Console.WriteLine("Localized cell values:");
            for (int col = 0; col < 9; col++)
            {
                // StringValue returns the displayed text according to globalization settings
                Console.WriteLine($"Cell[0,{col}]: {cells[0, col].StringValue}");
            }

            // Save the workbook to a new file to preserve the settings
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}