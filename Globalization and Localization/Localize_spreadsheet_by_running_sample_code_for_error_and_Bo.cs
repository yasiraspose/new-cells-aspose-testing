using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings to localize Boolean and error strings
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize Boolean values (e.g., Russian)
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize common Excel error values (e.g., Russian)
        public override string GetErrorValueString(string err)
        {
            switch (err)
            {
                case "#NAME?": return "#ИМЯ?";
                case "#DIV/0!": return "#ДЕЛ/0!";
                case "#REF!": return "#ССЫЛКА!";
                case "#VALUE!": return "#ЗНАЧ!";
                case "#N/A": return "#Н/Д";
                case "#NUM!": return "#ЧИСЛО!";
                case "#NULL!": return "#ПУСТО!";
                default: return base.GetErrorValueString(err);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (can be an empty workbook or existing file)
            string inputPath = "input.xlsx";

            // Load the workbook using the standard constructor (create/load rule)
            Workbook wb = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate sample Boolean values
            cells[0, 0].PutValue(true);   // A1
            cells[0, 1].PutValue(false);  // B1

            // Populate sample error values
            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };
            for (int i = 0; i < errors.Length; i++)
            {
                // Starting from column C (index 2)
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Apply the custom globalization settings to the workbook
            wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Display localized string values in the console
            for (int col = 0; col < 9; col++)
            {
                Console.WriteLine($"Cell[0,{col}] (localized): {cells[0, col].StringValue}");
            }

            // Save the localized workbook (save rule)
            string outputPath = "localized_output.xlsx";
            wb.Save(outputPath);
        }
    }
}