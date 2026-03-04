using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Apply custom Russian globalization settings
        workbook.Settings.GlobalizationSettings = new RussianGlobalizationSettings();

        // Recalculate formulas so that localized boolean and error strings are reflected
        workbook.CalculateFormula();

        // Save the localized workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }

    // Custom globalization settings for Russian language
    class RussianGlobalizationSettings : GlobalizationSettings
    {
        // Localize boolean values
        public override string GetBooleanValueString(bool value)
        {
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize error values
        public override string GetErrorValueString(string error)
        {
            switch (error)
            {
                case "#NAME?":   return "#ИМЯ?";
                case "#DIV/0!": return "#ДЕЛ/0!";
                case "#REF!":   return "#ССЫЛКА!";
                case "#VALUE!": return "#ЗНАЧ!";
                case "#N/A":    return "#Н/Д";
                case "#NUM!":   return "#ЧИСЛО!";
                case "#NULL!":  return "#ПУСТО!";
                default:        return base.GetErrorValueString(error);
            }
        }
    }
}