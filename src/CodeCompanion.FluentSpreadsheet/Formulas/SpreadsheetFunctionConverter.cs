using System;

namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public static class SpreadsheetFunctionConverter
    {
        public static string Convert(SpreadsheetFunction function) => function switch
        {
            SpreadsheetFunction.Sum => "SUM",
            SpreadsheetFunction.Average => "AVERAGE",
            _ => throw new InvalidOperationException($"{nameof(SpreadsheetFunction)}.{function} is unsupported")
        };
    }
}