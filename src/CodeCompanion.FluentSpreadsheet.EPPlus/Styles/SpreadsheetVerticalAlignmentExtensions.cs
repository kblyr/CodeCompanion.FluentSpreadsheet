using System;
using OfficeOpenXml.Style;

namespace CodeCompanion.FluentSpreadsheet.Styles
{
    public static class SpreadsheetVerticalAlignmentExtensions
    {
        public static ExcelVerticalAlignment ToRaw(this SpreadsheetVerticalAlignment alignment) => alignment switch
        {
            SpreadsheetVerticalAlignment.Top => ExcelVerticalAlignment.Top,
            SpreadsheetVerticalAlignment.Center => ExcelVerticalAlignment.Center,
            SpreadsheetVerticalAlignment.Bottom => ExcelVerticalAlignment.Bottom,
            _ => throw new InvalidOperationException($"{nameof(SpreadsheetVerticalAlignment)}.{alignment} is unsupported")
        };
    }
}