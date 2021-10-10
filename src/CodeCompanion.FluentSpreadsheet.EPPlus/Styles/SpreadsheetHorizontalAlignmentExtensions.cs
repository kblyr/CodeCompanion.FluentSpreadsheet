using System;
using OfficeOpenXml.Style;

namespace CodeCompanion.FluentSpreadsheet.Styles
{
    public static class SpreadsheetHorizontalAlignmentExtensions
    {
        public static ExcelHorizontalAlignment ToRaw(this SpreadsheetHorizontalAlignment alignment) => alignment switch
        {
            SpreadsheetHorizontalAlignment.Left => ExcelHorizontalAlignment.Left,
            SpreadsheetHorizontalAlignment.Center => ExcelHorizontalAlignment.Center,
            SpreadsheetHorizontalAlignment.Right => ExcelHorizontalAlignment.Right,
            _ => throw new InvalidOperationException($"{nameof(SpreadsheetHorizontalAlignment)}.{alignment} is unsupported")
        };
    }
}