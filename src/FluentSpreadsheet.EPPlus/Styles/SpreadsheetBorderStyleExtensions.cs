using System;
using OfficeOpenXml.Style;

namespace CodeCompanion.FluentSpreadsheet.Styles
{
    public static class SpreadsheetBorderStyleExtensions
    {
        public static ExcelBorderStyle ToRaw(this SpreadsheetBorderStyle style) => style switch
        {
            SpreadsheetBorderStyle.None => ExcelBorderStyle.None,
            SpreadsheetBorderStyle.Thin => ExcelBorderStyle.Thin,
            SpreadsheetBorderStyle.Thick => ExcelBorderStyle.Thick,
            SpreadsheetBorderStyle.Dotted => ExcelBorderStyle.Dotted,
            SpreadsheetBorderStyle.Dashed => ExcelBorderStyle.Dashed,
            SpreadsheetBorderStyle.Double => ExcelBorderStyle.Double,
            _ => throw new InvalidOperationException($"{nameof(SpreadsheetBorderStyle)}.{style} is unsupported")
        };
    }
}