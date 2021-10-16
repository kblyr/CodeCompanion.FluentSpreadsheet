using System;
using OfficeOpenXml.Style;

namespace CodeCompanion.FluentSpreadsheet.Styles
{
    public static class SpreadsheetUnderlineStyleExtensions
    {
        public static ExcelUnderLineType ToRaw(this SpreadsheetUnderlineStyle style) => style switch
        {
            SpreadsheetUnderlineStyle.None => ExcelUnderLineType.None,
            SpreadsheetUnderlineStyle.Single => ExcelUnderLineType.Single,
            SpreadsheetUnderlineStyle.Double => ExcelUnderLineType.Double,
            _ => throw new InvalidOperationException($"{nameof(SpreadsheetUnderlineStyle)}.{style} is unsupported")
        };
    }
}