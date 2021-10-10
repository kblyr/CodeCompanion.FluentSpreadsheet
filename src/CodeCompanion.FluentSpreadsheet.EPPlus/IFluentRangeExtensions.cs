using CodeCompanion.FluentSpreadsheet.Styles;
using OfficeOpenXml.Style;

namespace CodeCompanion.FluentSpreadsheet
{
    public static class IFluentRangeExtensions
    {
        public static IFluentRange WithBorder(this IFluentRange range, ExcelBorderItem item, SpreadsheetBorder border)
        {
            if (border.Color is not null)
                item.Color.SetColor(border.Color.Value.ToDotNetColor());

            if (border.Style is not null)
                item.Style = border.Style.Value.ToRaw();

            return range;
        }

        public static IFluentRange WithBorderIf(this IFluentRange range, ExcelBorderItem item, SpreadsheetBorder border, bool condition)
        {
            if (condition)
                return range.WithBorder(item, border);

            return range;
        }
    }
}