using System;
namespace CodeCompanion.FluentSpreadsheet
{
    public static class IFluentRangeExtensions
    {
        public static IFluentWorkbook DeleteWorksheet(this IFluentRange range, string name) => range.Worksheet.DeleteWorksheet(name);

        public static IFluentWorkbook DeleteWorksheet(this IFluentRange range, int index) => range.Worksheet.DeleteWorksheet(index);

        public static IFluentWorksheet OnWorksheet(this IFluentRange range, string name) => range.Worksheet.OnWorksheet(name);

        public static IFluentWorksheet OnWorksheet(this IFluentRange range, int index) => range.Worksheet.OnWorksheet(index);

        public static IFluentRange OnRange(this IFluentRange range, string address) => range.Worksheet.OnRange(address);

        public static IFluentRange OnRange(this IFluentRange range, int row, int column) => range.Worksheet.OnRange(row, column);

        public static IFluentRange OnRange(this IFluentRange range, int beginRow, int beginColumn, int endRow, int endColumn) => range.Worksheet.OnRange(beginRow, beginColumn, endRow, endColumn);
    }
}