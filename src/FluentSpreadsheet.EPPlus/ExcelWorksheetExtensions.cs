using OfficeOpenXml;

namespace CodeCompanion.FluentSpreadsheet
{
    public static class ExcelWorksheetExtensions
    {
        public static ExcelRange OnRange(this ExcelWorksheet worksheet, string address) => worksheet.Cells[address];

        public static ExcelRange OnRange(this ExcelWorksheet worksheet, int row, int column) => worksheet.Cells[row, column];

        public static ExcelRange OnRange(this ExcelWorksheet worksheet, int beginRow, int beginColumn, int endRow, int endColumn) => worksheet.Cells[beginRow, beginColumn, endRow, endColumn];
    }
}