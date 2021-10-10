using OfficeOpenXml;

namespace CodeCompanion.FluentSpreadsheet
{
    public static class ExcelWorkbookExtensions
    {
        public static ExcelWorksheet OnWorksheet(this ExcelWorkbook workbook, string name) => workbook.Worksheets[name];

        public static ExcelWorksheet OnWorksheet(this ExcelWorkbook workbook, int index) => workbook.Worksheets[index];
    }
}