using OfficeOpenXml;

namespace CodeCompanion.FluentSpreadsheet
{
    sealed class FluentWorksheet : IFluentWorksheet
    {
        private readonly ExcelWorksheet _rawWorksheet;

        public object RawWorksheet => _rawWorksheet;
        public IFluentWorkbook Workbook { get; }

        public FluentWorksheet(ExcelWorksheet rawWorksheet, IFluentWorkbook workbook)
        {
            _rawWorksheet = rawWorksheet;
            Workbook = workbook;
        }

        public IFluentRange OnRange(string address) => new FluentRange(_rawWorksheet.OnRange(address), this);

        public IFluentRange OnRange(int row, int column) => new FluentRange(_rawWorksheet.OnRange(row + 1, column + 1), this);

        public IFluentRange OnRange(int beginRow, int beginColumn, int endRow, int endColumn) => new FluentRange(_rawWorksheet.OnRange(beginRow + 1, beginColumn + 1, endRow + 1, endColumn + 1), this);
    }
}