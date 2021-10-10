namespace CodeCompanion.FluentSpreadsheet
{
    public interface IFluentWorksheet
    {
        object RawWorksheet { get; }
        IFluentWorkbook Workbook { get; }

        IFluentRange OnRange(string address);
        IFluentRange OnRange(int row, int column);
        IFluentRange OnRange(int beginRow, int beginColumn, int endRow, int endColumn);
    }
}