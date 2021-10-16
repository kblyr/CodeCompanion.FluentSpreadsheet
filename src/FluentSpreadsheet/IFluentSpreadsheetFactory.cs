namespace CodeCompanion.FluentSpreadsheet
{
    public interface IFluentSpreadsheetFactory
    {
        IFluentWorkbook Open(string path);
        IFluentWorkbook New();
    }
}