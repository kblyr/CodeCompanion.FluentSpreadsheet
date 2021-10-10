namespace CodeCompanion.FluentSpreadsheet
{
    public interface IFluentWorkbook
    {
        object RawWorkbook { get; }

        IFluentWorksheet AddWorksheet(string name);
        IFluentWorksheet CopyWorksheet(string templateName, string name);
        IFluentWorksheet CopyWorksheet(int templateIndex, string name);
        IFluentWorkbook DeleteWorksheet(string name);
        IFluentWorkbook DeleteWorksheet(int index);
        IFluentWorksheet OnWorksheet(string name); 
        IFluentWorksheet OnWorksheet(int index);
    }
}