using System.IO;

namespace CodeCompanion.FluentSpreadsheet
{
    public sealed class FluentSpreadsheetFactory : IFluentSpreadsheetFactory
    {
        public IFluentWorkbook New() => new FluentWorkbook(new());

        public IFluentWorkbook Open(string path) => new FluentWorkbook(new(new FileInfo(path)));
    }
}