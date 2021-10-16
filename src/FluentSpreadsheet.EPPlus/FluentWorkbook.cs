using System.IO;
using OfficeOpenXml;

namespace CodeCompanion.FluentSpreadsheet
{
    sealed class FluentWorkbook : IFluentWorkbook
    {
        private readonly ExcelPackage _package;
        private readonly ExcelWorkbook _rawWorkbook;

        public object RawWorkbook => _rawWorkbook;

        public FluentWorkbook(ExcelPackage package)
        {
            _package = package;
            _rawWorkbook = package.Workbook;
        }

        public IFluentWorksheet AddWorksheet(string name)
        {
            var rawWorksheet = _rawWorkbook.Worksheets.Add(name);
            return new FluentWorksheet(rawWorksheet, this);
        }

        public IFluentWorksheet CopyWorksheet(string templateName, string name)
        {
            var templateRawWorksheet = _rawWorkbook.Worksheets[templateName];
            var rawWorksheet = _rawWorkbook.Worksheets.Add(name, templateRawWorksheet);
            return new FluentWorksheet(rawWorksheet, this);
        }

        public IFluentWorksheet CopyWorksheet(int templateIndex, string name)
        {
            var templateRawWorksheet = _rawWorkbook.Worksheets[templateIndex];
            var rawWorksheet = _rawWorkbook.Worksheets.Add(name, templateRawWorksheet);
            return new FluentWorksheet(rawWorksheet, this);
        }

        public IFluentWorkbook DeleteWorksheet(string name)
        {
            _rawWorkbook.Worksheets.Delete(name);
            return this;
        }

        public IFluentWorkbook DeleteWorksheet(int index)
        {
            _rawWorkbook.Worksheets.Delete(index);
            return this;
        }

        public void Save()
        {
            _package.Save();
            _package.Dispose();
        }

        public void SaveAs(string path)
        {
            _package.SaveAs(new FileInfo(path));
            _package.Dispose();
        }

        public IFluentWorksheet OnWorksheet(string name) => new FluentWorksheet(_rawWorkbook.OnWorksheet(name), this);

        public IFluentWorksheet OnWorksheet(int index) => new FluentWorksheet(_rawWorkbook.OnWorksheet(index), this);
    }
}