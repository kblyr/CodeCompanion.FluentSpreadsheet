using CodeCompanion.FluentSpreadsheet.Styles;
using OfficeOpenXml;

namespace CodeCompanion.FluentSpreadsheet
{
    sealed class FluentRange : IFluentRange
    {
        private readonly ExcelRange _rawRange;

        public object RawRange => _rawRange;
        public IFluentWorksheet Worksheet { get; }

        public FluentRange(ExcelRange rawRange, IFluentWorksheet worksheet)
        {
            _rawRange = rawRange;
            Worksheet = worksheet;
        }

        public IFluentRange WithValue(object value)
        {
            _rawRange.Value = value;
            return this;
        }

        public IFluentRange WithVerticalAlignment(SpreadsheetVerticalAlignment alignment)
        {
            _rawRange.Style.VerticalAlignment = alignment.ToRaw();
            return this;
        }

        public IFluentRange WithHorizontalAlignment(SpreadsheetHorizontalAlignment alignment)
        {
            _rawRange.Style.HorizontalAlignment = alignment.ToRaw();
            return this;
        }

        public IFluentRange WithBorder(SpreadsheetBoxSide side, SpreadsheetBorder border)
        {
            return this
                .WithBorderIf(_rawRange.Style.Border.Top, border, side.HasFlag(SpreadsheetBoxSide.Top))
                .WithBorderIf(_rawRange.Style.Border.Bottom, border, side.HasFlag(SpreadsheetBoxSide.Bottom))
                .WithBorderIf(_rawRange.Style.Border.Left, border, side.HasFlag(SpreadsheetBoxSide.Left))
                .WithBorderIf(_rawRange.Style.Border.Right, border, side.HasFlag(SpreadsheetBoxSide.Right));
        }

        public IFluentRange Merge(bool isMerged = true)
        {
            _rawRange.Merge = isMerged;
            return this;
        }

        public IFluentRange WithFontBold(bool isbold = true)
        {
            _rawRange.Style.Font.Bold = isbold;
            return this;
        }

        public IFluentRange WithFontUnderline(bool isUnderlined = true)
        {
            _rawRange.Style.Font.UnderLine = isUnderlined;
            return this;
        }

        public IFluentRange WithFontItalic(bool isItalic = true)
        {
            _rawRange.Style.Font.Italic = isItalic;
            return this;
        }

        public IFluentRange WithNumberFormat(string format)
        {
            _rawRange.Style.Numberformat.Format = format;
            return this;
        }
    }
}