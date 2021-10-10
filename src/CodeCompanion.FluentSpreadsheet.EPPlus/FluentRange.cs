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

        public IFluentRange WithBorder(SpreadsheetBoxSide side, SpreadsheetBorderStyle style)
        {
            var rawStyle = style.ToRaw();

            if (side.HasFlag(SpreadsheetBoxSide.Top))
                _rawRange.Style.Border.Top.Style = rawStyle;

            if (side.HasFlag(SpreadsheetBoxSide.Bottom))
                _rawRange.Style.Border.Bottom.Style = rawStyle;

            if (side.HasFlag(SpreadsheetBoxSide.Left))
                _rawRange.Style.Border.Left.Style = rawStyle;

            if (side.HasFlag(SpreadsheetBoxSide.Right))
                _rawRange.Style.Border.Right.Style = rawStyle;

            return this;
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

        public IFluentRange WithFontUnderline(SpreadsheetUnderlineStyle style)
        {
            _rawRange.Style.Font.UnderLine = true;
            _rawRange.Style.Font.UnderLineType = style.ToRaw();
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