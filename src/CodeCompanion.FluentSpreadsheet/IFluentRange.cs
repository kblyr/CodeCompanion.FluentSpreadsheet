using CodeCompanion.FluentSpreadsheet.Styles;

namespace CodeCompanion.FluentSpreadsheet
{
    public interface IFluentRange
    {
        object RawRange { get; }
        IFluentWorksheet Worksheet { get; }

        IFluentRange WithValue(object value);
        IFluentRange WithVerticalAlignment(SpreadsheetVerticalAlignment alignment);
        IFluentRange WithHorizontalAlignment(SpreadsheetHorizontalAlignment alignment);
        IFluentRange WithBorder(SpreadsheetBoxSide side, SpreadsheetBorderStyle style);
        IFluentRange Merge(bool isMerged = true);
        IFluentRange WithFontBold(bool isbold = true);
        IFluentRange WithFontUnderline(bool isUnderlined = true);
        IFluentRange WithFontItalic(bool isItalic = true);
        IFluentRange WithNumberFormat(string format);
    }
}