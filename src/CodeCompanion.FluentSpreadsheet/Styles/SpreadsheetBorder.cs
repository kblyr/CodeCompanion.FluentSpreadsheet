namespace CodeCompanion.FluentSpreadsheet.Styles
{
    public record SpreadsheetBorder
    {
        public SpreadsheetColor? Color { get; init; }
        public SpreadsheetBorderStyle? Style { get; init; }
    }
}