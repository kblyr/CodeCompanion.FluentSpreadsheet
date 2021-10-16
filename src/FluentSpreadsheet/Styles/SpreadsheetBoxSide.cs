namespace CodeCompanion.FluentSpreadsheet.Styles
{
    [Flags]
    public enum SpreadsheetBoxSide
    {
        Top         = 1 << 0,
        Bottom      = 1 << 1,
        Left        = 1 << 2,
        Right       = 1 << 3,
        TopBottom   = Top | Bottom,
        LeftRight   = Left | Right,
        All         = TopBottom | LeftRight
    }
}