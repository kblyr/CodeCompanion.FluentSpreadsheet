namespace CodeCompanion.FluentSpreadsheet
{
    public sealed class EPPlusSpreadsheetContext<TSharedData> : SpreadsheetContextBase<TSharedData>, ISpreadsheetContext<TSharedData>
    {
        public override ISpreadsheetContext<TSharedData> Self => this;
    }
}