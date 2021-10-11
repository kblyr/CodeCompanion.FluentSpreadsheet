namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public class RangeFuncFormulaBuilder : RangeFuncFormulaBuilderBase<RangeFuncFormulaBuilder>
    {
        public RangeFuncFormulaBuilder(SpreadsheetFunction function, IFluentRange range) : base(function, range)
        {
        }

        protected override RangeFuncFormulaBuilder GetSelf() => this;
    }
}