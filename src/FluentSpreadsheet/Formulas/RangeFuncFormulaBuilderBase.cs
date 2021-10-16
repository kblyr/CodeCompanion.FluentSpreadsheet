namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public abstract class RangeFuncFormulaBuilderBase<TSelf> : FuncFormulaBuilderBase<TSelf>, IRangeFuncFormulaBuilder where TSelf : IRangeFuncFormulaBuilder
    {
        public IFluentRange Range { get; }

        protected RangeFuncFormulaBuilderBase(SpreadsheetFunction function, IFluentRange range) : base(function)
        {
            Range = range;
        }

        public virtual IFluentRange Build() => Range.WithFormula(BuildString());
    }
}