using CodeCompanion.FluentSpreadsheet.Formulas.Arguments;

namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public class FxRangeAverageBuilder : RangeFuncFormulaBuilderBase<FxRangeAverageBuilder>
    {
        public FxRangeAverageBuilder(IFluentRange range) : base(SpreadsheetFunction.Average, range)
        {
        }

        protected override FxRangeAverageBuilder GetSelf() => this;

        internal FxRangeAverageBuilder WithArg(IFuncFormulaArg arg) => AddArg(arg);

        public FxRangeAverageBuilder WithNumber(int value) => AddArg(new Int32FuncFormulaArg(value));

        public FxRangeAverageBuilder WithNumber(double value) => AddArg(new DoubleFuncFormulaArg(value));

        public FxRangeAverageBuilder WithNumber(decimal value) => AddArg(new DecimalFuncFormulaArg(value));
    }
}