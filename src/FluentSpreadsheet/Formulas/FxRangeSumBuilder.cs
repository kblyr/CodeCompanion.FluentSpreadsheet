using CodeCompanion.FluentSpreadsheet.Formulas.Arguments;

namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public class FxRangeSumBuilder : RangeFuncFormulaBuilderBase<FxRangeSumBuilder>
    {
        public FxRangeSumBuilder(IFluentRange range) : base(SpreadsheetFunction.Sum, range)
        {
        }

        protected override FxRangeSumBuilder GetSelf() => this;

        internal FxRangeSumBuilder WithArg(IFuncFormulaArg arg) => AddArg(arg);

        public FxRangeSumBuilder WithNumber(int value) => AddArg(new Int32FuncFormulaArg(value));

        public FxRangeSumBuilder WithNumber(double value) => AddArg(new DoubleFuncFormulaArg(value));

        public FxRangeSumBuilder WithNumber(decimal value) => AddArg(new DecimalFuncFormulaArg(value));
    }
}