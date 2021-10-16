using CodeCompanion.FluentSpreadsheet.Formulas.Arguments;

namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public class FxAverageBuilder : FuncFormulaBuilderBase<FxAverageBuilder>
    {
        public FxAverageBuilder() : base(SpreadsheetFunction.Average)
        {
        }

        protected override FxAverageBuilder GetSelf() => this;

        internal FxAverageBuilder WithArg(IFuncFormulaArg arg) => AddArg(arg);

        public FxAverageBuilder WithNumber(int value) => AddArg(new Int32FuncFormulaArg(value));

        public FxAverageBuilder WithNumber(double value) => AddArg(new DoubleFuncFormulaArg(value));

        public FxAverageBuilder WithNumber(decimal value) => AddArg(new DecimalFuncFormulaArg(value));
    }
}