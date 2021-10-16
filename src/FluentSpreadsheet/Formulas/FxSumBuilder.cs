using CodeCompanion.FluentSpreadsheet.Formulas.Arguments;

namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public class FxSumBuilder : FuncFormulaBuilderBase<FxSumBuilder>
    {
        public FxSumBuilder() : base(SpreadsheetFunction.Sum)
        {
        }

        protected override FxSumBuilder GetSelf() => this;

        internal FxSumBuilder WithArg(IFuncFormulaArg arg) => AddArg(arg);

        public FxSumBuilder WithNumber(int value) => AddArg(new Int32FuncFormulaArg(value));

        public FxSumBuilder WithNumber(double value) => AddArg(new DoubleFuncFormulaArg(value));
        
        public FxSumBuilder WithNumber(decimal value) => AddArg(new DecimalFuncFormulaArg(value));
    }
}