using CodeCompanion.FluentSpreadsheet.Formulas.Arguments;

namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public class FuncFormulaBuilder : FuncFormulaBuilderBase<FuncFormulaBuilder>
    {
        public FuncFormulaBuilder(SpreadsheetFunction function) : base(function)
        {
        }

        protected override FuncFormulaBuilder GetSelf() => this;
    }
}