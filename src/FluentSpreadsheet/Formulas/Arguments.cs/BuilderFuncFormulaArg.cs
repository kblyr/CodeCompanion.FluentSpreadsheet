namespace CodeCompanion.FluentSpreadsheet.Formulas.Arguments
{
    public struct BuilderFuncFormulaArg : IFuncFormulaArg
    {
        public IFuncFormulaBuilder Value { get; }

        public BuilderFuncFormulaArg(IFuncFormulaBuilder value)
        {
            Value = value;
        }

        public string Materialize() => Value.BuildString();
    }
}