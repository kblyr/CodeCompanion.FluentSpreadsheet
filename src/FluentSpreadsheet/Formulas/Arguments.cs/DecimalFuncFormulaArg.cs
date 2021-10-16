namespace CodeCompanion.FluentSpreadsheet.Formulas.Arguments
{
    public struct DecimalFuncFormulaArg : IFuncFormulaArg
    {
        public decimal Value { get; }

        public DecimalFuncFormulaArg(decimal value)
        {
            Value = value;
        }

        public string Materialize() => Value.ToString();
    }
}