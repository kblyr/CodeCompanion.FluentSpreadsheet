namespace CodeCompanion.FluentSpreadsheet.Formulas.Arguments
{
    public struct DoubleFuncFormulaArg : IFuncFormulaArg
    {
        public double Value { get; }

        public DoubleFuncFormulaArg(double value)
        {
            Value = value;
        }

        public string Materialize() => Value.ToString();
    }
}