namespace CodeCompanion.FluentSpreadsheet.Formulas.Arguments
{
    public struct Int32FuncFormulaArg : IFuncFormulaArg
    {
        public int Value { get; }

        public Int32FuncFormulaArg(int value)
        {
            Value = value;
        }

        public string Materialize() => Value.ToString();
    }
}