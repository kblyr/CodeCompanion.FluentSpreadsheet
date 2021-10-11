namespace CodeCompanion.FluentSpreadsheet.Formulas.Arguments
{
    public struct RangeRefFuncFormulaArg : IFuncFormulaArg
    {
        public string Address { get; }

        public RangeRefFuncFormulaArg(string address)
        {
            Address = address;
        }

        public string Materialize() => Address;
    }
}