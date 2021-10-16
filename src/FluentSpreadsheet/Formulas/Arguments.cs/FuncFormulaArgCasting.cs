using System;

namespace CodeCompanion.FluentSpreadsheet.Formulas.Arguments
{
    public static class FuncFormulaArgCasting
    {
        public static IFuncFormulaArg TryCast(object obj) => obj switch
        {
            int value => new Int32FuncFormulaArg(value),
            decimal value => new DecimalFuncFormulaArg(value),
            double value => new DoubleFuncFormulaArg(value),
            IFuncFormulaBuilder value => new BuilderFuncFormulaArg(value),
            IFluentRange value => new RangeRefFuncFormulaArg(value.FullAddress),
            RangeRef value => new RangeRefFuncFormulaArg(value.Address),
            _ => throw new InvalidOperationException($"Type '{obj.GetType()}' is not valid number")
        };
    }
}