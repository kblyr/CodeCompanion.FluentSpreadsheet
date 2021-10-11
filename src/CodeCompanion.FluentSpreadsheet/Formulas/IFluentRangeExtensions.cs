using System;
using CodeCompanion.FluentSpreadsheet.Formulas.Arguments;

namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public static class IFluentRangeExtensions
    {
        public static RangeFuncFormulaBuilder Fx(this IFluentRange range, SpreadsheetFunction function) => new(function, range);

        public static FxRangeSumBuilder FxSum(this IFluentRange range) => new(range);
        
        public static IFluentRange FxSum(this IFluentRange range, params object[] numbers)
        {
            var builder = new FxRangeSumBuilder(range);

            for (int index = 0; index < numbers.Length; index++)
                builder.WithArg(FuncFormulaArgCasting.TryCast(numbers[index]));
            
            return builder.Build();
        }

        public static FxRangeAverageBuilder FxAverage(this IFluentRange range) => new(range);

        public static IFluentRange FxAverage(this IFluentRange range, params object[] numbers)
        {
            var builder = new FxRangeAverageBuilder(range);

            for (int index = 0; index < numbers.Length; index++)
                builder.WithArg(FuncFormulaArgCasting.TryCast(numbers[index]));

            return builder.Build();
        }
    }
}