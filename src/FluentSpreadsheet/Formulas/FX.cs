using CodeCompanion.FluentSpreadsheet.Formulas.Arguments;

namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public static class FX
    {
        public static FxSumBuilder Sum() => new();

        public static FxSumBuilder Sum(params object[] numbers)
        {
            var builder = new FxSumBuilder();

            for (int index = 0; index < numbers.Length; index++)
                builder.WithArg(FuncFormulaArgCasting.TryCast(numbers[index]));

            return builder;
        }

        public static FxAverageBuilder Average() => new();

        public static FxAverageBuilder AverageBuilder(params object[] numbers)
        {
            var builder = new FxAverageBuilder();

            for (int index = 0; index < numbers.Length; index++)
                builder.WithArg(FuncFormulaArgCasting.TryCast(numbers[index]));

            return builder;
        }
    }
}