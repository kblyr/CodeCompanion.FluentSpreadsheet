namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public interface IRangeFuncFormulaBuilder : IFuncFormulaBuilder
    {
        IFluentRange Range { get; }
        IFluentRange Build();
    }
}