using System.Collections.Generic;
using System.Linq;
using System.Text;
using CodeCompanion.FluentSpreadsheet.Formulas.Arguments;

namespace CodeCompanion.FluentSpreadsheet.Formulas
{
    public abstract class FuncFormulaBuilderBase<TSelf> : IFuncFormulaBuilder where TSelf : IFuncFormulaBuilder 
    {
        private readonly List<IFuncFormulaArg> _args = new();
        private readonly SpreadsheetFunction _function;

        protected FuncFormulaBuilderBase(SpreadsheetFunction function)
        {
            _function = function;
        }

        protected abstract TSelf GetSelf();

        protected virtual TSelf AddArg(IFuncFormulaArg arg)
        {
            _args.Add(arg);
            return GetSelf();
        }

        public string BuildString() => new StringBuilder(SpreadsheetFunctionConverter.Convert(_function))
            .Append("(")
            .Append(_args.Select(_ => _.Materialize()).Aggregate((set, item) => $"{set},{item}"))
            .Append(")")
            .ToString();
    }
}