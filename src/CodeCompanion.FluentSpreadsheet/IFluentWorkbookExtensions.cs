using System;
using System.Collections.Generic;

namespace CodeCompanion.FluentSpreadsheet
{
    public static class IFluentWorkbookExtensions
    {
        public static IFluentWorkbook ForEach<T>(this IFluentWorkbook workbook, IEnumerable<T> values, Action<IFluentWorkbook, T> handle)
        {
            foreach (var value in values)
                handle(workbook, value);

            return workbook;
        }

        public static IFluentWorkbook ForEach<T>(this IFluentWorkbook workbook, IEnumerable<T> values, Action<IFluentWorkbook, T, int> handle)
        {
            var index = 0;

            foreach (var value in values)
                handle(workbook, value, index++);

            return workbook;
        }

        public static IFluentWorkbook While(this IFluentWorkbook workbook, Func<bool> predicate, Action<IFluentWorkbook> handle)
        {
            while(predicate())
                handle(workbook);

            return workbook;
        }
    }
}