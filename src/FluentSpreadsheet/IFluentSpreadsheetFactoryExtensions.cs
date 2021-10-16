namespace CodeCompanion.FluentSpreadsheet
{
    public static class IFluentSpreadsheetFactoryExtensions
    {
        public static IFluentSpreadsheetFactory ForEach<T>(this IFluentSpreadsheetFactory factory, IEnumerable<T> values, Action<IFluentSpreadsheetFactory, T> handle)
        {
            foreach (var value in values)
                handle(factory, value);

            return factory;
        }

        public static IFluentSpreadsheetFactory ForEach<T>(this IFluentSpreadsheetFactory factory, IEnumerable<T> values, Action<IFluentSpreadsheetFactory, T, int> handle)
        {
            var index = 0;

            foreach(var value in values)
                handle(factory, value, index++);

            return factory;
        }
    }
}