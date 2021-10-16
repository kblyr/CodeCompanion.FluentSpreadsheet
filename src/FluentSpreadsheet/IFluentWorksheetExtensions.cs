namespace CodeCompanion.FluentSpreadsheet
{
    public static class IFluentWorksheetExtensions
    {
        public static IFluentWorkbook DeleteWorksheet(this IFluentWorksheet worksheet, string name) => worksheet.Workbook.DeleteWorksheet(name);

        public static IFluentWorkbook DeleteWorksheet(this IFluentWorksheet worksheet, int index) => worksheet.Workbook.DeleteWorksheet(index);

        public static IFluentWorksheet OnWorksheet(this IFluentWorksheet worksheet, string name) => worksheet.Workbook.OnWorksheet(name);

        public static IFluentWorksheet OnWorksheet(this IFluentWorksheet worksheet, int index) => worksheet.Workbook.OnWorksheet(index);

        public static IFluentWorksheet ForEach<T>(this IFluentWorksheet worksheet, IEnumerable<T> values, Action<IFluentWorksheet, T> handle)
        {
            foreach (var value in values)
                handle(worksheet, value);

            return worksheet;
        }

        public static IFluentWorksheet ForEach<T>(this IFluentWorksheet worksheet, IEnumerable<T> values, Action<IFluentWorksheet, T, int> handle)
        {
            var index = 0;

            foreach(var value in values)
                handle(worksheet, value, index++);

            return worksheet;
        }

        public static IFluentWorksheet While(this IFluentWorksheet worksheet, Func<bool> predicate, Action<IFluentWorksheet> handle)
        {
            while(predicate())
                handle(worksheet);

            return worksheet;
        }

        public static void Save(this IFluentWorksheet worksheet) => worksheet.Workbook.Save();

        public static void SaveAs(this IFluentWorksheet worksheet, string path) => worksheet.Workbook.SaveAs(path);

        public static IFluentWorksheet AddWorksheet(this IFluentWorksheet worksheet, string name) => worksheet.Workbook.AddWorksheet(name);

        public static IFluentWorksheet CopyWorksheet(this IFluentWorksheet worksheet, string templateName, string name) => worksheet.Workbook.CopyWorksheet(templateName, name);

        public static IFluentWorksheet CopyWorksheet(this IFluentWorksheet worksheet, int templateIndex, string name) => worksheet.Workbook.CopyWorksheet(templateIndex, name);
    }
}