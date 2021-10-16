using System.Text;

namespace CodeCompanion.FluentSpreadsheet
{
    public static class StringBuilderExtensions
    {
        public static StringBuilder AppendIf(this StringBuilder builder, Func<string> valueCallback, bool condition) 
        {
            if (condition)
                builder.Append(valueCallback());

            return builder;
        }
    }
}