using System.Drawing;

namespace CodeCompanion.FluentSpreadsheet.Styles
{
    public static class SpreadsheetColorExtensions
    {
        public static Color ToDotNetColor(this SpreadsheetColor color) => Color.FromArgb(color.Alpha, color.Red, color.Green, color.Blue);
    }
}