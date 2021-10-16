using System.Drawing;

namespace CodeCompanion.FluentSpreadsheet.Styles
{
    public struct SpreadsheetColor
    {
        public byte Red { get; }
        public byte Green { get; }
        public byte Blue { get; }
        public byte Alpha { get; }

        public SpreadsheetColor(byte red, byte green, byte blue, byte alpha)
        {
            Red = red;
            Green = green;
            Blue = blue;
            Alpha = alpha;
        }

        public static SpreadsheetColor FromRgba(byte red, byte green, byte blue, byte alpha) => new(red, green, blue, alpha);

        public static SpreadsheetColor FromDotNetColor(Color color) => new(color.R, color.G, color.B, color.A);
    }
}