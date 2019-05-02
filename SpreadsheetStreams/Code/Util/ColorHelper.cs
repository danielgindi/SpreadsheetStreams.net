using System.Drawing;
using System.Globalization;

namespace SpreadsheetStreams.Util
{
    internal static class ColorHelper
    {
        private static CultureInfo InvariantCulture = CultureInfo.InvariantCulture;

        public static bool IsTransparentOrEmpty(Color Input)
        {
            return Input.IsEmpty || Input == Color.Transparent;
        }

        internal static string GetHexRgb(Color color)
        {
            if (color == Color.Empty) return @"";
            return string.Format(InvariantCulture, @"{0:x2}{1:x2}{2:x2}", color.R, color.G, color.B);
        }

        internal static string GetHexArgb(Color color)
        {
            if (color == Color.Empty) return @"";
            return string.Format(InvariantCulture, @"{0:x2}{1:x2}{2:x2}{3:x2}", color.A, color.R, color.G, color.B);
        }
    }
}
