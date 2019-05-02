using System.Drawing;
using System.Globalization;

namespace SpreadsheetStreams.Util
{
    internal static class XmlHelper
    {
        internal static string Escape(string value)
        {
            value = value.Replace("&", "&amp;");
            value = value.Replace("<", "&lt;");
            value = value.Replace(">", "&gt;");
            value = value.Replace(@"""", "&quot;");
            value = value.Replace("'", "&apos;");
            value = value.Replace("\r", "&#xD;");
            value = value.Replace("\n", "&#xA;");

            return value;
        }
    }
}
