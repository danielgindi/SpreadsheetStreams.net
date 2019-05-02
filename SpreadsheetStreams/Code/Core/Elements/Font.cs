using System.Drawing;

namespace SpreadsheetStreams
{
    public struct Font
    {
        public Font(float size = 10.0f, string name = null, bool bold = false)
        {
            Size = size;
            Bold = bold;
            Name = name;
            Color = Color.Black;
            Italic = false;
            Outline = false;
            Shadow = false;
            StrikeThrough = false;
            Underline = FontUnderline.None;
            VerticalAlign = FontVerticalAlign.None;
            Charset = null;
            Family = FontFamily.Automatic;
        }

        public float Size;
        public string Name;
        public bool Bold;
        public Color Color;
        public bool Italic;
        public bool Outline;
        public bool Shadow;
        public bool StrikeThrough;
        public FontUnderline Underline;
        public FontVerticalAlign VerticalAlign;
        public Charset? Charset;
        public FontFamily Family;
    }
}