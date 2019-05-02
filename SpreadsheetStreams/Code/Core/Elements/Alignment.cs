using System;

namespace SpreadsheetStreams
{
    public struct Alignment
    {
        public HorizontalAlignment Horizontal;
        public VerticalAlignment Vertical;
        public int Indent;
        public HorizontalReadingOrder ReadingOrder;
        public double Rotate;
        public bool ShrinkToFit;
        public bool VerticalText;
        public bool WrapText;
    }
}