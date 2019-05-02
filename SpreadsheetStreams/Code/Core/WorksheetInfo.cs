using System;

namespace SpreadsheetStreams
{
    public class WorksheetInfo
    {
        internal int Id;
        internal Uri Uri;

        public string Name;
        public float[] ColumnWidths;
        public float? DefaultRowHeight = 15f;
        public float? DefaultColumnWidth = 10f;
    }
}
