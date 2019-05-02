using System.Collections.Generic;

namespace SpreadsheetStreams
{
    public class Style
    {
        public NumberFormat NumberFormat = NumberFormat.General;
        public Alignment? Alignment;
        public List<Border> Borders;
        public Fill? Fill;
        public Font? Font;
    }
}