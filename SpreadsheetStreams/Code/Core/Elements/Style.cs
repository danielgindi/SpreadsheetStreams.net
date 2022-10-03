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

        public Style Clone()
        {
            return new Style
            {
                NumberFormat = NumberFormat,
                Alignment = Alignment,
                Borders = Borders == null ? null : new List<Border>(Borders),
                Fill = Fill,
                Font = Font,
            };
        }
    }
}