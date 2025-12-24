using System;
#nullable enable

namespace SpreadsheetStreams.Code.Excel
{
    public class AutoFitConfig
    {
        public AutoFitConfig() { }


        public AutoFitMeasureDelegate? Measure { get; set; }
        public bool Multiline { get; set; }

        /// <summary>
        /// Multiplier for size measurement (character size -> column size).
        /// Not applied for <see cref="Measure"/> results.
        /// </summary>
        public float Multiplier { get; set; } = 1.2f;

        /// <summary>
        /// Max length for auto fit measurement.
        /// </summary>
        public float MaxLength { get; set; } = 120f;
    }
}
