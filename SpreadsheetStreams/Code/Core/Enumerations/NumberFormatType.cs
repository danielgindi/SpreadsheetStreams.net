namespace SpreadsheetStreams
{
    public enum NumberFormatType
    {
        /// <summary>
        /// No format specified or applied
        /// </summary>
        None,

        /// <summary>
        /// ...
        /// </summary>
        Custom,

        /// <summary>
        /// General
        /// </summary>
        General,

        /// <summary>
        /// 0
        /// </summary>
        GeneralNumber,

        /// <summary>
        /// dd/mm/yyyy h:mm
        /// </summary>
        GeneralDate,

        /// <summary>
        /// d/m/yyyy
        /// </summary>
        ShortDate,

        /// <summary>
        /// d-mmm-yy
        /// </summary>
        MediumDate,

        /// <summary>
        /// [$]dddd, mmmm d, yyyy;@
        /// </summary>
        LongDate,

        /// <summary>
        /// h:mm
        /// </summary>
        ShortTime,

        /// <summary>
        /// mm AM/PM
        /// </summary>
        MediumTime,

        /// <summary>
        /// h:mm:ss AM/PM
        /// </summary>
        LongTime,

        /// <summary>
        /// 0.00
        /// </summary>
        Fixed,

        /// <summary>
        /// #,##0.00
        /// </summary>
        Standard,

        /// <summary>
        /// 0.00%
        /// </summary>
        Percent,

        /// <summary>
        /// ##0.0E+0
        /// </summary>
        Scientific,

        /// <summary>
        /// "Yes";"Yes";"No"
        /// </summary>
        YesNo,

        /// <summary>
        /// "True";"True";"False"
        /// </summary>
        TrueFalse,

        /// <summary>
        /// "On";"On";"Off"
        /// </summary>
        OnOff,
    }
}
