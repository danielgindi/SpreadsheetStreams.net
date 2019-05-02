namespace SpreadsheetStreams
{
    public struct NumberFormat
    {
        public NumberFormat(NumberFormatType type)
        {
            Type = type;
            Custom = null;
        }

        public NumberFormat(string custom)
        {
            Type = NumberFormatType.Custom;
            Custom = custom;
        }

        public NumberFormatType Type;
        public string Custom;

        #region Predefined

        public static NumberFormat None = new NumberFormat(NumberFormatType.None);
        public static NumberFormat General = new NumberFormat(NumberFormatType.General);
        public static NumberFormat GeneralNumber = new NumberFormat(NumberFormatType.GeneralNumber);
        public static NumberFormat GeneralDate = new NumberFormat(NumberFormatType.GeneralDate);
        public static NumberFormat ShortDate = new NumberFormat(NumberFormatType.ShortDate);
        public static NumberFormat MediumDate = new NumberFormat(NumberFormatType.MediumDate);
        public static NumberFormat LongDate = new NumberFormat(NumberFormatType.LongDate);
        public static NumberFormat ShortTime = new NumberFormat(NumberFormatType.ShortTime);
        public static NumberFormat MediumTime = new NumberFormat(NumberFormatType.MediumTime);
        public static NumberFormat LongTime = new NumberFormat(NumberFormatType.LongTime);
        public static NumberFormat Currency(string code)
        {
            return new NumberFormat("$#,##0.00;[Red]-$#,##0.00".Replace("$", code));
        }
        public static NumberFormat Fixed = new NumberFormat(NumberFormatType.Fixed);
        public static NumberFormat Standard = new NumberFormat(NumberFormatType.Standard);
        public static NumberFormat Percent = new NumberFormat(NumberFormatType.Percent);
        public static NumberFormat Scientific = new NumberFormat(NumberFormatType.Scientific);
        public static NumberFormat YesNo = new NumberFormat(NumberFormatType.YesNo);
        public static NumberFormat TrueFalse = new NumberFormat(NumberFormatType.TrueFalse);
        public static NumberFormat OnOff = new NumberFormat(NumberFormatType.OnOff);

        #endregion
    }
}
