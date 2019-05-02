using SpreadsheetStreams.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;

namespace SpreadsheetStreams
{
    public class XmlSpreadsheetWriter : SpreadsheetWriter
    {
        #region Constructors

        public XmlSpreadsheetWriter(Stream outputStream, Encoding fileEncoding)
            : base(outputStream ?? new MemoryStream())
        {
            _FileEncoding = fileEncoding ?? Encoding.UTF8;
            _Writer = new StreamWriter(OutputStream, _FileEncoding);
        }

        public XmlSpreadsheetWriter(Stream outputStream)
            : this(outputStream, Encoding.UTF8)
        {

        }

        public XmlSpreadsheetWriter() : this(null, null)
        {
        }

        #endregion

        #region IDisposable

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            if (disposing)
            {
                if (_Writer != null)
                {
                    _Writer.Close();
                    _Writer = null;
                }
            }
        }

        #endregion

        #region Private data

        private CultureInfo _Culture = CultureInfo.InvariantCulture;

        private Encoding _FileEncoding = Encoding.UTF8;
        private StreamWriter _Writer = null;

        private bool _WroteFileStart = false;
        private bool _WroteFileEnd = false;
        private bool _ShouldBeginWorksheet = false;
        private bool _ShouldEndWorksheet = false;
        private bool _ShouldEndRow = false;

        private int _NextStyleIndex = 1;
        private Dictionary<Style, int> _Styles = new Dictionary<Style, int>();
        private int _WorksheetCount = 0;
        private WorksheetInfo _CurrentWorksheetInfo = null;
        private List<string> _Columns = new List<string>();

        #endregion

        #region SpreadsheetWriter - Basic properties

        public override string FileExtension => "xml";
        public override string FileContentType => "text/xml";
        public override bool IsFlatFormat => false;
        public bool ShouldWriteBOM { get; set; } = true;

        #endregion

        #region Public Properties

        #endregion

        #region Private helpers
        
        private static double COLUMN_WIDTH_MULTIPLIER = 5.2875f;
        
        private string GetCellMergeString(int horz, int vert)
        {
            if (horz < 2 && vert < 2) return "";

            return (horz > 1 && vert > 1) ?
                    string.Format(_Culture, @" ss:MergeAcross=""{0}"" ss:MergeDown=""{1}""", horz - 1, vert - 1) :
                    (
                        (horz > 1) ?
                        string.Format(_Culture, @" ss:MergeAcross=""{0}""", horz - 1) :
                        string.Format(_Culture, @" ss:MergeDown=""{0}""", vert - 1)
                    );
        }

        private void Write(string data)
        {
            _Writer.Write(data);
        }

        private string GetStyleId(Style style, bool allowRegister = true)
        {
            if (!_Styles.TryGetValue(style, out var index))
            {
                if (!allowRegister)
                {
                    throw new InvalidOperationException("Registering new styles is not allowed at this time. For XML spreadsheets, styles must be registered before starting to add rows.");
                }

                index = _NextStyleIndex++;
                _Styles[style] = index;
            }

            return "s" + index;
        }

        private string ConvertNumberFormat(NumberFormat format)
        {
            switch (format.Type)
            {
                default:
                case NumberFormatType.General: return "General";
                case NumberFormatType.Custom: return format.Custom;
                case NumberFormatType.GeneralNumber: return "General Number";
                case NumberFormatType.GeneralDate: return "General Date";
                case NumberFormatType.ShortDate: return "Short Date";
                case NumberFormatType.MediumDate: return "Medium Date";
                case NumberFormatType.LongDate: return "Long Date";
                case NumberFormatType.ShortTime: return "Short Time";
                case NumberFormatType.MediumTime: return "Medium Time";
                case NumberFormatType.LongTime: return "[$]dddd, mmmm d, yyyy;@";
                case NumberFormatType.Fixed: return "Fixed";
                case NumberFormatType.Standard: return "Standard";
                case NumberFormatType.Percent: return "Percent";
                case NumberFormatType.Scientific: return "Scientific";
                case NumberFormatType.YesNo: return "Yes/No";
                case NumberFormatType.TrueFalse: return "True/False";
                case NumberFormatType.OnOff: return "On/Off";
            }
        }

        private void WriteStyle(Style style, string id)
        {
            Write(string.Format(@"<ss:Style ss:ID=""{0}"">", id));

            if (style.Alignment != null)
            {
                var align = style.Alignment.Value;

                Write(@"<ss:Alignment"); // Opening tag

                string horz = null;
                switch (align.Horizontal)
                {
                    default:
                    case HorizontalAlignment.Automatic: break;
                    case HorizontalAlignment.Left: horz = @"Left"; break;
                    case HorizontalAlignment.Center: horz = @"Center"; break;
                    case HorizontalAlignment.Right: horz = @"Right"; break;
                    case HorizontalAlignment.Fill: horz = @"Fill"; break;
                    case HorizontalAlignment.Justify: horz = @"Justify"; break;
                    case HorizontalAlignment.CenterAcrossSelection: horz = @"CenterAcrossSelection"; break;
                    case HorizontalAlignment.Distributed: horz = @"Distributed"; break;
                }

                if (horz != null)
                    Write(string.Format(@" ss:Horizontal=""{0}""", horz));

                if (align.Indent > 0) // 0 is default
                    Write(string.Format(@" ss:Indent=""{0}""", align.Indent));

                string readingOrder = null;
                switch (align.ReadingOrder)
                {
                    default:
                    case HorizontalReadingOrder.Context: break;
                    case HorizontalReadingOrder.RightToLeft: readingOrder = @"RightToLeft"; break;
                    case HorizontalReadingOrder.LeftToRight: readingOrder = @"LeftToRight"; break;
                }

                if (readingOrder != null)
                    Write(string.Format(@" ss:ReadingOrder=""{0}""", readingOrder));

                if (align.Rotate != 0.0) // 0 is default
                    Write(string.Format(_Culture, @" ss:Rotate=""{0}""", align.Rotate));

                if (align.ShrinkToFit) // FALSE is default
                    Write(@" ss:ShrinkToFit=""1""");

                string vert = null;
                switch (align.Vertical)
                {
                    default:
                    case VerticalAlignment.Automatic: break;
                    case VerticalAlignment.Top: vert = @"Top"; break;
                    case VerticalAlignment.Bottom: vert = @"Bottom"; break;
                    case VerticalAlignment.Center: vert = @"Center"; break;
                    case VerticalAlignment.Justify: vert = @"Justify"; break;
                    case VerticalAlignment.Distributed: vert = @"Distributed"; break;
                    case VerticalAlignment.JustifyDistributed: vert = @"JustifyDistributed"; break;
                }

                if (vert != null)
                    Write(string.Format(@" ss:Vertical=""{0}""", vert));

                if (align.VerticalText) // FALSE is default
                {
                    Write(@" ss:VerticalText=""1""");
                }

                if (align.WrapText) // FALSE is default
                {
                    Write(@" ss:WrapText=""1""");
                }

                Write("/>"); // Closing tag
            }

            if (style.NumberFormat.Type != NumberFormatType.None)
            {
                Write(string.Format(@"<ss:NumberFormat ss:Format=""{0}""/>", XmlHelper.Escape(ConvertNumberFormat(style.NumberFormat))));
            }

            if (style.Borders != null && style.Borders.Count > 0)
            {
                Write(@"<ss:Borders>"); // Opening tag

                foreach (Border border in style.Borders)
                {
                    Write(@"<ss:Border"); // Opening tag

                    string Position = null;
                    switch (border.Position)
                    {
                        default:
                        case BorderPosition.Left:
                            Position = @"Left";
                            break;
                        case BorderPosition.Top:
                            Position = @"Top";
                            break;
                        case BorderPosition.Right:
                            Position = @"Right";
                            break;
                        case BorderPosition.Bottom:
                            Position = @"Bottom";
                            break;
                        case BorderPosition.DiagonalLeft:
                            Position = @"DiagonalLeft";
                            break;
                        case BorderPosition.DiagonalRight:
                            Position = @"DiagonalRight";
                            break;
                    }
                    Write($@" ss:Position=""{Position}"""); // Required

                    if (!ColorHelper.IsTransparentOrEmpty(border.Color))
                    {
                        Write($@" ss:Color=""#{ColorHelper.GetHexRgb(border.Color)}""");
                    }

                    string LineStyle = null;
                    switch (border.LineStyle)
                    {
                        default:
                        case BorderLineStyle.None: // Default
                            break;
                        case BorderLineStyle.Continuous:
                            LineStyle = @"Continuous";
                            break;
                        case BorderLineStyle.Dash:
                            LineStyle = @"Dash";
                            break;
                        case BorderLineStyle.Dot:
                            LineStyle = @"Dot";
                            break;
                        case BorderLineStyle.DashDot:
                            LineStyle = @"DashDot";
                            break;
                        case BorderLineStyle.DashDotDot:
                            LineStyle = @"DashDotDot";
                            break;
                        case BorderLineStyle.SlantDashDot:
                            LineStyle = @"SlantDashDot";
                            break;
                        case BorderLineStyle.Double:
                            LineStyle = @"Double";
                            break;
                    }
                    if (LineStyle != null)
                    {
                        Write($@" ss:LineStyle=""{LineStyle}""");
                    }

                    if (border.Weight > 0.0) // 0 is default
                    {
                        Write(string.Format(_Culture, @" ss:Weight=""{0}""", border.Weight));
                    }

                    Write("/>"); // Closing tag
                }

                Write("</ss:Borders>"); // Closing tag
            }

            if (style.Fill != null)
            {
                var fill = style.Fill.Value;

                Write(@"<ss:Interior"); // Opening tag

                var bgColor = fill.Color;
                if (fill.Pattern == FillPattern.Solid && bgColor == Color.Empty)
                    bgColor = fill.PatternColor;

                if (!ColorHelper.IsTransparentOrEmpty(bgColor))
                    Write($@" ss:Color=""#{ColorHelper.GetHexRgb(bgColor)}""");

                string pattern = null;
                switch (fill.Pattern)
                {
                    default:
                    case FillPattern.None: break;
                    case FillPattern.Solid: pattern = @"Solid"; break;
                    case FillPattern.Gray75: pattern = @"Gray75"; break;
                    case FillPattern.Gray50: pattern = @"Gray50"; break;
                    case FillPattern.Gray25: pattern = @"Gray25"; break;
                    case FillPattern.Gray125: pattern = @"Gray125"; break;
                    case FillPattern.Gray0625: pattern = @"Gray0625"; break;
                    case FillPattern.HorzStripe: pattern = @"HorzStripe"; break;
                    case FillPattern.VertStripe: pattern = @"VertStripe"; break;
                    case FillPattern.ReverseDiagStripe: pattern = @"ReverseDiagStripe"; break;
                    case FillPattern.DiagCross: pattern = @"DiagCross"; break;
                    case FillPattern.ThickDiagCross: pattern = @"ThickDiagCross"; break;
                    case FillPattern.ThinHorzStripe: pattern = @"ThinHorzStripe"; break;
                    case FillPattern.ThinVertStripe: pattern = @"ThinVertStripe"; break;
                    case FillPattern.ThinReverseDiagStripe: pattern = @"ThinReverseDiagStripe"; break;
                    case FillPattern.ThinDiagStripe: pattern = @"ThinDiagStripe"; break;
                    case FillPattern.ThinHorzCross: pattern = @"ThinHorzCross"; break;
                    case FillPattern.ThinDiagCross: pattern = @"ThinDiagCross"; break;
                }

                if (pattern != null)
                    Write($@" ss:Pattern=""{pattern}""");

                if (fill.Pattern != FillPattern.Solid)
                {
                    if (!ColorHelper.IsTransparentOrEmpty(fill.PatternColor))
                        Write($@" ss:PatternColor=""#{ColorHelper.GetHexRgb(fill.PatternColor)}""");
                }

                Write("/>"); // Closing tag
            }

            if (style.Font != null)
            {
                var font = style.Font.Value;

                Write("<ss:Font"); // Opening tag

                if (font.Bold) // FALSE is default
                    Write(@" ss:Bold=""1""");

                if (!ColorHelper.IsTransparentOrEmpty(font.Color))
                {
                    Write($@" ss:Color=""#{ColorHelper.GetHexRgb(font.Color)}""");
                }

                if (font.Name != null && font.Name.Length > 0)
                {
                    Write($@" ss:FontName=""{XmlHelper.Escape(font.Name)}""");
                }

                if (font.Italic) // FALSE is default
                    Write(@" ss:Italic=""1""");

                if (font.Outline) // FALSE is default
                    Write(@" ss:Outline=""1""");

                if (font.Shadow) // FALSE is default
                    Write(@" ss:Shadow=""1""");

                if (font.Size != 10.0) // 10 is default
                    Write(string.Format(_Culture, @" ss:Size=""{0}""", font.Size));

                if (font.StrikeThrough) // FALSE is default
                    Write(@" ss:StrikeThrough=""1""");

                string underline = null;
                switch (font.Underline)
                {
                    default:
                    case FontUnderline.None: break;
                    case FontUnderline.Single: underline = @"Single"; break;
                    case FontUnderline.Double: underline = @"Double"; break;
                    case FontUnderline.SingleAccounting: underline = @"SingleAccounting"; break;
                    case FontUnderline.DoubleAccounting: underline = @"DoubleAccounting"; break;
                }

                if (underline != null)
                    Write($@" ss:Underline=""{underline}""");

                string verticalAlign = null;
                switch (font.VerticalAlign)
                {
                    default:
                    case FontVerticalAlign.None: break;
                    case FontVerticalAlign.Subscript: verticalAlign = @"Subscript"; break;
                    case FontVerticalAlign.Superscript: verticalAlign = @"Superscript"; break;
                }

                if (verticalAlign != null)
                    Write($@" ss:VerticalAlign=""{verticalAlign}""");

                if (font.Charset != null && (int)font.Charset.Value > 0) // 0 is default
                    Write(string.Format(_Culture, @" ss:CharSet=""{0}""", (int)font.Charset.Value));

                string family = null;
                switch (font.Family)
                {
                    default:
                    case FontFamily.Automatic: break;
                    case FontFamily.Decorative: family = @"Decorative"; break;
                    case FontFamily.Modern: family = @"Modern"; break;
                    case FontFamily.Roman: family = @"Roman"; break;
                    case FontFamily.Script: family = @"Script"; break;
                    case FontFamily.Swiss: family = @"Swiss"; break;
                }

                if (family != null)
                    Write($@" ss:Family=""{family}""");

                Write("/>"); // Closing tag
            }

            Write("</ss:Style>");
        }

        #endregion

        #region Syling

        public override void RegisterStyle(Style style)
        {
            GetStyleId(style, true);
        }

        #endregion

        #region SpreadsheetWriter - Document Lifespan (private)

        private void WritePendingBeginWorksheet()
        {
            if (_ShouldBeginWorksheet)
            {
                Write(string.Format("<Worksheet ss:Name=\"{0}\">", XmlHelper.Escape(_CurrentWorksheetInfo.Name ?? $"Worksheet{_WorksheetCount}")));
                Write("<Table");

                if (_CurrentWorksheetInfo.DefaultRowHeight != null)
                    Write($" ss:DefaultRowHeight=\"{Math.Max(0f, Math.Min(_CurrentWorksheetInfo.DefaultRowHeight.Value, 409.5f)).ToString("G", _Culture)}\"");

                if (_CurrentWorksheetInfo.DefaultColumnWidth != null)
                    Write($" ss:DefaultColumnWidth=\"{(_CurrentWorksheetInfo.DefaultColumnWidth.Value * COLUMN_WIDTH_MULTIPLIER).ToString("G", _Culture)}\"");

                Write(">");

                foreach (string col in _Columns)
                {
                    if (string.IsNullOrEmpty(col)) Write("<Column/>");
                    else Write($"<Column ss:Width=\"{col}\"/>");
                }

                _ShouldBeginWorksheet = false;
            }
        }

        private void WritePendingEndRow()
        {
            if (!_ShouldEndRow) return;

            Write("</Row>");

            _ShouldEndRow = false;
        }

        private void WritePendingEndWorksheet()
        {
            if (_ShouldEndWorksheet)
            {
                WritePendingBeginWorksheet();
                WritePendingEndRow();

                Write("</Table>");
                Write("</Worksheet>");

                _ShouldEndWorksheet = false;
            }
        }

        private void WriteBeginFile()
        {
            if (_WroteFileStart) return;
            _WroteFileStart = true;

            if (OutputStream != null && ShouldWriteBOM)
            {
                OutputStream.Write(_FileEncoding.GetPreamble(), 0, _FileEncoding.GetPreamble().Length);
            }

            Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            Write("<?mso-application progid=\"Excel.Sheet\"?>");
            Write("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"");
            Write(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
            Write(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"");
            Write(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");

            Write("<Styles>");
            Write("<ss:Style ss:ID=\"Default\" ss:Name=\"Normal\">");
            Write("<ss:Alignment ss:Vertical=\"Bottom\"/>");
            Write("</ss:Style>");

            foreach (var style in _Styles.Keys)
            {
                WriteStyle(style, GetStyleId(style, true));
            }

            Write("</Styles>");
        }

        #endregion

        #region SpreadsheetWriter - Document Lifespan (public)

        public override void NewWorksheet(WorksheetInfo info)
        {
            WritePendingEndWorksheet();

            _Columns.Clear();
            _CurrentWorksheetInfo = info;
            _WorksheetCount++;

            if (info.ColumnWidths != null)
            {
                foreach (var width in info.ColumnWidths)
                {
                    string columnString = "";

                    if (width != 0.0)
                    {
                        columnString = (width * COLUMN_WIDTH_MULTIPLIER).ToString("G", _Culture);
                    }

                    _Columns.Add(columnString);
                }
            }

            _ShouldBeginWorksheet = true;
            _ShouldEndWorksheet = true;
        }

        public override void AddRow(Style style = null, float height = 0f)
        {
            if (!_ShouldEndWorksheet)
            {
                throw new InvalidOperationException("Adding new rows is not allowed at this time. Please call NewWorksheet(...) first.");
            }

            if (!_WroteFileStart)
            {
                if (style != null)
                {
                    // A chance to register this style
                    RegisterStyle(style);
                }

                WriteBeginFile();
            }

            WritePendingBeginWorksheet();
            WritePendingEndRow();

            Write("<Row");

            if (style != null)
            {
                Write(string.Format(" ss:StyleID=\"{0}\"", GetStyleId(style, false)));
            }

            if (height != 0)
            {
                Write(string.Format(" ss:Height=\"{0:0.##}\"", height));
            }

            Write(">");

            _ShouldEndRow = true;
        }

        public override void Finish()
        {
            if (!_WroteFileStart)
            {
                WriteBeginFile();
                NewWorksheet(new WorksheetInfo { });
            }

            WritePendingEndWorksheet();

            if (!_WroteFileEnd)
            {
                Write("</Workbook>");
                _WroteFileEnd = true;

                _Writer.Flush();
                _Writer.Close();
            }
        }

        #endregion

        #region SpreadsheetWriter - Cell methods

        public override void AddCell(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";

            Write(string.Format("<Cell{0}{1}><Data ss:Type=\"String\">{2}</Data></Cell>", styleString, merge, XmlHelper.Escape(data)));
        }

        public override void AddCellStringAutoType(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";

            var type = "String";

            if (style?.NumberFormat != null && (
                style.NumberFormat.Type == NumberFormatType.GeneralNumber ||
                style.NumberFormat.Type == NumberFormatType.Scientific ||
                style.NumberFormat.Type == NumberFormatType.Fixed ||
                style.NumberFormat.Type == NumberFormatType.Standard))
            {
                type = "Number";
            }

            Write(string.Format("<Cell{0}{1}><Data ss:Type=\"{2}\">{3}</Data></Cell>", styleString, merge, type, XmlHelper.Escape(data)));
        }

        public override void AddCellForcedString(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            AddCell(data, style, horzCellCount, vertCellCount);
        }

        public override void AddCell(Int32 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            Write(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:G}</Data></Cell>", styleString, merge, data));
        }

        public override void AddCell(Int64 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            Write(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:G}</Data></Cell>", styleString, merge, data));
        }

        public override void AddCell(float data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            Write(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:G}</Data></Cell>", styleString, merge, data));
        }

        public override void AddCell(double data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            Write(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:G}</Data></Cell>", styleString, merge, data));
        }

        public override void AddCell(decimal data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            Write(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:G}</Data></Cell>", styleString, merge, data));
        }

        public override void AddCell(DateTime data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            Write(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"DateTime\">{2}</Data></Cell>",
                styleString, merge,
                data.Year <= 1 ? "" : data.ToString("yyyy-MM-ddTHH:mm:ss.fff")));
        }

        public override void AddCellFormula(string formula, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";

            Write(string.Format("<Cell{0}{1} ss:Formula=\"{2}\" />",
                styleString, merge, XmlHelper.Escape(formula)));
        }

        #endregion
    }
}
