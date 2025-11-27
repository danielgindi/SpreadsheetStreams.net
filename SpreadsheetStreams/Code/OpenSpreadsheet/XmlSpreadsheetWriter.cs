using SpreadsheetStreams.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

#nullable enable

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

                if (_XmlWriterHelper != null)
                {
                    _XmlWriterHelper.Dispose();
                    _XmlWriterHelper = null;
                }
            }
        }

        #endregion

        #region Private data

        private CultureInfo _Culture = CultureInfo.InvariantCulture;

        private Encoding _FileEncoding = Encoding.UTF8;
        private StreamWriter? _Writer = null;

        private bool _WroteFileStart = false;
        private bool _WroteFileEnd = false;
        private bool _ShouldBeginWorksheet = false;
        private bool _ShouldEndWorksheet = false;
        private bool _ShouldEndRow = false;

        private int _NextStyleIndex = 1;
        private Dictionary<Style, int> _Styles = new Dictionary<Style, int>();
        private int _WorksheetCount = 0;
        private WorksheetInfo? _CurrentWorksheetInfo = null;
        private List<string> _Columns = new List<string>();

        private XmlWriterHelper? _XmlWriterHelper = new XmlWriterHelper();

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
                    $@" ss:MergeAcross=""{horz - 1}"" ss:MergeDown=""{vert - 1}""" :
                    (
                        (horz > 1) ?
                        $@" ss:MergeAcross=""{horz - 1}""" :
                        $@" ss:MergeDown=""{vert - 1}"""
                    );
        }
        
        private async Task WriteAsync(string data)
        {
            await _Writer!.WriteAsync(data).ConfigureAwait(false);
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

        private async Task WriteStyleAsync(Style style, string id)
        {
            await WriteAsync($@"<ss:Style ss:ID=""{id}"">");

            if (style.Alignment != null)
            {
                var align = style.Alignment.Value;

                await WriteAsync(@"<ss:Alignment"); // Opening tag

                string? horz = null;
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
                    await WriteAsync(string.Format(@" ss:Horizontal=""{0}""", horz));

                if (align.Indent > 0) // 0 is default
                    await WriteAsync(string.Format(@" ss:Indent=""{0}""", align.Indent));

                string? readingOrder = null;
                switch (align.ReadingOrder)
                {
                    default:
                    case HorizontalReadingOrder.Context: break;
                    case HorizontalReadingOrder.RightToLeft: readingOrder = @"RightToLeft"; break;
                    case HorizontalReadingOrder.LeftToRight: readingOrder = @"LeftToRight"; break;
                }

                if (readingOrder != null)
                    await WriteAsync(string.Format(@" ss:ReadingOrder=""{0}""", readingOrder));

                if (align.Rotate != 0.0) // 0 is default
                    await WriteAsync(string.Format(_Culture, @" ss:Rotate=""{0}""", align.Rotate));

                if (align.ShrinkToFit) // FALSE is default
                    await WriteAsync(@" ss:ShrinkToFit=""1""");

                string? vert = null;
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
                    await WriteAsync(string.Format(@" ss:Vertical=""{0}""", vert));

                if (align.VerticalText) // FALSE is default
                {
                    await WriteAsync(@" ss:VerticalText=""1""");
                }

                if (align.WrapText) // FALSE is default
                {
                    await WriteAsync(@" ss:WrapText=""1""");
                }

                await WriteAsync("/>"); // Closing tag
            }

            if (style.NumberFormat.Type != NumberFormatType.None)
            {
                await WriteAsync(string.Format(@"<ss:NumberFormat ss:Format=""{0}""/>", XmlWriterHelper.EscapeSimpleXmlAttr(ConvertNumberFormat(style.NumberFormat))));
            }

            if (style.Borders != null && style.Borders.Count > 0)
            {
                await WriteAsync(@"<ss:Borders>"); // Opening tag

                foreach (Border border in style.Borders)
                {
                    await WriteAsync(@"<ss:Border"); // Opening tag

                    string? position = null;
                    switch (border.Position)
                    {
                        default:
                        case BorderPosition.Left:
                            position = @"Left";
                            break;
                        case BorderPosition.Top:
                            position = @"Top";
                            break;
                        case BorderPosition.Right:
                            position = @"Right";
                            break;
                        case BorderPosition.Bottom:
                            position = @"Bottom";
                            break;
                        case BorderPosition.DiagonalLeft:
                            position = @"DiagonalLeft";
                            break;
                        case BorderPosition.DiagonalRight:
                            position = @"DiagonalRight";
                            break;
                    }
                    await WriteAsync($@" ss:Position=""{position}"""); // Required

                    if (!ColorHelper.IsTransparentOrEmpty(border.Color))
                    {
                        await WriteAsync($@" ss:Color=""#{ColorHelper.GetHexRgb(border.Color)}""");
                    }

                    string? lineStyle = null;
                    switch (border.LineStyle)
                    {
                        default:
                        case BorderLineStyle.None: // Default
                            break;
                        case BorderLineStyle.Continuous:
                            lineStyle = @"Continuous";
                            break;
                        case BorderLineStyle.Dash:
                            lineStyle = @"Dash";
                            break;
                        case BorderLineStyle.Dot:
                            lineStyle = @"Dot";
                            break;
                        case BorderLineStyle.DashDot:
                            lineStyle = @"DashDot";
                            break;
                        case BorderLineStyle.DashDotDot:
                            lineStyle = @"DashDotDot";
                            break;
                        case BorderLineStyle.SlantDashDot:
                            lineStyle = @"SlantDashDot";
                            break;
                        case BorderLineStyle.Double:
                            lineStyle = @"Double";
                            break;
                    }
                    if (lineStyle != null)
                    {
                        await WriteAsync($@" ss:LineStyle=""{lineStyle}""");
                    }

                    if (border.Weight > 0.0) // 0 is default
                    {
                        await WriteAsync(string.Format(_Culture, @" ss:Weight=""{0}""", border.Weight));
                    }

                    await WriteAsync("/>"); // Closing tag
                }

                await WriteAsync("</ss:Borders>"); // Closing tag
            }

            if (style.Fill != null)
            {
                var fill = style.Fill.Value;

                await WriteAsync(@"<ss:Interior"); // Opening tag

                var bgColor = fill.Color;
                if (fill.Pattern == FillPattern.Solid && bgColor == Color.Empty)
                    bgColor = fill.PatternColor;

                if (!ColorHelper.IsTransparentOrEmpty(bgColor))
                    await WriteAsync($@" ss:Color=""#{ColorHelper.GetHexRgb(bgColor)}""");

                string? pattern = null;
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
                    await WriteAsync($@" ss:Pattern=""{pattern}""");

                if (fill.Pattern != FillPattern.Solid)
                {
                    if (!ColorHelper.IsTransparentOrEmpty(fill.PatternColor))
                        await WriteAsync($@" ss:PatternColor=""#{ColorHelper.GetHexRgb(fill.PatternColor)}""");
                }

                await WriteAsync("/>"); // Closing tag
            }

            if (style.Font != null)
            {
                var font = style.Font.Value;

                await WriteAsync("<ss:Font"); // Opening tag

                if (font.Bold) // FALSE is default
                    await WriteAsync(@" ss:Bold=""1""");

                if (!ColorHelper.IsTransparentOrEmpty(font.Color))
                {
                    await WriteAsync($@" ss:Color=""#{ColorHelper.GetHexRgb(font.Color)}""");
                }

                if (font.Name != null && font.Name.Length > 0)
                {
                    await WriteAsync($@" ss:FontName=""{XmlWriterHelper.EscapeSimpleXmlAttr(font.Name)}""");
                }

                if (font.Italic) // FALSE is default
                    await WriteAsync(@" ss:Italic=""1""");

                if (font.Outline) // FALSE is default
                    await WriteAsync(@" ss:Outline=""1""");

                if (font.Shadow) // FALSE is default
                    await WriteAsync(@" ss:Shadow=""1""");

                if (font.Size != 10.0) // 10 is default
                    await WriteAsync(string.Format(_Culture, @" ss:Size=""{0}""", font.Size));

                if (font.StrikeThrough) // FALSE is default
                    await WriteAsync(@" ss:StrikeThrough=""1""");

                string? underline = null;
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
                    await WriteAsync($@" ss:Underline=""{underline}""");

                string? verticalAlign = null;
                switch (font.VerticalAlign)
                {
                    default:
                    case FontVerticalAlign.None: break;
                    case FontVerticalAlign.Subscript: verticalAlign = @"Subscript"; break;
                    case FontVerticalAlign.Superscript: verticalAlign = @"Superscript"; break;
                }

                if (verticalAlign != null)
                    await WriteAsync($@" ss:VerticalAlign=""{verticalAlign}""");

                if (font.Charset != null && (int)font.Charset.Value > 0) // 0 is default
                    await WriteAsync(string.Format(_Culture, @" ss:CharSet=""{0}""", (int)font.Charset.Value));

                string? family = null;
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
                    await WriteAsync($@" ss:Family=""{family}""");

                await WriteAsync("/>"); // Closing tag
            }

            await WriteAsync("</ss:Style>");
        }

        #endregion

        #region Syling

        public override void RegisterStyle(Style style)
        {
            GetStyleId(style, true);
        }

        #endregion

        #region SpreadsheetWriter - Document Lifespan (private)
        
        private async Task WritePendingBeginWorksheetAsync()
        {
            if (_ShouldBeginWorksheet)
            {
                await WriteAsync($"<Worksheet ss:Name=\"{_XmlWriterHelper!.EscapeAttribute(_CurrentWorksheetInfo!.Name ?? $"Worksheet{_WorksheetCount}")}\"");

                if (_CurrentWorksheetInfo.RightToLeft != null)
                    await WriteAsync($" ss:RightToLeft=\"{(_CurrentWorksheetInfo.RightToLeft == true ? "1" : "0")}\"");

                await WriteAsync("><Table");

                if (_CurrentWorksheetInfo.DefaultRowHeight != null)
                    await WriteAsync($" ss:DefaultRowHeight=\"{Math.Max(0f, Math.Min(_CurrentWorksheetInfo.DefaultRowHeight.Value, 409.5f)).ToString("G", _Culture)}\"");

                if (_CurrentWorksheetInfo.DefaultColumnWidth != null)
                    await WriteAsync($" ss:DefaultColumnWidth=\"{(_CurrentWorksheetInfo.DefaultColumnWidth.Value * COLUMN_WIDTH_MULTIPLIER).ToString("G", _Culture)}\"");


                await WriteAsync(">");

                foreach (string col in _Columns)
                {
                    if (string.IsNullOrEmpty(col)) await WriteAsync("<Column/>");
                    else await WriteAsync($"<Column ss:Width=\"{col}\"/>");
                }

                _ShouldBeginWorksheet = false;
            }
        }

        private async Task WritePendingEndRowAsync()
        {
            if (!_ShouldEndRow) return;

            await WriteAsync("</Row>");

            _ShouldEndRow = false;
        }

        private async Task WritePendingEndWorksheetAsync()
        {
            if (_ShouldEndWorksheet)
            {
                await WritePendingBeginWorksheetAsync();
                await WritePendingEndRowAsync();

                await WriteAsync("</Table>");
                await WriteAsync("</Worksheet>");

                _ShouldEndWorksheet = false;
            }
        }

        private async Task WriteBeginFileAsync()
        {
            if (_WroteFileStart) return;
            _WroteFileStart = true;

            if (OutputStream != null && ShouldWriteBOM)
            {
                OutputStream.Write(_FileEncoding.GetPreamble(), 0, _FileEncoding.GetPreamble().Length);
            }

            await WriteAsync("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            await WriteAsync("<?mso-application progid=\"Excel.Sheet\"?>");
            await WriteAsync("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"");
            await WriteAsync(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
            await WriteAsync(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"");
            await WriteAsync(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");

            await WriteAsync("<Styles>");
            await WriteAsync("<ss:Style ss:ID=\"Default\" ss:Name=\"Normal\">");
            await WriteAsync("<ss:Alignment ss:Vertical=\"Bottom\"/>");
            await WriteAsync("</ss:Style>");

            foreach (var style in _Styles.Keys)
            {
                await WriteStyleAsync(style, GetStyleId(style, true));
            }

            await WriteAsync("</Styles>");
        }

        #endregion

        #region SpreadsheetWriter - Document Lifespan (public)

        public override async Task NewWorksheetAsync(WorksheetInfo info)
        {
            await WritePendingEndWorksheetAsync();

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

        public override Task SkipRowAsync()
        {
            return SkipRowsAsync(1);
        }

        public override async Task SkipRowsAsync(int count)
        {
            for (int i = 0; i < count; i++)
            {
                await AddRowAsync();
            }
        }
        
        public override async Task AddRowAsync(Style? style = null, float height = 0f, bool autoFit = true)
        {
            if (!_ShouldEndWorksheet)
            {
                throw new InvalidOperationException("Adding new rows is not allowed at this time. Please call NewWorksheetAsync(...) first.");
            }

            if (!_WroteFileStart)
            {
                if (style != null)
                {
                    // A chance to register this style
                    RegisterStyle(style);
                }

                await WriteBeginFileAsync();
            }

            await WritePendingBeginWorksheetAsync();
            await WritePendingEndRowAsync();

            await WriteAsync("<Row");

            if (style != null)
            {
                await WriteAsync(string.Format(" ss:StyleID=\"{0}\"", GetStyleId(style, false)));
            }

            if (height != 0)
            {
                await WriteAsync(string.Format(" ss:Height=\"{0:0.##}\"", height));
            }

            await WriteAsync(">");

            _ShouldEndRow = true;
        }

        public override async Task FinishAsync()
        {
            if (!_WroteFileStart)
            {
                await WriteBeginFileAsync();
                await NewWorksheetAsync(new WorksheetInfo { });
            }

            await WritePendingEndWorksheetAsync();

            if (!_WroteFileEnd)
            {
                await WriteAsync("</Workbook>");
                _WroteFileEnd = true;

                _Writer!.Flush();
                _Writer.Close();
            }
        }

        #endregion
        
        #region SpreadsheetWriter - Cell methods

        public override Task SkipCellAsync()
        {
            return SkipCellsAsync(1);
        }

        public override async Task SkipCellsAsync(int count)
        {
            for (int i = 0; i < count; i++)
            {
                await AddCellAsync("");
            }
        }
        
        public override async Task AddCellAsync(string? data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";

            await WriteAsync(string.Format("<Cell{0}{1}><Data ss:Type=\"String\">{2}</Data></Cell>", styleString, merge, _XmlWriterHelper!.EscapeValue(data)));
        }

        public override async Task AddCellStringAutoTypeAsync(string? data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
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

            await WriteAsync(string.Format("<Cell{0}{1}><Data ss:Type=\"{2}\">{3}</Data></Cell>", styleString, merge, type, _XmlWriterHelper!.EscapeValue(data)));
        }

        public override async Task AddCellForcedStringAsync(string? data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            await AddCellAsync(data, style, horzCellCount, vertCellCount);
        }

        public override async Task AddCellAsync(Int32 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            await WriteAsync(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:G}</Data></Cell>", styleString, merge, data));
        }

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public override async Task AddCellAsync(UInt32 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            await WriteAsync(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:G}</Data></Cell>", styleString, merge, data));
        }
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public override async Task AddCellAsync(Int64 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            await WriteAsync(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:G}</Data></Cell>", styleString, merge, data));
        }

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public override async Task AddCellAsync(UInt64 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            await WriteAsync(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:G}</Data></Cell>", styleString, merge, data));
        }
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public override async Task AddCellAsync(float data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            await WriteAsync(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:R15}</Data></Cell>", styleString, merge, data));
        }

        public override async Task AddCellAsync(double data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            await WriteAsync(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:R15}</Data></Cell>", styleString, merge, data));
        }

        public override async Task AddCellAsync(decimal data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            await WriteAsync(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"Number\">{2:G15}</Data></Cell>", styleString, merge, data));
        }

        public override async Task AddCellAsync(DateTime data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";
            await WriteAsync(string.Format(_Culture, "<Cell{0}{1}><Data ss:Type=\"DateTime\">{2}</Data></Cell>",
                styleString, merge,
                data.Year <= 1 ? "" : data.ToString("yyyy-MM-ddTHH:mm:ss.fff")));
        }

        public override async Task AddCellFormulaAsync(string formula, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            string merge = GetCellMergeString(horzCellCount, vertCellCount);
            string styleString = style != null ? $" ss:StyleID=\"{GetStyleId(style, false)}\"" : "";

            await WriteAsync(string.Format("<Cell{0}{1} ss:Formula=\"{2}\" />",
                styleString, merge, XmlWriterHelper.EscapeSimpleXmlAttr(formula)));
        }
        public override Task AddCellImageAsync(
            Image image,
            Style? style = null,
            int horzCellCount = 0,
            int vertCellCount = 0,
            CancellationToken cancellationToken = default)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
