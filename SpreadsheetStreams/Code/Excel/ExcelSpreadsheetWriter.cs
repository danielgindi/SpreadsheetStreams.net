using SpreadsheetStreams.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SpreadsheetStreams
{
    public class ExcelSpreadsheetWriter : SpreadsheetWriter
    {
        #region Constructors

        public ExcelSpreadsheetWriter(Stream outputStream, CompressionLevel compressionLevel = CompressionLevel.Fastest, bool leaveOpen = false)
            : base(outputStream ?? new MemoryStream())
        {
            if (outputStream.GetType().FullName == "System.Web.HttpResponseStream")
            {
                outputStream = new WriteStreamWrapper(outputStream);
            }

            _Package = new PackageWriteStream(outputStream, leaveOpen);

            _CompressionLevel = compressionLevel;
        }

        public ExcelSpreadsheetWriter() : this(null)
        {
        }

        #endregion

        #region IDisposable

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            if (disposing)
            {
                if (_Package != null)
                {
                    _Package.Dispose();
                    _Package = null;
                }

                if (_CurrentWorksheetPartWriter != null)
                {
                    _CurrentWorksheetPartWriter.Dispose();
                    _CurrentWorksheetPartWriter = null;
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

        private PackageWriteStream _Package = null;
        private CompressionLevel _CompressionLevel = CompressionLevel.Fastest;

        private ZipArchiveEntry _CurrentWorksheetEntry = null;
        private Stream _CurrentWorksheetPartStream = null;
        private StreamWriter _CurrentWorksheetPartWriter = null;
        private WorksheetInfo _CurrentWorksheetInfo = null;
        private FrozenPaneState? _CurrentWorksheePane = null;
        private int _RowCount = 0;
        private int _CellCount = 0;
        private List<string> _MergeCells = new List<string>();

        private List<WorksheetInfo> _WorksheetInfos = new List<WorksheetInfo>();

        private bool _WroteFileStart = false;
        private bool _WroteFileEnd = false;
        private bool _ShouldBeginWorksheet = false;
        private bool _ShouldEndWorksheet = false;
        private bool _ShouldEndRow = false;

        private int _NextStyleIndex = 1;
        private Dictionary<Style, int> _Styles = new Dictionary<Style, int>();
        private Dictionary<int, int> _StyleIdBorderIdMap = new Dictionary<int, int>();
        private Dictionary<int, int> _StyleIdFontIdMap = new Dictionary<int, int>();
        private Dictionary<int, int> _StyleIdFillIdMap = new Dictionary<int, int>();
        private Dictionary<int, int> _StyleIdNumberFormatIdMap = new Dictionary<int, int>();
        private List<string> _Columns = new List<string>();

        private XmlWriterHelper _XmlWriterHelper = new XmlWriterHelper();

        #endregion

        #region SpreadsheetWriter - Basic properties

        public override string FileExtension => "xlsx";
        public override string FileContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public override bool IsFlatFormat => false;

        #endregion

        #region Public Properties

        public const int MIN_COLUMN_NUMBER = 1;
        public const int MAX_COLUMN_NUMBER = 16384;

        #endregion

        #region Private helpers

        private string GetStyleId(Style style)
        {
            if (!_Styles.TryGetValue(style, out var index))
            {
                index = _NextStyleIndex++;
                _Styles[style] = index;
            }

            return index.ToString();
        }

        private void MergeNextCell(int horzCellCount, int vertCellCount)
        {
            if (horzCellCount > 1 || vertCellCount > 1)
            {
                _MergeCells.Add(
                    ConvertColumnAddress(_CellCount + 1) + _RowCount +
                    ":" +
                    ConvertColumnAddress(_CellCount + (horzCellCount > 1 ? horzCellCount : 1)) + (_RowCount + (vertCellCount > 1 ? vertCellCount - 1 : 0)));
            }
        }

        private void FillHorzCellCount(int horzCellCount)
        {
            if (horzCellCount < 2)
                return;

            _CellCount += horzCellCount - 1;
        }

        private static string GetBorderLineStyleName(BorderLineStyle style, float weight)
        {
            switch (style)
            {
                default:
                case BorderLineStyle.None: return "";

                case BorderLineStyle.Continuous:
                    return weight >= 2.0f ? "medium"
                        : weight >= 3.0f ? "thick"
                        : "thin";

                case BorderLineStyle.Dot: return "dotted";
                case BorderLineStyle.DashDot: return weight >= 2.0f ? "mediumDashDot" : "dashDot";
                case BorderLineStyle.DashDotDot: return weight >= 2.0f ? "mediumDashDotDot" : "dashDotDot";
                case BorderLineStyle.Dash: return weight >= 2.0f ? "mediumDashDot" : "dashed";
                case BorderLineStyle.SlantDashDot: return "slantDashDot";
                case BorderLineStyle.Double: return "double";
            }
        }

        private static int GetFontFamilyNumbering(FontFamily style)
        {
            switch (style)
            {
                default:
                case FontFamily.Automatic: return 0;
                case FontFamily.Roman: return 1;
                case FontFamily.Swiss: return 2;
                case FontFamily.Modern: return 3;
                case FontFamily.Script: return 4;
                case FontFamily.Decorative: return 5;
            }
        }

        private static string GetPatternName(FillPattern pattern)
        {
            switch (pattern)
            {
                default:
                case FillPattern.None: return "none";
                case FillPattern.Solid: return "solid";
                case FillPattern.Gray75: return "darkGray";
                case FillPattern.Gray50: return "mediumGray";
                case FillPattern.Gray25: return "lightGray";
                case FillPattern.Gray125: return "gray125";
                case FillPattern.Gray0625: return "gray0625";
                case FillPattern.HorzStripe: return "darkHorizontal";
                case FillPattern.VertStripe: return "darkVertical";
                case FillPattern.ReverseDiagStripe: return "darkDown";
                case FillPattern.DiagStripe: return "darkUp";
                case FillPattern.DiagCross: return "darkGrid";
                case FillPattern.ThickDiagCross: return "darkTrellis";
                case FillPattern.ThinHorzStripe: return "lightHorizontal";
                case FillPattern.ThinVertStripe: return "lightVertical";
                case FillPattern.ThinReverseDiagStripe: return "lightDown";
                case FillPattern.ThinDiagStripe: return "lightUp";
                case FillPattern.ThinHorzCross: return "lightGrid";
                case FillPattern.ThinDiagCross: return "lightTrellis";
            }
        }

        private (int, string) ConvertNumberFormat(NumberFormat format)
        {
            switch (format.Type)
            {
                default:
                case NumberFormatType.General: return (0, null);
                case NumberFormatType.None: return (-1, null);
                case NumberFormatType.Custom: return (-2, format.Custom);
                case NumberFormatType.GeneralNumber: return (0, null);
                case NumberFormatType.GeneralDate: return (22, null);
                case NumberFormatType.ShortDate: return (13, null);
                case NumberFormatType.MediumDate: return (14, null);
                case NumberFormatType.LongDate: return (-2, "[$]dddd, mmmm d, yyyy;@");
                case NumberFormatType.ShortTime: return (19, null);
                case NumberFormatType.MediumTime: return (17, null);
                case NumberFormatType.LongTime: return (18, null);
                case NumberFormatType.Fixed: return (1, null);
                case NumberFormatType.Standard: return (3, null);
                case NumberFormatType.Percent: return (10, null);
                case NumberFormatType.Scientific: return (47, null);
                case NumberFormatType.YesNo: return (-2, "\"Yes\";\"Yes\";\"No\"");
                case NumberFormatType.TrueFalse: return (-2, "\"True\";\"True\";\"False\"");
                case NumberFormatType.OnOff: return (-2, "\"On\";\"On\";\"Off\"");
            }
        }

        private int GetBorderPositionOrdering(BorderPosition pos)
        {
            switch (pos)
            {
                default: return -1;
                case BorderPosition.Left: return 0;
                case BorderPosition.Right: return 1;
                case BorderPosition.Top: return 2;
                case BorderPosition.Bottom: return 3;
                case BorderPosition.DiagonalRight: return 4;
                case BorderPosition.DiagonalLeft: return 4;
            }
        }

        private bool WriteStyleBorderXml(StreamWriter writer, List<Border> items, bool writeEmpty, int id, int? styleId)
        {
            if (items == null)
            {
                if (writeEmpty)
                    writer.Write($"<border><left/><right/><top/><bottom/><diagonal/></border>");

                if (styleId != null) _StyleIdBorderIdMap[styleId.Value] = -1;

                return writeEmpty;
            }
            else
            {
                var hasDiagonalUp = items.Any(x => x.Position == BorderPosition.DiagonalRight);
                var hasDiagonalDown = items.Any(x => x.Position == BorderPosition.DiagonalLeft);

                writer.Write($"<border{(hasDiagonalUp ? "diagonalUp=\"1\"" : "")}{(hasDiagonalDown ? "diagonalDown =\"1\"" : "")}>");

                foreach (var item in items
                    .Distinct()
                    .Where(x => !items.Any(y => y != x && GetBorderPositionOrdering(y.Position) == GetBorderPositionOrdering(x.Position)))
                    .OrderBy(x => GetBorderPositionOrdering(x.Position))
                    )
                {
                    string position;

                    switch (item.Position)
                    {
                        default:
                        case BorderPosition.Left: position = "left"; break;
                        case BorderPosition.Right: position = "right"; break;
                        case BorderPosition.Top: position = "top"; break;
                        case BorderPosition.Bottom: position = "bottom"; break;
                        case BorderPosition.DiagonalRight:
                        case BorderPosition.DiagonalLeft: position = "diagonal"; break;
                    }

                    writer.Write($"<{position} style=\"{GetBorderLineStyleName(item.LineStyle, item.Weight)}\">");
                    {
                        if (item.Color != System.Drawing.Color.Empty)
                            writer.Write($"<color rgb=\"{ColorHelper.GetHexArgb(item.Color)}\"/>");
                        else writer.Write("<color auto=\"1\"/>");
                    }
                    writer.Write($"</{position}>");
                }

                writer.Write("</border>");

                if (styleId != null) _StyleIdBorderIdMap[styleId.Value] = id;

                return true;
            }
        }

        private bool WriteStyleFontXml(StreamWriter writer, Font? font, bool writeEmpty, int id, int? styleId)
        {
            if (font == null)
            {
                if (writeEmpty)
                    writer.Write($"<font></font>");

                if (styleId != null) _StyleIdFontIdMap[styleId.Value] = -1;

                return writeEmpty;
            }
            else
            {
                var item = font.Value;

                writer.Write("<font>");
                {
                    if (item.Bold) writer.Write("<b/>");
                    if (item.Italic) writer.Write("<i/>");

                    if (item.Underline == FontUnderline.Single) writer.Write("<u/>");
                    if (item.Underline == FontUnderline.SingleAccounting) writer.Write("<u val=\"singleAccounting\"/>");
                    if (item.Underline == FontUnderline.Double) writer.Write("<u val=\"double\"/>");
                    if (item.Underline == FontUnderline.DoubleAccounting) writer.Write("<u val=\"doubleAccounting\"/>");

                    if (item.StrikeThrough) writer.Write("<strike/>");

                    if (item.VerticalAlign == FontVerticalAlign.Subscript) writer.Write("<vertAlign val=\"subscript\"/>");
                    if (item.VerticalAlign == FontVerticalAlign.Superscript) writer.Write("<vertAlign val=\"superscript\"/>");

                    writer.Write($"<sz val=\"{item.Size.ToString("G", _Culture)}\"/>");

                    if (item.Color != Color.Empty)
                        writer.Write($"<color rgb=\"{ColorHelper.GetHexArgb(item.Color)}\"/>");
                    else writer.Write("<color auto=\"1\"/>");

                    writer.Write($"<name val=\"{item.Name}\"/>");
                    writer.Write($"<family val=\"{GetFontFamilyNumbering(item.Family)}\"/>");

                    if (item.Charset != null)
                    {
                        writer.Write($"<charset val=\"{(int)item.Charset.Value}\"/>");
                    }
                }
                writer.Write("</font>");

                if (styleId != null) _StyleIdFontIdMap[styleId.Value] = id;

                return true;
            }
        }

        private bool WriteStyleFill(StreamWriter writer, Fill? fill, bool writeEmpty, int id, int? styleId)
        {
            if (fill == null || fill?.Pattern == FillPattern.None)
            {
                if (writeEmpty)
                    writer.Write($"<fill><patternFill patternType=\"none\"/></fill>");

                if (styleId != null) _StyleIdFillIdMap[styleId.Value] = -1;

                return writeEmpty;
            }
            else
            {
                var item = fill.Value;

                writer.Write("<fill>");
                {
                    writer.Write($"<patternFill patternType=\"{GetPatternName(item.Pattern)}\">");

                    var fgColor = item.PatternColor;
                    if (item.Pattern == FillPattern.Solid && fgColor == Color.Empty)
                        fgColor = item.Color;

                    if (fgColor != Color.Empty)
                        writer.Write($"<fgColor rgb=\"{ColorHelper.GetHexArgb(fgColor)}\"/>");
                    else writer.Write("<fgColor auto=\"1\"/>");

                    if (item.Pattern != FillPattern.Solid)
                    {
                        if (item.Color != Color.Empty)
                            writer.Write($"<bgColor rgb=\"{ColorHelper.GetHexArgb(item.Color)}\"/>");
                        else writer.Write("<bgColor auto=\"1\"/>");
                    }

                    writer.Write("</patternFill>");
                }
                writer.Write("</fill>");

                if (styleId != null) _StyleIdFillIdMap[styleId.Value] = id;

                return true;
            }
        }

        private bool WriteStyleNumberFormatXml(StreamWriter writer, NumberFormat format, bool writeEmpty, int id, int? styleId)
        {
            var convertedFormat = ConvertNumberFormat(format);

            if (convertedFormat.Item1 == -2)
            {
                writer.Write($"<numFmt numFmtId=\"{id}\" formatCode=\"{_XmlWriterHelper.EscapeAttribute(convertedFormat.Item2)}\"/>");
                if (styleId != null) _StyleIdNumberFormatIdMap[styleId.Value] = id;
                return true;
            }
            else if (convertedFormat.Item1 == -1)
            {
                if (writeEmpty)
                    writer.Write($"<numFmt numFmtId=\"0\"/>");

                if (styleId != null) _StyleIdNumberFormatIdMap[styleId.Value] = -1;
                return false;
            }
            else
            {
                if (styleId != null) _StyleIdNumberFormatIdMap[styleId.Value] = convertedFormat.Item1;
                return false;
            }
        }

        private void WriteStyleXfXml(StreamWriter writer, int styleId, Style style)
        {
            var numFmtId = _StyleIdNumberFormatIdMap[styleId];
            var borderId = _StyleIdBorderIdMap[styleId];
            var fillId = _StyleIdFillIdMap[styleId];
            var fontId = _StyleIdFontIdMap[styleId];

            writer.Write($"<xf");
            writer.Write($" numFmtId=\"{(numFmtId == -1 ? 0 : numFmtId)}\"");
            writer.Write($" borderId=\"{(borderId == -1 ? 0 : borderId)}\"");
            writer.Write($" fillId=\"{(fillId == -1 ? 0 : fillId)}\"");
            writer.Write($" fontId=\"{(fontId == -1 ? 0 : fontId)}\"");

            if (numFmtId > -1)
                writer.Write(" applyNumberFormat=\"1\"");
            if (borderId > -1)
                writer.Write(" applyBorder=\"1\"");
            if (fillId > -1)
                writer.Write(" applyFill=\"1\"");
            if (fontId > -1)
                writer.Write(" applyFont=\"1\"");
            if (style.Alignment != null)
                writer.Write(" applyAlignment=\"1\"");

            if (style.Alignment != null)
            {
                var align = style.Alignment.Value;

                writer.Write(">");
                {
                    writer.Write("<alignment");

                    if (align.VerticalText == true)
                    {
                        writer.Write(" textRotation=\"255\"");
                    }
                    else if (align.Rotate != 0d)
                    {
                        writer.Write($" textRotation=\"{Math.Max(-90d, Math.Min(90d, align.Rotate)).ToString("G", _Culture)}\"");
                    }

                    string horz = null;
                    switch (align.Horizontal)
                    {
                        default:
                        case HorizontalAlignment.Automatic: break;
                        case HorizontalAlignment.Center: horz = "center"; break;
                        case HorizontalAlignment.CenterAcrossSelection: horz = "centerContinuous"; break;
                        case HorizontalAlignment.Distributed: horz = "distributed"; break;
                        case HorizontalAlignment.Fill: horz = "fill"; break;
                        case HorizontalAlignment.Justify: horz = "justify"; break;
                        case HorizontalAlignment.Left: horz = "left"; break;
                        case HorizontalAlignment.Right: horz = "right"; break;
                    }

                    if (horz != null)
                        writer.Write($" horizontal=\"{horz}\"");

                    string vert = null;
                    switch (align.Vertical)
                    {
                        default:
                        case VerticalAlignment.Automatic: break;
                        case VerticalAlignment.Bottom: vert = "bottom"; break;
                        case VerticalAlignment.Center: vert = "center"; break;
                        case VerticalAlignment.Distributed: vert = "distributed"; break;
                        case VerticalAlignment.Justify: vert = "justify"; break;
                        case VerticalAlignment.Top: vert = "top"; break;
                    }

                    if (vert != null)
                        writer.Write($" vertical=\"{vert}\"");

                    if (align.WrapText)
                        writer.Write($" wrapText=\"1\"");

                    if (align.ShrinkToFit)
                        writer.Write($" shrinkToFit=\"1\"");

                    if (align.ReadingOrder == HorizontalReadingOrder.LeftToRight)
                        writer.Write($" readingOrder=\"1\"");
                    else if (align.ReadingOrder == HorizontalReadingOrder.RightToLeft)
                        writer.Write($" readingOrder=\"2\"");

                    if (align.Indent > 0)
                        writer.Write($" indent=\"{align.Indent}\"");

                    writer.Write("/>");
                }
                writer.Write("</xf>");
            }
            else
            {
                writer.Write("/>");
            }
        }

        private void WriteStylesXml(Stream stream)
        {
            using (var writer = new StreamWriter(stream))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
                writer.Write("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
                {
                    var orderedStyles = _Styles.OrderBy(p => p.Value);

                    var nextId = 124;
                    var customFormatCount = orderedStyles.Count(x => x.Key.NumberFormat.Type == NumberFormatType.Custom);
                    writer.Write($"<numFmts count=\"{customFormatCount}\">");
                    foreach (var pair in orderedStyles)
                    {
                        if (WriteStyleNumberFormatXml(writer, pair.Key.NumberFormat, false, nextId, pair.Value))
                            nextId++;
                    }
                    writer.Write("</numFmts>");

                    nextId = 1;
                    writer.Write($"<fonts x14ac:knownFonts=\"1\" count=\"{_Styles.Count(x => x.Key.Font != null) + 1}\">");
                    WriteStyleFontXml(writer, null, true, 0, null);
                    foreach (var pair in orderedStyles)
                    {
                        if (WriteStyleFontXml(writer, pair.Key.Font, false, nextId, pair.Value))
                            nextId++;
                    }
                    writer.Write("</fonts>");

                    nextId = 2; // We need 2 dummies, for some reason
                    writer.Write($"<fills count=\"{_Styles.Count(x => x.Key.Fill != null && x.Key.Fill.Value.Pattern != FillPattern.None) + 2}\">");
                    WriteStyleFill(writer, null, true, 0, null);
                    WriteStyleFill(writer, null, true, 0, null);
                    foreach (var pair in orderedStyles)
                    {
                        if (WriteStyleFill(writer, pair.Key.Fill, false, nextId, pair.Value))
                            nextId++;
                    }
                    writer.Write("</fills>");

                    nextId = 1;
                    writer.Write($"<borders count=\"{_Styles.Count(x => x.Key.Borders != null) + 1}\">");
                    WriteStyleBorderXml(writer, null, true, 0, null);
                    foreach (var pair in orderedStyles)
                    {
                        if (WriteStyleBorderXml(writer, pair.Key.Borders, false, nextId, pair.Value))
                            nextId++;
                    }
                    writer.Write("</borders>");

                    writer.Write($"<cellXfs count=\"{_Styles.Count + 1}\">");
                    writer.Write($"<xf numFmtId=\"0\" borderId=\"0\" fillId=\"0\" fontId=\"0\"/>");
        
                    foreach (var pair in orderedStyles)
                    {
                        WriteStyleXfXml(writer, pair.Value, pair.Key);
                    }
                    writer.Write("</cellXfs>");

                }
                writer.Write("</styleSheet>");
            }
        }

        private void WriteWorkbookXml(Stream stream)
        {
            using (var writer = new StreamWriter(stream))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
                writer.Write("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

                writer.Write("<sheets>");

                foreach (var item in _WorksheetInfos)
                {
                    var name = item.Name ?? $"Worksheet{item.Id}";

                    // Remove invalid characters
                    name = Regex.Replace(name, "[:/\\?[\\]]", " ");

                    // Cannot begin/end with an apostrophe
                    name = Regex.Replace(name, "^'+|'+$", "");

                    // Excel limits to 31 characters. Otherwise it's an error
                    name = name.Substring(0, Math.Min(name.Length, 31));

                    writer.Write($"<sheet r:id=\"rId{item.Id}\" sheetId=\"{item.Id}\" name=\"{_XmlWriterHelper.EscapeAttribute(name)}\"/>");
                }

                writer.Write("</sheets>");
                writer.Write("</workbook>");
            }
        }

        /// <summary>
        /// Gets the column address (A - XFD)
        /// </summary>
        /// <param name="columnNumber">Column number (one-based)</param>
        /// <returns>Column address (A - XFD)</returns>
        /// <exception cref="RangeException">Throws an RangeException if the passed column number was out of range</exception>
        public static string ConvertColumnAddress(int columnNumber)
        {
            if (columnNumber > MAX_COLUMN_NUMBER || columnNumber < MIN_COLUMN_NUMBER)
            {
                throw new Exception($"The column number ({columnNumber}) is out of range. Range is from {MIN_COLUMN_NUMBER} to {MAX_COLUMN_NUMBER}.");

            }

            int a = 0, b = 0, c = 0;
            var address = "";
            for (int i = 0; i < columnNumber; i++)
            {
                if (a > 25)
                {
                    b++;
                    a = 0;
                }
                if (b > 25)
                {
                    c++;
                    b = 0;
                }
                a++;
            }

            if (c > 0) { address += ((char)(c + 64)); }
            if (b > 0) { address += ((char)(b + 64)); }
            address += ((char)(a + 64));

            return address;
        }

        #endregion

        #region Syling

        public override void RegisterStyle(Style style)
        {
            GetStyleId(style);
        }

        #endregion

        #region SpreadsheetWriter - Document Lifespan (private)

        private void BeginPackage()
        {
        }

        private void EndPackage()
        {
            string workbookPath = "/xl/workbook.xml";
            string stylesheetPath = "/xl/styles.xml";
            string sharedStringsPath = "/xl/sharedStrings.xml";
            string docPropsCorePath = "/docProps/core.xml";
            string docPropsAppPath = "/docProps/app.xml";

            var idCounter = _WorksheetInfos.Count + 1;

            int ridWb = idCounter++;
            int ridStyles = idCounter++;
            int ridSharedStrings = idCounter++;
            int ridDocPropsCore = idCounter++;
            int ridDocPropsApp = idCounter++;

            var wbEntry = _Package.CreateStream(workbookPath, _CompressionLevel);
            using (var stream = wbEntry.Open())
            {
                _Package.AddPackageRelationship(workbookPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "rId" + ridWb);
                _Package.AddContentType(workbookPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");

                WriteWorkbookXml(stream);
            }

            var stylesEntry = _Package.CreateStream(stylesheetPath, _CompressionLevel);
            using (var stream = stylesEntry.Open())
            {
                _Package.AddPartRelationship(workbookPath, stylesheetPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "rId" + ridStyles);
                _Package.AddContentType(stylesheetPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");

                WriteStylesXml(stream);
            }

            var sharedStringsEntry = _Package.CreateStream(sharedStringsPath, _CompressionLevel);
            _Package.AddPartRelationship(workbookPath, sharedStringsPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "rId" + ridSharedStrings);
            _Package.AddContentType(sharedStringsPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");

            using (var stream = sharedStringsEntry.Open())
            using (var writer = new StreamWriter(stream))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
                writer.Write("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>");
            }

            foreach (var ws in _WorksheetInfos)
            {
                _Package.AddPartRelationship(workbookPath, ws.Path, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "rId" + ws.Id);
                _Package.AddContentType(ws.Path, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            }

            if (SpreadsheetInfo != null)
            {
                var docPropsCoreEntry = _Package.CreateStream(docPropsCorePath, _CompressionLevel);
                _Package.AddPackageRelationship(docPropsCorePath, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "rId" + ridDocPropsCore);
                _Package.AddContentType(docPropsCorePath, "application/vnd.openxmlformats-package.core-properties+xml");

                using (var stream = docPropsCoreEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
                    writer.Write("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"" + 
                        " xmlns:dc=\"http://purl.org/dc/elements/1.1/\"" +
                        " xmlns:dcterms=\"http://purl.org/dc/terms/\"" + 
                        " xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\"" +
                        " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");

                    if (SpreadsheetInfo.Title != null)
                    {
                        writer.Write($"<dc:title>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.Title)}</dc:title>");
                    }

                    if (SpreadsheetInfo.Subject != null)
                    {
                        writer.Write($"<dc:subject>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.Subject)}</dc:subject>");
                    }

                    if (SpreadsheetInfo.Author != null)
                    {
                        writer.Write($"<dc:creator>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.Author)}</dc:creator>");
                    }

                    if (SpreadsheetInfo.Keywords != null)
                    {
                        writer.Write($"<cp:keywords>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.Keywords)}</cp:keywords>");
                    }

                    if (SpreadsheetInfo.Comments != null)
                    {
                        writer.Write($"<dc:description>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.Comments)}</dc:description>");
                    }

                    if (SpreadsheetInfo.Status != null)
                    {
                        writer.Write($"<cp:contentStatus>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.Status)}</cp:contentStatus>");
                    }

                    if (SpreadsheetInfo.Category != null)
                    {
                        writer.Write($"<cp:category>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.Category)}</cp:category>");
                    }

                    if (SpreadsheetInfo.LastModifiedBy != null)
                    {
                        writer.Write($"<cp:lastModifiedBy>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.LastModifiedBy)}</cp:lastModifiedBy>");
                    }

                    if (SpreadsheetInfo.CreatedOn != null)
                    {
                        writer.Write($"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{SpreadsheetInfo.CreatedOn.Value.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")}</dcterms:created>");
                    }

                    if (SpreadsheetInfo.ModifiedOn != null)
                    {
                        writer.Write($"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{SpreadsheetInfo.ModifiedOn.Value.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")}</dcterms:modified>");
                    }

                    writer.Write("</cp:coreProperties>");
                }

                var docPropsAppEntry = _Package.CreateStream(docPropsAppPath, _CompressionLevel);
                _Package.AddPackageRelationship(docPropsAppPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "rId" + ridDocPropsApp);
                _Package.AddContentType(docPropsAppPath, "application/vnd.openxmlformats-officedocument.extended-properties+xml");

                using (var stream = docPropsAppEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
                    writer.Write("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"" +
                        " xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");

                    if (SpreadsheetInfo.Application != null)
                    {
                        writer.Write($"<Application>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.Application)}</Application>");
                    }

                    if (SpreadsheetInfo.ScaleCrop != null)
                    {
                        writer.Write($"<ScaleCrop>{(SpreadsheetInfo.ScaleCrop.Value ? "true" : "false")}</ScaleCrop>");
                    }

                    if (SpreadsheetInfo.Manager != null)
                    {
                        writer.Write($"<Manager>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.Manager)}</Manager>");
                    }

                    if (SpreadsheetInfo.Company != null)
                    {
                        writer.Write($"<Company>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.Company)}</Company>");
                    }

                    if (SpreadsheetInfo.LinksUpToDate != null)
                    {
                        writer.Write($"<LinksUpToDate>{(SpreadsheetInfo.LinksUpToDate.Value ? "true" : "false")}</LinksUpToDate>");
                    }

                    if (SpreadsheetInfo.SharedDoc != null)
                    {
                        writer.Write($"<SharedDoc>{(SpreadsheetInfo.SharedDoc.Value ? "true" : "false")}</SharedDoc>");
                    }

                    if (SpreadsheetInfo.HyperlinksChanged != null)
                    {
                        writer.Write($"<HyperlinksChanged>{(SpreadsheetInfo.HyperlinksChanged.Value ? "true" : "false")}</HyperlinksChanged>");
                    }

                    if (SpreadsheetInfo.AppVersion != null)
                    {
                        writer.Write($"<AppVersion>{_XmlWriterHelper.EscapeValue(SpreadsheetInfo.AppVersion)}</AppVersion>");
                    }

                    writer.Write("</Properties>");
                }
            }

            _Package.CommitRelationships(_CompressionLevel);
            _Package.CommitContentTypes(_CompressionLevel);
            _Package.Close();

            _Package = null;
        }

        private void WritePendingBeginWorksheet()
        {
            if (_ShouldBeginWorksheet)
            {
                if (_CurrentWorksheetPartWriter != null)
                {
                    _CurrentWorksheetPartWriter.Dispose();
                }

                _CurrentWorksheetEntry = _Package.CreateStream(_CurrentWorksheetInfo.Path, _CompressionLevel);
                _CurrentWorksheetPartStream = _CurrentWorksheetEntry.Open();
                _CurrentWorksheetPartWriter = new StreamWriter(_CurrentWorksheetPartStream, Encoding.UTF8);

                _CurrentWorksheetPartWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
                _CurrentWorksheetPartWriter.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
                {
                    _CurrentWorksheetPartWriter.Write("<sheetViews><sheetView");
                    if (_CurrentWorksheetInfo.RightToLeft != null)
                    {
                        _CurrentWorksheetPartWriter.Write($" rightToLeft=\"{(_CurrentWorksheetInfo.RightToLeft == true ? "1" : "0")}\"");
                    }
                    _CurrentWorksheetPartWriter.Write(" workbookViewId=\"0\">");

                    if (_CurrentWorksheePane != null &&
                        (_CurrentWorksheePane.Value.Column > 1 || _CurrentWorksheePane.Value.Row > 1))
                    {
                        _CurrentWorksheetPartWriter.Write(@"<pane");

                        if (_CurrentWorksheePane.Value.Column > 1)
                            _CurrentWorksheetPartWriter.Write($@" xSplit=""{_CurrentWorksheePane.Value.Column - 1}""");

                        if (_CurrentWorksheePane.Value.Row > 1)
                            _CurrentWorksheetPartWriter.Write($@" ySplit=""{_CurrentWorksheePane.Value.Row - 1}""");

                        _CurrentWorksheetPartWriter.Write($@" topLeftCell=""{ConvertColumnAddress(_CurrentWorksheePane.Value.Column)}{_CurrentWorksheePane.Value.Row}""");
                        _CurrentWorksheetPartWriter.Write($@" activePane=""bottomRight""");
                        _CurrentWorksheetPartWriter.Write(@" state=""frozen""/>");
                    }

                    _CurrentWorksheetPartWriter.Write("</sheetView></sheetViews>");

                    _CurrentWorksheetPartWriter.Write("<sheetFormatPr");

                    if (_CurrentWorksheetInfo.DefaultRowHeight != null)
                        _CurrentWorksheetPartWriter.Write(string.Format(
                            " defaultRowHeight=\"{0:G}\"",
                            Math.Max(0f, Math.Min(_CurrentWorksheetInfo.DefaultRowHeight.Value, 409.5f))));

                    if (_CurrentWorksheetInfo.DefaultColumnWidth != null)
                        _CurrentWorksheetPartWriter.Write(string.Format(" baseColWidth=\"{0:G}\"", _CurrentWorksheetInfo.DefaultColumnWidth));

                    _CurrentWorksheetPartWriter.Write("/>");

                    var sb = new StringBuilder();
                    if (_CurrentWorksheetInfo.ColumnWidths != null)
                    {
                        for (int i = 0; i < _CurrentWorksheetInfo.ColumnWidths.Length; i++)
                        {
                            var w = _CurrentWorksheetInfo.ColumnWidths[i];
                            if (w == 0f || w == _CurrentWorksheetInfo.DefaultColumnWidth) continue;

                            sb.Append(string.Format(
                                "<col min=\"{0}\" max=\"{0}\" width=\"{1:G}\" bestFit=\"1\" customWidth=\"1\"/>",
                                i + 1,
                                w
                                ));
                        }
                    }

                    if (sb.Length > 0)
                    {
                        _CurrentWorksheetPartWriter.Write("<cols>");
                        _CurrentWorksheetPartWriter.Write(sb.ToString());
                        _CurrentWorksheetPartWriter.Write("</cols>");
                    }

                    _CurrentWorksheetPartWriter.Write("<sheetData>");
                }

                _ShouldBeginWorksheet = false;
            }
        }

        private void WritePendingEndRow()
        {
            if (!_ShouldEndRow) return;

            _CurrentWorksheetPartWriter.Write("</row>");

            _ShouldEndRow = false;
        }

        private void WritePendingEndWorksheet()
        {
            if (_ShouldEndWorksheet)
            {
                WritePendingBeginWorksheet();
                WritePendingEndRow();

                _CurrentWorksheetPartWriter.Write("</sheetData>");

                if (_MergeCells.Count > 0)
                {
                    _CurrentWorksheetPartWriter.Write($"<mergeCells count=\"{_MergeCells.Count}\">");

                    foreach (var merge in _MergeCells)
                    {
                        _CurrentWorksheetPartWriter.Write($"<mergeCell ref=\"{merge}\" />");
                    }

                    _CurrentWorksheetPartWriter.Write("</mergeCells>");
                }

                _CurrentWorksheetPartWriter.Write("</worksheet>");

                if (_CurrentWorksheetPartWriter != null)
                {
                    _CurrentWorksheetPartWriter.Dispose();
                    _CurrentWorksheetPartWriter = null;
                }

                _ShouldEndWorksheet = false;
            }
        }

        private void WriteBeginFile()
        {
            if (_WroteFileStart) return;
            _WroteFileStart = true;
        }

        #endregion

        #region SpreadsheetWriter - Document Lifespan (public)

        public override void NewWorksheet(WorksheetInfo info)
        {
            WritePendingEndWorksheet();

            if (_WorksheetInfos.Contains(info))
                throw new InvalidOperationException("This WorksheetInfo object has already been added to the workbook");

            _CurrentWorksheetInfo = info;
            _CurrentWorksheetInfo.Id = _WorksheetInfos.Count > 0 ? _WorksheetInfos.Last().Id + 1 : 1;
            _CurrentWorksheetInfo.Path = "/xl/worksheets/sheet" + _CurrentWorksheetInfo.Id + ".xml";
            _WorksheetInfos.Add(_CurrentWorksheetInfo);

            _CurrentWorksheePane = null;

            _ShouldBeginWorksheet = true;
            _ShouldEndWorksheet = true;
            _MergeCells.Clear();

            _RowCount = 0;
        }

        public void SetWorksheetFrozenPane(FrozenPaneState? pane)
        {
            _CurrentWorksheePane = pane;
        }

        public override void AddRow(Style style = null, float height = 0f, bool autoFit = true)
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

            _RowCount++;
            _CellCount = 0;
            _CurrentWorksheetPartWriter.Write($"<row r=\"{_RowCount}\"");

            if (height != 0f && height != _CurrentWorksheetInfo.DefaultRowHeight)
            {
                if (!autoFit)
                    _CurrentWorksheetPartWriter.Write(" customHeight=\"1\"");

                _CurrentWorksheetPartWriter.Write($" ht=\"{height.ToString("G", _Culture)}\"");
            }
            else if (_CurrentWorksheetInfo.DefaultRowHeight != null)
            {
                if (!autoFit)
                    _CurrentWorksheetPartWriter.Write(" customHeight=\"1\"");
            }

            if (style != null)
                _CurrentWorksheetPartWriter.Write($" s=\"{GetStyleId(style)}\"");

            _CurrentWorksheetPartWriter.Write(">");

            _ShouldEndRow = true;
        }

        private void WriteCellHeader(int cellIndex, int rowIndex, bool closed, string type, Style style = null)
        {
            _CurrentWorksheetPartWriter.Write($"<c r=\"{ConvertColumnAddress(cellIndex)}{rowIndex}\"");

            if (type != null)
                _CurrentWorksheetPartWriter.Write(" t=\"" + type + "\"");

            if (style != null)
                _CurrentWorksheetPartWriter.Write($" s=\"{GetStyleId(style)}\"");

            _CurrentWorksheetPartWriter.Write(closed ? "/>" : ">");
        }

        private void WriteCellFooter()
        {
            _CurrentWorksheetPartWriter.Write("</c>");
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
                EndPackage();
                _WroteFileEnd = true;
            }
        }

        #endregion

        #region SpreadsheetWriter - Cell methods

        public override void AddCell(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            MergeNextCell(horzCellCount, vertCellCount);

            WriteCellHeader(_CellCount++ + 1, _RowCount, false, "str", style);
            _CurrentWorksheetPartWriter.Write("<v>" + _XmlWriterHelper.EscapeValue(data) + "</v>");
            WriteCellFooter();

            if (horzCellCount > 1)
                FillHorzCellCount(horzCellCount);
        }

        public override void AddCellStringAutoType(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            MergeNextCell(horzCellCount, vertCellCount);

            var type = "str";

            if (style?.NumberFormat != null && (
                style.NumberFormat.Type == NumberFormatType.GeneralNumber ||
                style.NumberFormat.Type == NumberFormatType.Scientific ||
                style.NumberFormat.Type == NumberFormatType.Fixed ||
                style.NumberFormat.Type == NumberFormatType.Standard))
            {
                type = "n";
            }

            WriteCellHeader(_CellCount++ + 1, _RowCount, false, type, style);
            _CurrentWorksheetPartWriter.Write("<v>" + _XmlWriterHelper.EscapeValue(data) + "</v>");
            WriteCellFooter();

            if (horzCellCount > 1)
                FillHorzCellCount(horzCellCount);
        }

        public override void AddCellForcedString(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            AddCell(data, style, horzCellCount, vertCellCount);
        }

        public override void AddCell(Int32 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            MergeNextCell(horzCellCount, vertCellCount);

            WriteCellHeader(_CellCount++ + 1, _RowCount, false, "n", style);
            _CurrentWorksheetPartWriter.Write(string.Format(_Culture, "<v>{0:G}</v>", data));
            WriteCellFooter();

            if (horzCellCount > 1)
                FillHorzCellCount(horzCellCount);
        }

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public override void AddCell(UInt32 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            MergeNextCell(horzCellCount, vertCellCount);

            WriteCellHeader(_CellCount++ + 1, _RowCount, false, "n", style);
            _CurrentWorksheetPartWriter.Write(string.Format(_Culture, "<v>{0:G}</v>", data));
            WriteCellFooter();

            if (horzCellCount > 1)
                FillHorzCellCount(horzCellCount);
        }
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public override void AddCell(Int64 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            MergeNextCell(horzCellCount, vertCellCount);

            WriteCellHeader(_CellCount++ + 1, _RowCount, false, "n", style);
            _CurrentWorksheetPartWriter.Write(string.Format(_Culture, "<v>{0:G}</v>", data));
            WriteCellFooter();

            if (horzCellCount > 1)
                FillHorzCellCount(horzCellCount);
        }

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public override void AddCell(UInt64 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            MergeNextCell(horzCellCount, vertCellCount);

            WriteCellHeader(_CellCount++ + 1, _RowCount, false, "n", style);
            _CurrentWorksheetPartWriter.Write(string.Format(_Culture, "<v>{0:G}</v>", data));
            WriteCellFooter();

            if (horzCellCount > 1)
                FillHorzCellCount(horzCellCount);
        }
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public override void AddCell(float data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            MergeNextCell(horzCellCount, vertCellCount);

            WriteCellHeader(_CellCount++ + 1, _RowCount, false, "n", style);
            _CurrentWorksheetPartWriter.Write(string.Format(_Culture, "<v>{0:G}</v>", data));
            WriteCellFooter();

            if (horzCellCount > 1)
                FillHorzCellCount(horzCellCount);
        }

        public override void AddCell(double data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            MergeNextCell(horzCellCount, vertCellCount);

            WriteCellHeader(_CellCount++ + 1, _RowCount, false, "n", style);
            _CurrentWorksheetPartWriter.Write(string.Format(_Culture, "<v>{0:G}</v>", data));
            WriteCellFooter();

            if (horzCellCount > 1)
                FillHorzCellCount(horzCellCount);
        }

        public override void AddCell(decimal data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            MergeNextCell(horzCellCount, vertCellCount);

            WriteCellHeader(_CellCount++ + 1, _RowCount, false, "n", style);
            _CurrentWorksheetPartWriter.Write(string.Format(_Culture, "<v>{0:G}</v>", data));
            WriteCellFooter();

            if (horzCellCount > 1)
                FillHorzCellCount(horzCellCount);
        }

        public override void AddCell(DateTime data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            MergeNextCell(horzCellCount, vertCellCount);

            if (data.Year >= 1900)
            {
                WriteCellHeader(_CellCount++ + 1, _RowCount, false, "n", style);
                _CurrentWorksheetPartWriter.Write(string.Format(_Culture, "<v>{0:G}</v>", data.ToOADate()));
                WriteCellFooter();
            }
            else
            {
                WriteCellHeader(_CellCount++ + 1, _RowCount, true, null, style);
            }

            if (horzCellCount > 1)
                FillHorzCellCount(horzCellCount);
        }

        public override void AddCellFormula(string formula, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            MergeNextCell(horzCellCount, vertCellCount);

            WriteCellHeader(_CellCount++ + 1, _RowCount, false, "str", style);
            _CurrentWorksheetPartWriter.Write("<f>" + _XmlWriterHelper.EscapeValue(formula) + "</f>");
            WriteCellFooter();

            if (horzCellCount > 1)
                FillHorzCellCount(horzCellCount);
        }

        #endregion
    }
}
