using SpreadsheetStreams.Code.Excel;
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
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

#nullable enable

namespace SpreadsheetStreams
{
    public class ExcelSpreadsheetWriter : SpreadsheetWriter
    {
        #region Constructors

        public ExcelSpreadsheetWriter(Stream outputStream, CompressionLevel compressionLevel = CompressionLevel.Fastest, bool leaveOpen = false)
            : base(outputStream)
        {
            if (outputStream.GetType().FullName == "System.Web.HttpResponseStream")
            {
                outputStream = new WriteStreamWrapper(outputStream);
            }

            _Package = new PackageWriteStream(outputStream, leaveOpen);

            _CompressionLevel = compressionLevel;
        }

        public ExcelSpreadsheetWriter() : this(new MemoryStream())
        {
        }

        #endregion

        #region IDisposable

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            if (disposing)
            {
                // owned by package, dispose first
                if (_CurrentWorksheetPartWriter != null)
                {
                    _CurrentWorksheetPartWriter.Dispose();
                    _CurrentWorksheetPartWriter = null;
                    _CurrentWorksheetPartStream = null; // closed by writer's Close
                }

                if (_CurrentWorksheetTempPath != null)
                {
                    try
                    {
                        File.Delete(_CurrentWorksheetTempPath);
                    }
                    catch { }
                }

                if (_Package != null)
                {
                    _Package.Dispose();
                    _Package = null;
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

        private PackageWriteStream? _Package = null;
        private CompressionLevel _CompressionLevel = CompressionLevel.Fastest;

        private string? _CurrentWorksheetTempPath = null;
        private Stream? _CurrentWorksheetPartStream = null;
        private StreamWriter? _CurrentWorksheetPartWriter = null;
        private WorksheetInfo? _CurrentWorksheetInfo = null;
        private FrozenPaneState? _CurrentWorksheePane = null;
        private Int64 _WrittenColsPosStart = 0;
        private Int64 _WrittenColsPosEnd = 0;
        private int _RowCount = 0;
        private int _CellCount = 0;

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

        private List<string> _MergeCells = new List<string>();
        private Dictionary<int, List<(int x, Style? style)>> _QueuedMergedCellStyles = new Dictionary<int, List<(int x, Style? style)>>();
        private List<int> _QueuedRowIndexes = new List<int>();
        private int _NextQueuedRowIndex = 0;
        private int _MaxQueuedRowIndex = 0;
        
        private XmlWriterHelper? _XmlWriterHelper = new XmlWriterHelper();

        #endregion

        #region SpreadsheetWriter - Basic properties

        public override string FileExtension => "xlsx";
        public override string FileContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public override bool IsFlatFormat => false;

        #endregion

        #region Public Properties

        public const int MIN_COLUMN_NUMBER = 1;
        public const int MAX_COLUMN_NUMBER = 16384;

        public bool EnableFormulaProtection { get; set; } = true;

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

        private async System.Threading.Tasks.Task HandleMergedCells(int horzCellCount, int vertCellCount, Style? style)
        {
            if (horzCellCount > 1 && (style?.Borders?.Count ?? 0) > 0)
            {
                for (int x = 2; x <= horzCellCount; x++)
                {
                    await WriteCellHeaderAsync(_CellCount + x - 1, _RowCount, true, null, style);
                }
            }
            
            if (vertCellCount > 1 && (style?.Borders?.Count ?? 0) > 0)
            {
                var maxY = _RowCount + vertCellCount - 1;
                
                for (int y = _RowCount + 1; y <= maxY; y++)
                {
                    if (!_QueuedMergedCellStyles.TryGetValue(y, out var mergedCellStyles))
                    {
                        mergedCellStyles = new List<(int x, Style? style)>();
                        _QueuedMergedCellStyles[y] = mergedCellStyles;

                        if (_MaxQueuedRowIndex < y)
                        {
                            _MaxQueuedRowIndex = y;
                            _QueuedRowIndexes.Add(y);

                            if (_NextQueuedRowIndex == 0)
                                _NextQueuedRowIndex = y;
                        }
                    }

                    int insertIndex = mergedCellStyles.Count;

                    if (mergedCellStyles.Count > 0 && mergedCellStyles[0].x > _CellCount)
                        insertIndex = 0;

                    for (int x = 1; x <= horzCellCount; x++)
                    {
                        mergedCellStyles.Insert(insertIndex++, (_CellCount + x - 1, style));
                    }
                }
            }

            if (horzCellCount > 1)
            {
                _CellCount += horzCellCount - 1;
            }
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

        private (int, string?) ConvertNumberFormat(NumberFormat format)
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

        private async Task<bool> WriteStyleBorderXmlAsync(StreamWriter writer, List<Border>? items, bool writeEmpty, int id, int? styleId)
        {
            if (items == null)
            {
                if (writeEmpty)
                    await writer.WriteAsync($"<border><left/><right/><top/><bottom/><diagonal/></border>").ConfigureAwait(false);

                if (styleId != null) _StyleIdBorderIdMap[styleId.Value] = -1;

                return writeEmpty;
            }
            else
            {
                var hasDiagonalUp = items.Any(x => x.Position == BorderPosition.DiagonalRight);
                var hasDiagonalDown = items.Any(x => x.Position == BorderPosition.DiagonalLeft);

                await writer.WriteAsync($"<border{(hasDiagonalUp ? "diagonalUp=\"1\"" : "")}{(hasDiagonalDown ? "diagonalDown =\"1\"" : "")}>").ConfigureAwait(false);

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

                    await writer.WriteAsync($"<{position} style=\"{GetBorderLineStyleName(item.LineStyle, item.Weight)}\">").ConfigureAwait(false);
                    {
                        if (item.Color != System.Drawing.Color.Empty)
                            await writer.WriteAsync($"<color rgb=\"{ColorHelper.GetHexArgb(item.Color)}\"/>").ConfigureAwait(false);
                        else await writer.WriteAsync("<color auto=\"1\"/>").ConfigureAwait(false);
                    }
                    await writer.WriteAsync($"</{position}>").ConfigureAwait(false);
                }

                await writer.WriteAsync("</border>").ConfigureAwait(false);

                if (styleId != null) _StyleIdBorderIdMap[styleId.Value] = id;

                return true;
            }
        }

        private async Task<bool> WriteStyleFontXmlAsync(StreamWriter writer, Font? font, bool writeEmpty, int id, int? styleId)
        {
            if (font == null)
            {
                if (writeEmpty)
                    await writer.WriteAsync($"<font></font>").ConfigureAwait(false);

                if (styleId != null) _StyleIdFontIdMap[styleId.Value] = -1;

                return writeEmpty;
            }
            else
            {
                var item = font.Value;

                await writer.WriteAsync("<font>").ConfigureAwait(false);
                {
                    if (item.Bold) await writer.WriteAsync("<b/>").ConfigureAwait(false);
                    if (item.Italic) await writer.WriteAsync("<i/>").ConfigureAwait(false);

                    if (item.Underline == FontUnderline.Single) await writer.WriteAsync("<u/>").ConfigureAwait(false);
                    if (item.Underline == FontUnderline.SingleAccounting) await writer.WriteAsync("<u val=\"singleAccounting\"/>").ConfigureAwait(false);
                    if (item.Underline == FontUnderline.Double) await writer.WriteAsync("<u val=\"double\"/>").ConfigureAwait(false);
                    if (item.Underline == FontUnderline.DoubleAccounting) await writer.WriteAsync("<u val=\"doubleAccounting\"/>").ConfigureAwait(false);

                    if (item.StrikeThrough) await writer.WriteAsync("<strike/>").ConfigureAwait(false);

                    if (item.VerticalAlign == FontVerticalAlign.Subscript) await writer.WriteAsync("<vertAlign val=\"subscript\"/>").ConfigureAwait(false);
                    if (item.VerticalAlign == FontVerticalAlign.Superscript) await writer.WriteAsync("<vertAlign val=\"superscript\"/>").ConfigureAwait(false);

                    if (item.Size > 0)
                        await writer.WriteAsync($"<sz val=\"{item.Size.ToString("G", _Culture)}\"/>").ConfigureAwait(false);

                    if (item.Color != Color.Empty)
                        await writer.WriteAsync($"<color rgb=\"{ColorHelper.GetHexArgb(item.Color)}\"/>").ConfigureAwait(false);
                    else await writer.WriteAsync("<color auto=\"1\"/>").ConfigureAwait(false);

                    if (!string.IsNullOrEmpty(item.Name))
                        await writer.WriteAsync($"<name val=\"{item.Name}\"/>").ConfigureAwait(false);
                    
                    await writer.WriteAsync($"<family val=\"{GetFontFamilyNumbering(item.Family)}\"/>").ConfigureAwait(false);

                    if (item.Charset != null)
                    {
                        await writer.WriteAsync($"<charset val=\"{(int)item.Charset.Value}\"/>").ConfigureAwait(false);
                    }
                }
                await writer.WriteAsync("</font>").ConfigureAwait(false);

                if (styleId != null) _StyleIdFontIdMap[styleId.Value] = id;

                return true;
            }
        }

        private async Task<bool> WriteStyleFillAsync(StreamWriter writer, Fill? fill, bool writeEmpty, int id, int? styleId)
        {
            if (fill == null || fill?.Pattern == FillPattern.None)
            {
                if (writeEmpty)
                    await writer.WriteAsync($"<fill><patternFill patternType=\"none\"/></fill>").ConfigureAwait(false);

                if (styleId != null) _StyleIdFillIdMap[styleId.Value] = -1;

                return writeEmpty;
            }
            else
            {
                var item = fill!.Value;

                await writer.WriteAsync("<fill>").ConfigureAwait(false);
                {
                    await writer.WriteAsync($"<patternFill patternType=\"{GetPatternName(item.Pattern)}\">").ConfigureAwait(false);

                    var fgColor = item.PatternColor;
                    if (item.Pattern == FillPattern.Solid && fgColor == Color.Empty)
                        fgColor = item.Color;

                    if (fgColor != Color.Empty)
                        await writer.WriteAsync($"<fgColor rgb=\"{ColorHelper.GetHexArgb(fgColor)}\"/>").ConfigureAwait(false);
                    else await writer.WriteAsync("<fgColor auto=\"1\"/>").ConfigureAwait(false);

                    if (item.Pattern != FillPattern.Solid)
                    {
                        if (item.Color != Color.Empty)
                            await writer.WriteAsync($"<bgColor rgb=\"{ColorHelper.GetHexArgb(item.Color)}\"/>").ConfigureAwait(false);
                        else await writer.WriteAsync("<bgColor auto=\"1\"/>").ConfigureAwait(false);
                    }

                    await writer.WriteAsync("</patternFill>").ConfigureAwait(false);
                }
                await writer.WriteAsync("</fill>").ConfigureAwait(false);

                if (styleId != null) _StyleIdFillIdMap[styleId.Value] = id;

                return true;
            }
        }

        private async Task<bool> WriteStyleNumberFormatXmlAsync(StreamWriter writer, NumberFormat format, bool writeEmpty, int id, int? styleId)
        {
            var convertedFormat = ConvertNumberFormat(format);

            if (convertedFormat.Item1 == -2)
            {
                await writer.WriteAsync($"<numFmt numFmtId=\"{id}\" formatCode=\"{XmlWriterHelper.EscapeSimpleXmlAttr(convertedFormat.Item2 ?? "")}\"/>").ConfigureAwait(false);
                if (styleId != null) _StyleIdNumberFormatIdMap[styleId.Value] = id;
                return true;
            }
            else if (convertedFormat.Item1 == -1)
            {
                if (writeEmpty)
                    await writer.WriteAsync($"<numFmt numFmtId=\"0\"/>").ConfigureAwait(false);

                if (styleId != null) _StyleIdNumberFormatIdMap[styleId.Value] = -1;
                return false;
            }
            else
            {
                if (styleId != null) _StyleIdNumberFormatIdMap[styleId.Value] = convertedFormat.Item1;
                return false;
            }
        }

        private async Task WriteStyleXfXmlAsync(StreamWriter writer, int styleId, Style style)
        {
            var numFmtId = _StyleIdNumberFormatIdMap[styleId];
            var borderId = _StyleIdBorderIdMap[styleId];
            var fillId = _StyleIdFillIdMap[styleId];
            var fontId = _StyleIdFontIdMap[styleId];

            await writer.WriteAsync($"<xf").ConfigureAwait(false);
            await writer.WriteAsync($" numFmtId=\"{(numFmtId == -1 ? 0 : numFmtId)}\"").ConfigureAwait(false);
            await writer.WriteAsync($" borderId=\"{(borderId == -1 ? 0 : borderId)}\"").ConfigureAwait(false);
            await writer.WriteAsync($" fillId=\"{(fillId == -1 ? 0 : fillId)}\"").ConfigureAwait(false);
            await writer.WriteAsync($" fontId=\"{(fontId == -1 ? 0 : fontId)}\"").ConfigureAwait(false);

            if (numFmtId > -1)
                await writer.WriteAsync(" applyNumberFormat=\"1\"").ConfigureAwait(false);
            if (borderId > -1)
                await writer.WriteAsync(" applyBorder=\"1\"").ConfigureAwait(false);
            if (fillId > -1)
                await writer.WriteAsync(" applyFill=\"1\"").ConfigureAwait(false);
            if (fontId > -1)
                await writer.WriteAsync(" applyFont=\"1\"").ConfigureAwait(false);
            if (style.Alignment != null)
                await writer.WriteAsync(" applyAlignment=\"1\"").ConfigureAwait(false);

            if (style.Alignment != null)
            {
                var align = style.Alignment.Value;

                await writer.WriteAsync(">").ConfigureAwait(false);
                {
                    await writer.WriteAsync("<alignment").ConfigureAwait(false);

                    if (align.VerticalText == true)
                    {
                        await writer.WriteAsync(" textRotation=\"255\"").ConfigureAwait(false);
                    }
                    else if (align.Rotate != 0d)
                    {
                        await writer.WriteAsync($" textRotation=\"{Math.Max(-90d, Math.Min(90d, align.Rotate)).ToString("G", _Culture)}\"").ConfigureAwait(false);
                    }

                    string? horz = null;
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
                        await writer.WriteAsync($" horizontal=\"{horz}\"").ConfigureAwait(false);

                    string? vert = null;
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
                        await writer.WriteAsync($" vertical=\"{vert}\"").ConfigureAwait(false);

                    if (align.WrapText)
                        await writer.WriteAsync($" wrapText=\"1\"").ConfigureAwait(false);

                    if (align.ShrinkToFit)
                        await writer.WriteAsync($" shrinkToFit=\"1\"").ConfigureAwait(false);

                    if (align.ReadingOrder == HorizontalReadingOrder.LeftToRight)
                        await writer.WriteAsync($" readingOrder=\"1\"").ConfigureAwait(false);
                    else if (align.ReadingOrder == HorizontalReadingOrder.RightToLeft)
                        await writer.WriteAsync($" readingOrder=\"2\"").ConfigureAwait(false);

                    if (align.Indent > 0)
                        await writer.WriteAsync($" indent=\"{align.Indent}\"").ConfigureAwait(false);

                    await writer.WriteAsync("/>").ConfigureAwait(false);
                }
                await writer.WriteAsync("</xf>").ConfigureAwait(false);
            }
            else
            {
                await writer.WriteAsync("/>").ConfigureAwait(false);
            }
        }

        private async Task WriteStylesXmlAsync(Stream stream)
        {
            using (var writer = new StreamWriter(stream))
            {
                await writer.WriteAsync("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n").ConfigureAwait(false);
                await writer.WriteAsync("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">").ConfigureAwait(false);
                {
                    var orderedStyles = _Styles.OrderBy(p => p.Value);

                    var nextId = 124;
                    var customFormatCount = orderedStyles.Count(x => x.Key.NumberFormat.Type == NumberFormatType.Custom);
                    await writer.WriteAsync($"<numFmts count=\"{customFormatCount}\">").ConfigureAwait(false);
                    foreach (var pair in orderedStyles)
                    {
                        if (await WriteStyleNumberFormatXmlAsync(writer, pair.Key.NumberFormat, false, nextId, pair.Value))
                            nextId++;
                    }
                    await writer.WriteAsync("</numFmts>").ConfigureAwait(false);

                    nextId = 1;
                    await writer.WriteAsync($"<fonts x14ac:knownFonts=\"1\" count=\"{_Styles.Count(x => x.Key.Font != null) + 1}\">").ConfigureAwait(false);
                    await WriteStyleFontXmlAsync(writer, null, true, 0, null);
                    foreach (var pair in orderedStyles)
                    {
                        if (await WriteStyleFontXmlAsync(writer, pair.Key.Font, false, nextId, pair.Value))
                            nextId++;
                    }
                    await writer.WriteAsync("</fonts>").ConfigureAwait(false);

                    nextId = 2; // We need 2 dummies, for some reason
                    await writer.WriteAsync($"<fills count=\"{_Styles.Count(x => x.Key.Fill != null && x.Key.Fill.Value.Pattern != FillPattern.None) + 2}\">").ConfigureAwait(false);
                    await WriteStyleFillAsync(writer, null, true, 0, null);
                    await WriteStyleFillAsync(writer, null, true, 0, null);
                    foreach (var pair in orderedStyles)
                    {
                        if (await WriteStyleFillAsync(writer, pair.Key.Fill, false, nextId, pair.Value))
                            nextId++;
                    }
                    await writer.WriteAsync("</fills>").ConfigureAwait(false);

                    nextId = 1;
                    await writer.WriteAsync($"<borders count=\"{_Styles.Count(x => x.Key.Borders != null) + 1}\">").ConfigureAwait(false);
                    await WriteStyleBorderXmlAsync(writer, null, true, 0, null);
                    foreach (var pair in orderedStyles)
                    {
                        if (await WriteStyleBorderXmlAsync(writer, pair.Key.Borders, false, nextId, pair.Value))
                            nextId++;
                    }
                    await writer.WriteAsync("</borders>").ConfigureAwait(false);

                    await writer.WriteAsync($"<cellXfs count=\"{_Styles.Count + 1}\">").ConfigureAwait(false);
                    await writer.WriteAsync($"<xf numFmtId=\"0\" borderId=\"0\" fillId=\"0\" fontId=\"0\"/>").ConfigureAwait(false);
        
                    foreach (var pair in orderedStyles)
                    {
                        await WriteStyleXfXmlAsync(writer, pair.Value, pair.Key);
                    }
                    await writer.WriteAsync("</cellXfs>").ConfigureAwait(false);

                }
                await writer.WriteAsync("</styleSheet>").ConfigureAwait(false);
            }
        }

        private async Task WriteWorkbookXmlAsync(Stream stream)
        {
            using (var writer = new StreamWriter(stream))
            {
                await writer.WriteAsync("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n").ConfigureAwait(false);
                await writer.WriteAsync("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">").ConfigureAwait(false);

                await writer.WriteAsync("<sheets>").ConfigureAwait(false);

                foreach (var item in _WorksheetInfos)
                {
                    var name = item.Name ?? $"Worksheet{item.Id}";

                    // Remove invalid characters
                    name = Regex.Replace(name, "[:/\\?[\\]]", " ");

                    // Cannot begin/end with an apostrophe
                    name = Regex.Replace(name, "^'+|'+$", "");

                    // Excel limits to 31 characters. Otherwise it's an error
                    name = name.Substring(0, Math.Min(name.Length, 31));

                    await writer.WriteAsync($"<sheet r:id=\"rId{item.Id}\" sheetId=\"{item.Id}\" name=\"{_XmlWriterHelper!.EscapeAttribute(name)}\"/>").ConfigureAwait(false);
                }

                await writer.WriteAsync("</sheets>").ConfigureAwait(false);
                await writer.WriteAsync("</workbook>").ConfigureAwait(false);
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

            const int @base = 26;
            columnNumber -= 1;
            
            var address = string.Empty; 
            
            do 
            {
                address = Convert.ToChar('A' + columnNumber % @base) + address;
                columnNumber = columnNumber / @base - 1; 
            } 
            while (columnNumber >= 0);

            return address;
        }

        #endregion

        #region Syling

        public override void RegisterStyle(Style style)
        {
            GetStyleId(style);
        }

        #endregion

        #region AutoFit

        private AutoFitConfig? _DefaultAutoFitConfig = null;
        private Dictionary<int, AutoFitConfig> _AutoFitConfig = new Dictionary<int, AutoFitConfig>();
        private Dictionary<int, float> _AutoFitState = new Dictionary<int, float>();

        private void UpdateAutoFitForCell(int index, object? value)
        {
            AutoFitConfig? conf = null;

            if (!_AutoFitConfig.TryGetValue(index, out conf))
            {
                conf = _DefaultAutoFitConfig;
            }

            if (conf != null)
            {
                if (!_AutoFitState.TryGetValue(index, out var current))
                    current = -1f;

                if (current > -1f && current >= conf.MaxLength)
                    return;

                var size = 0f;

                if (conf.Measure != null)
                    size = conf.Measure(index, value);
                else
                {
                    var text = value?.ToString() ?? "";
                    if (conf.Multiline)
                    {
                        var max = 0;
                        foreach (var c in text)
                        {
                            if (c == '\n')
                            {
                                size = Math.Max(size, max);
                                max = 0;
                            }
                            else
                            {
                                max++;
                            }
                        }

                        size = Math.Max(size, max);
                    }
                    else
                    {
                        size = text.Length;
                    }

                    size *= conf.Multiplier;
                }

                if (size > current)
                    _AutoFitState[index] = size;
            }
        }

        public void EnableAutoFitForDefaultColumn(int index, AutoFitConfig config)
        {
            _DefaultAutoFitConfig = config;
        }

        public void DisableAutoFitForDefaultColumn(int index)
        {
            _DefaultAutoFitConfig = null;
        }

        public void EnableAutoFitForColumn(int index, AutoFitConfig config)
        {
            _AutoFitConfig[index] = config;
        }

        public void DisableAutoFitForColumn(int index)
        {
            _AutoFitConfig.Remove(index);
        }

        #endregion

        #region SpreadsheetWriter - Document Lifespan (private)

        private const string _WORKBOOK_PATH = "/xl/workbook.xml";

        private async Task EndPackageAsync()
        {
            string stylesheetPath = "/xl/styles.xml";
            string sharedStringsPath = "/xl/sharedStrings.xml";
            string docPropsCorePath = "/docProps/core.xml";
            string docPropsAppPath = "/docProps/app.xml";

            var wbEntry = _Package!.CreateEntry(_WORKBOOK_PATH, _CompressionLevel);
            using (var stream = wbEntry.Open())
            {
                _Package.AddPackageRelationship(_WORKBOOK_PATH, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
                _Package.AddContentType(_WORKBOOK_PATH, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");

                await WriteWorkbookXmlAsync(stream);
            }

            var stylesEntry = _Package.CreateEntry(stylesheetPath, _CompressionLevel);
            using (var stream = stylesEntry.Open())
            {
                _Package.AddPartRelationship(_WORKBOOK_PATH, stylesheetPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
                _Package.AddContentType(stylesheetPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");

                await WriteStylesXmlAsync(stream);
            }

            var sharedStringsEntry = _Package.CreateEntry(sharedStringsPath, _CompressionLevel);
            _Package.AddPartRelationship(_WORKBOOK_PATH, sharedStringsPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
            _Package.AddContentType(sharedStringsPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");

            using (var stream = sharedStringsEntry.Open())
            using (var writer = new StreamWriter(stream))
            {
                await writer.WriteAsync("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n").ConfigureAwait(false);
                await writer.WriteAsync("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>").ConfigureAwait(false);
            }

            if (SpreadsheetInfo != null)
            {
                var docPropsCoreEntry = _Package.CreateEntry(docPropsCorePath, _CompressionLevel);
                _Package.AddPackageRelationship(docPropsCorePath, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
                _Package.AddContentType(docPropsCorePath, "application/vnd.openxmlformats-package.core-properties+xml");

                using (var stream = docPropsCoreEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    await writer.WriteAsync("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n").ConfigureAwait(false);
                    await writer.WriteAsync("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"" + 
                        " xmlns:dc=\"http://purl.org/dc/elements/1.1/\"" +
                        " xmlns:dcterms=\"http://purl.org/dc/terms/\"" + 
                        " xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\"" +
                        " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">").ConfigureAwait(false);

                    var xmlWriter = _XmlWriterHelper!;

                    if (SpreadsheetInfo.Title != null)
                    {
                        await writer.WriteAsync($"<dc:title>{xmlWriter.EscapeValue(SpreadsheetInfo.Title)}</dc:title>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.Subject != null)
                    {
                        await writer.WriteAsync($"<dc:subject>{xmlWriter.EscapeValue(SpreadsheetInfo.Subject)}</dc:subject>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.Author != null)
                    {
                        await writer.WriteAsync($"<dc:creator>{xmlWriter.EscapeValue(SpreadsheetInfo.Author)}</dc:creator>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.Keywords != null)
                    {
                        await writer.WriteAsync($"<cp:keywords>{xmlWriter.EscapeValue(SpreadsheetInfo.Keywords)}</cp:keywords>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.Comments != null)
                    {
                        await writer.WriteAsync($"<dc:description>{xmlWriter.EscapeValue(SpreadsheetInfo.Comments)}</dc:description>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.Status != null)
                    {
                        await writer.WriteAsync($"<cp:contentStatus>{xmlWriter.EscapeValue(SpreadsheetInfo.Status)}</cp:contentStatus>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.Category != null)
                    {
                        await writer.WriteAsync($"<cp:category>{xmlWriter.EscapeValue(SpreadsheetInfo.Category)}</cp:category>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.LastModifiedBy != null)
                    {
                        await writer.WriteAsync($"<cp:lastModifiedBy>{xmlWriter.EscapeValue(SpreadsheetInfo.LastModifiedBy)}</cp:lastModifiedBy>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.CreatedOn != null)
                    {
                        await writer.WriteAsync($"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{SpreadsheetInfo.CreatedOn.Value.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")}</dcterms:created>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.ModifiedOn != null)
                    {
                        await writer.WriteAsync($"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{SpreadsheetInfo.ModifiedOn.Value.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")}</dcterms:modified>").ConfigureAwait(false);
                    }

                    await writer.WriteAsync("</cp:coreProperties>").ConfigureAwait(false);
                }

                var docPropsAppEntry = _Package.CreateEntry(docPropsAppPath, _CompressionLevel);
                _Package.AddPackageRelationship(docPropsAppPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
                _Package.AddContentType(docPropsAppPath, "application/vnd.openxmlformats-officedocument.extended-properties+xml");

                using (var stream = docPropsAppEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    await writer.WriteAsync("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n").ConfigureAwait(false);
                    await writer.WriteAsync("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"" +
                        " xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">").ConfigureAwait(false);

                    var xmlWriter = _XmlWriterHelper!;

                    if (SpreadsheetInfo.Application != null)
                    {
                        await writer.WriteAsync($"<Application>{xmlWriter.EscapeValue(SpreadsheetInfo.Application)}</Application>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.ScaleCrop != null)
                    {
                        await writer.WriteAsync($"<ScaleCrop>{(SpreadsheetInfo.ScaleCrop.Value ? "true" : "false")}</ScaleCrop>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.Manager != null)
                    {
                        await writer.WriteAsync($"<Manager>{xmlWriter.EscapeValue(SpreadsheetInfo.Manager)}</Manager>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.Company != null)
                    {
                        await writer.WriteAsync($"<Company>{xmlWriter.EscapeValue(SpreadsheetInfo.Company)}</Company>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.LinksUpToDate != null)
                    {
                        await writer.WriteAsync($"<LinksUpToDate>{(SpreadsheetInfo.LinksUpToDate.Value ? "true" : "false")}</LinksUpToDate>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.SharedDoc != null)
                    {
                        await writer.WriteAsync($"<SharedDoc>{(SpreadsheetInfo.SharedDoc.Value ? "true" : "false")}</SharedDoc>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.HyperlinksChanged != null)
                    {
                        await writer.WriteAsync($"<HyperlinksChanged>{(SpreadsheetInfo.HyperlinksChanged.Value ? "true" : "false")}</HyperlinksChanged>").ConfigureAwait(false);
                    }

                    if (SpreadsheetInfo.AppVersion != null)
                    {
                        await writer.WriteAsync($"<AppVersion>{xmlWriter.EscapeValue(SpreadsheetInfo.AppVersion)}</AppVersion>").ConfigureAwait(false);
                    }

                    await writer.WriteAsync("</Properties>").ConfigureAwait(false);
                }
            }

            await _Package.CommitRichDataAsync(_CompressionLevel);
            await _Package.CommitRelationshipsAsync(_CompressionLevel);
            await _Package.CommitContentTypesAsync(_CompressionLevel);
            _Package.Close();
            _Package.Dispose();

            _Package = null;
        }

        private async Task WritePendingBeginWorksheetAsync()
        {
            if (!_ShouldBeginWorksheet)
                return;
            
            if (_CurrentWorksheetPartWriter != null)
            {
                _CurrentWorksheetPartWriter.Dispose();
            }

            _CurrentWorksheetTempPath = System.IO.Path.GetTempFileName();

            _CurrentWorksheetPartStream = System.IO.File.Open(_CurrentWorksheetTempPath, FileMode.Create, FileAccess.ReadWrite, FileShare.Read);
            _CurrentWorksheetPartWriter = new StreamWriter(_CurrentWorksheetPartStream, Encoding.UTF8);

            await _CurrentWorksheetPartWriter.WriteAsync("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n").ConfigureAwait(false);
            await _CurrentWorksheetPartWriter.WriteAsync("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">").ConfigureAwait(false);
            {
                await _CurrentWorksheetPartWriter.WriteAsync("<sheetViews><sheetView").ConfigureAwait(false);
                if (_CurrentWorksheetInfo!.RightToLeft != null)
                {
                    await _CurrentWorksheetPartWriter.WriteAsync($" rightToLeft=\"{(_CurrentWorksheetInfo.RightToLeft == true ? "1" : "0")}\"").ConfigureAwait(false);
                }
                await _CurrentWorksheetPartWriter.WriteAsync(" workbookViewId=\"0\">").ConfigureAwait(false);

                if (_CurrentWorksheePane != null &&
                    (_CurrentWorksheePane.Value.Column > 1 || _CurrentWorksheePane.Value.Row > 1))
                {
                    await _CurrentWorksheetPartWriter.WriteAsync(@"<pane").ConfigureAwait(false);

                    if (_CurrentWorksheePane.Value.Column > 1)
                        await _CurrentWorksheetPartWriter.WriteAsync($@" xSplit=""{_CurrentWorksheePane.Value.Column - 1}""").ConfigureAwait(false);

                    if (_CurrentWorksheePane.Value.Row > 1)
                        await _CurrentWorksheetPartWriter.WriteAsync($@" ySplit=""{_CurrentWorksheePane.Value.Row - 1}""").ConfigureAwait(false);

                    await _CurrentWorksheetPartWriter.WriteAsync($@" topLeftCell=""{ConvertColumnAddress(_CurrentWorksheePane.Value.Column)}{_CurrentWorksheePane.Value.Row}""").ConfigureAwait(false);
                    await _CurrentWorksheetPartWriter.WriteAsync($@" activePane=""bottomRight""").ConfigureAwait(false);
                    await _CurrentWorksheetPartWriter.WriteAsync(@" state=""frozen""/>").ConfigureAwait(false);
                }

                await _CurrentWorksheetPartWriter.WriteAsync("</sheetView></sheetViews>").ConfigureAwait(false);

                await _CurrentWorksheetPartWriter.WriteAsync("<sheetFormatPr").ConfigureAwait(false);

                if (_CurrentWorksheetInfo.DefaultRowHeight != null)
                    await _CurrentWorksheetPartWriter.WriteAsync(string.Format(
                        " defaultRowHeight=\"{0:R15}\"",
                        Math.Max(0f, Math.Min(_CurrentWorksheetInfo.DefaultRowHeight.Value, 409.5f)))).ConfigureAwait(false);

                if (_CurrentWorksheetInfo.DefaultColumnWidth != null)
                    await _CurrentWorksheetPartWriter.WriteAsync(string.Format(" baseColWidth=\"{0:R15}\"", _CurrentWorksheetInfo.DefaultColumnWidth)).ConfigureAwait(false);

                await _CurrentWorksheetPartWriter.WriteAsync("/>").ConfigureAwait(false);

                await _CurrentWorksheetPartWriter.FlushAsync();
                _WrittenColsPosStart = _CurrentWorksheetPartWriter.BaseStream.Position;
                var colsXml = GenerateColumnsXml();
                if (colsXml != null)
                {
                    await _CurrentWorksheetPartWriter.WriteAsync(colsXml).ConfigureAwait(false);
                    await _CurrentWorksheetPartWriter.FlushAsync();
                }
                _WrittenColsPosEnd = _CurrentWorksheetPartWriter.BaseStream.Position;

                await _CurrentWorksheetPartWriter.WriteAsync("<sheetData>").ConfigureAwait(false);
            }

            _ShouldBeginWorksheet = false;
        }

        private string? GenerateColumnsXml()
        {
            if (_CurrentWorksheetInfo == null)
                return null;

            var sb = new StringBuilder();

            var columnInfos = new List<ColumnInfo>();

            if (_CurrentWorksheetInfo.ColumnInfos != null)
                columnInfos.AddRange(_CurrentWorksheetInfo.ColumnInfos);

            foreach (var c in _AutoFitState.Keys)
            {
                if (!columnInfos.Any(x => c >= x.FromColumn && c <= x.ToColumn))
                {
                    columnInfos.Add(new ColumnInfo
                    {
                        FromColumn = c,
                        ToColumn = c,
                    });
                }
            }

            if (columnInfos != null)
            {
                var colsIndicesWithAutoFit = _AutoFitState.Keys.ToHashSet();

                void outputRange(ColumnInfo ci, int min, int max, float? autoFitWidth = null)
                {
                    sb.Append($"<col min=\"{min + 1}\" max=\"{max + 1}\" bestFit=\"0\"");

                    if (autoFitWidth != null)
                        sb.Append($" width=\"{autoFitWidth.Value:G}\" customWidth=\"1\"");
                    else if (ci.Width != null && ci.Width != 0f && ci.Width != _CurrentWorksheetInfo.DefaultColumnWidth)
                        sb.Append($" width=\"{ci.Width.Value:G}\" customWidth=\"1\"");

                    if (ci.Hidden)
                        sb.Append($" hidden=\"1\"");


                    sb.Append($" />");
                };

                for (int i = 0; i < columnInfos.Count; i++)
                {
                    var ci = columnInfos[i];

                    var matches = colsIndicesWithAutoFit
                        .Where(i => i >= ci.FromColumn && i <= ci.ToColumn)
                        .OrderBy(i => i)
                        .ToList();

                    // Case 1: No auto-fit columns inside → output whole range
                    if (matches.Count == 0)
                    {
                        outputRange(ci, ci.FromColumn, ci.ToColumn);
                        continue;
                    }

                    // Case 2: Need to split the range
                    int current = ci.FromColumn;

                    foreach (var m in matches)
                    {
                        colsIndicesWithAutoFit.Remove(m);

                        // Output the non-auto-fit chunk before the match
                        if (current <= m - 1)
                            outputRange(ci, current, m - 1);

                        // Output the auto-fit column as its own range
                        outputRange(ci, m, m, _AutoFitState[m]); // single column

                        current = m + 1;
                    }

                    // Output the tail chunk after the last auto-fit index
                    if (current <= ci.ToColumn)
                        outputRange(ci, current, ci.ToColumn);
                }

                foreach (var m in colsIndicesWithAutoFit)
                {
                    outputRange(new ColumnInfo { }, m, m, _AutoFitState[m]);
                }
            }

            if (sb.Length > 0)
            {
                sb.Insert(0, "<cols>");
                sb.Append("</cols>");
            }

            return sb.ToString();
        }

        private async Task WritePendingEndRowAsync()
        {
            if (!_ShouldEndRow) return;

            // write missing merge cells after last the cell that was written in this row
            if (_NextQueuedRowIndex == _RowCount)
            {
                await WriteMergedCellCounterparts(_RowCount);
            }

            await _CurrentWorksheetPartWriter!.WriteAsync("</row>").ConfigureAwait(false);

            _ShouldEndRow = false;
        }
        
        private async Task WritePendingEndWorksheetAsync()
        {
            if (!_ShouldEndWorksheet)
                return;
                
            if (_ShouldBeginWorksheet)
                await WritePendingBeginWorksheetAsync();

            await WritePendingEndRowAsync();

            await _CurrentWorksheetPartWriter!.WriteAsync("</sheetData>").ConfigureAwait(false);

            if (_MergeCells.Count > 0)
            {
                await _CurrentWorksheetPartWriter.WriteAsync($"<mergeCells count=\"{_MergeCells.Count}\">").ConfigureAwait(false);

                foreach (var merge in _MergeCells)
                {
                    await _CurrentWorksheetPartWriter.WriteAsync($"<mergeCell ref=\"{merge}\" />").ConfigureAwait(false);
                }

                await _CurrentWorksheetPartWriter.WriteAsync("</mergeCells>").ConfigureAwait(false);
            }

            await _CurrentWorksheetPartWriter.WriteAsync("</worksheet>").ConfigureAwait(false);

            await _CurrentWorksheetPartWriter!.FlushAsync();
            _CurrentWorksheetPartWriter.Dispose();
            _CurrentWorksheetPartWriter = null;
            _CurrentWorksheetPartStream = null;

            if (_AutoFitState.Count > 0)
            {
                var xml = GenerateColumnsXml();

                if (xml != null)
                {
                    var autoFitTmpFile = System.IO.Path.GetTempFileName();
                    using var autoFitStream = System.IO.File.Open(autoFitTmpFile, FileMode.Create, FileAccess.ReadWrite, FileShare.Read);
                    using var autoFitWriter = new StreamWriter(autoFitStream, Encoding.UTF8);

                    using (var input = new FileStream(_CurrentWorksheetTempPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        CopyBytes(input, autoFitStream, _WrittenColsPosStart);
                        await autoFitWriter.WriteAsync(xml).ConfigureAwait(false);
                        await autoFitWriter.FlushAsync().ConfigureAwait(false);
                        input.Position = _WrittenColsPosEnd;
                        await input.CopyToAsync(autoFitStream).ConfigureAwait(false);
                    }

                    try
                    {
                        File.Delete(_CurrentWorksheetTempPath);
                    }
                    catch { }

                    _CurrentWorksheetTempPath = autoFitTmpFile;
                }
            }

            var packageEntry = _Package!.CreateEntry(_CurrentWorksheetInfo!.Path, _CompressionLevel);
            var entryStream = packageEntry.Open();
            using (var readStream = File.Open(_CurrentWorksheetTempPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                await readStream.CopyToAsync(entryStream);
            entryStream.Dispose();

            try
            {
                File.Delete(_CurrentWorksheetTempPath);
            }
            catch { }
            _CurrentWorksheetTempPath = null;

            _ShouldEndWorksheet = false;
        }

        private static void CopyBytes(Stream input, Stream output, Int64 count)
        {
            byte[] buffer = new byte[4096];
            long remaining = count;

            while (remaining > 0)
            {
                int toRead = (int)Math.Min(buffer.Length, remaining);
                int read = input.Read(buffer, 0, toRead);
                if (read == 0) break;
                output.Write(buffer, 0, read);
                remaining -= read;
            }
        }

        private void BeginFile()
        {
            if (_WroteFileStart) return;
            _WroteFileStart = true;
        }

        #endregion
        
        #region SpreadsheetWriter - Document Lifespan (public)

        public override async Task NewWorksheetAsync(WorksheetInfo info)
        {
            await WritePendingEndWorksheetAsync().ConfigureAwait(false);

            if (_WorksheetInfos.Contains(info))
                throw new InvalidOperationException("This WorksheetInfo object has already been added to the workbook");

            info.Path = $"/xl/worksheets/sheet{_WorksheetInfos.Count + 1}.xml";

            info.Id = _Package!.AddPartRelationship(_WORKBOOK_PATH, info.Path, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            _Package!.AddContentType(info.Path, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");

            _CurrentWorksheetInfo = info;
            _WorksheetInfos.Add(_CurrentWorksheetInfo);

            _CurrentWorksheePane = null;

            _ShouldBeginWorksheet = true;
            _ShouldEndWorksheet = true;
            
            _RowCount = 0;
            _MergeCells.Clear();
            _QueuedMergedCellStyles.Clear();
            _QueuedRowIndexes.Clear();
            _NextQueuedRowIndex = 0;
            _MaxQueuedRowIndex = 0;

            _DefaultAutoFitConfig = null;
            _AutoFitConfig.Clear();
            _AutoFitState.Clear();
        }

        public void SetWorksheetFrozenPane(FrozenPaneState? pane)
        {
            _CurrentWorksheePane = pane;
        }

        public override Task SkipRowAsync()
        {
            return SkipRowsAsync(1);
        }
        
        private async Task SkipRowsAsyncImpl(int count)
        {
            if (count > 0 && _NextQueuedRowIndex > _RowCount && _NextQueuedRowIndex <= _RowCount + count)
            {
                var max = Math.Min(_MaxQueuedRowIndex, _RowCount + count);
                
                for (int y = _NextQueuedRowIndex; y <= max; y++)
                {
                    await _CurrentWorksheetPartWriter!.WriteAsync($"<row r=\"{y}\">").ConfigureAwait(false);
                    await WriteMergedCellCounterparts(y);
                    await _CurrentWorksheetPartWriter.WriteAsync($"</row>").ConfigureAwait(false);

                    if (_NextQueuedRowIndex > 0)
                        y = _NextQueuedRowIndex - 1;
                    else
                        y = max;
                }
            }
            
            _RowCount += count;
        }
        
        private async Task WriteMergedCellCounterparts(int row, int maxX = 0)
        {
            if (!_QueuedMergedCellStyles.TryGetValue(row, out var cells))
                return;

            int writtenCount = 0;
            foreach (var cell in cells)
            {
                if (maxX > 0 && cell.x > maxX) break;
                
                await WriteCellHeaderAsync(cell.x, row, true, null, cell.style);
                writtenCount++;
            }
            
            if (writtenCount < cells.Count)
            {
                cells.RemoveRange(0, writtenCount);
                return;
            }

            _QueuedMergedCellStyles.Remove(row);
            _QueuedRowIndexes.RemoveAt(0);

            if (_QueuedRowIndexes.Count > 0)
            {
                _NextQueuedRowIndex = _QueuedRowIndexes[0];
            }
            else
            {
                _NextQueuedRowIndex = 0;
                _MaxQueuedRowIndex = 0;
            }
        }

        public override async Task SkipRowsAsync(int count)
        {
            if (!_ShouldEndWorksheet)
            {
                throw new InvalidOperationException("Adding new rows is not allowed at this time. Please call NewWorksheet(...) first.");
            }

            if (!_WroteFileStart)
            {
                BeginFile();
            }
            
            await WritePendingBeginWorksheetAsync();
            await WritePendingEndRowAsync();
            await SkipRowsAsyncImpl(count);

            _CellCount = 0;
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

                BeginFile();
            }

            await WritePendingBeginWorksheetAsync();
            await WritePendingEndRowAsync();

            _RowCount ++;
            _CellCount = 0;
            await _CurrentWorksheetPartWriter!.WriteAsync($"<row r=\"{_RowCount}\"").ConfigureAwait(false);

            if (height != 0f && height != _CurrentWorksheetInfo!.DefaultRowHeight)
            {
                if (!autoFit)
                    await _CurrentWorksheetPartWriter.WriteAsync(" customHeight=\"1\"").ConfigureAwait(false);

                await _CurrentWorksheetPartWriter.WriteAsync($" ht=\"{height.ToString("G", _Culture)}\"").ConfigureAwait(false);
            }
            else if (_CurrentWorksheetInfo!.DefaultRowHeight != null)
            {
                if (!autoFit)
                    await _CurrentWorksheetPartWriter.WriteAsync(" customHeight=\"1\"").ConfigureAwait(false);
            }

            if (style != null)
                await _CurrentWorksheetPartWriter.WriteAsync($" s=\"{GetStyleId(style)}\"").ConfigureAwait(false);

            await _CurrentWorksheetPartWriter.WriteAsync(">").ConfigureAwait(false);

            _ShouldEndRow = true;
        }

        private async Task WriteCellHeaderAsync(int cellIndex, int rowIndex, bool closed, string? type, Style? style = null, int valueMetadataIndex = -1)
        {
            await _CurrentWorksheetPartWriter!.WriteAsync($"<c r=\"{ConvertColumnAddress(cellIndex)}{rowIndex}\"").ConfigureAwait(false);

            if (type != null)
                await _CurrentWorksheetPartWriter.WriteAsync(" t=\"" + type + "\"").ConfigureAwait(false);

            if (style != null)
                await _CurrentWorksheetPartWriter.WriteAsync($" s=\"{GetStyleId(style)}\"").ConfigureAwait(false);

            if (valueMetadataIndex > 0)
                await _CurrentWorksheetPartWriter.WriteAsync($" vm=\"{valueMetadataIndex}\"").ConfigureAwait(false);

            await _CurrentWorksheetPartWriter.WriteAsync(closed ? "/>" : ">").ConfigureAwait(false);
        }

        private async Task WriteCellFooterAsync()
        {
            await _CurrentWorksheetPartWriter!.WriteAsync("</c>").ConfigureAwait(false);
        }

        public override async Task FinishAsync()
        {
            if (!_WroteFileStart)
            {
                BeginFile();
                await NewWorksheetAsync(new WorksheetInfo { }).ConfigureAwait(false);
            }

            await WritePendingEndWorksheetAsync().ConfigureAwait(false);

            if (!_WroteFileEnd)
            {
                await EndPackageAsync().ConfigureAwait(false);
                _WroteFileEnd = true;
            }
        }

        #endregion

        #region SpreadsheetWriter - Cell methods

        public override Task SkipCellAsync()
        {
            return SkipCellsAsync(1);
        }

        public override Task SkipCellsAsync(int count)
        {
            _CellCount += count;
            return Task.CompletedTask;
        }

        public override async Task AddCellAsync(string? data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);
            
            MergeNextCell(horzCellCount, vertCellCount);

            if (data != null && EnableFormulaProtection && data.StartsWith("="))
                data = "'" + data;

            if (data != null && data.Length > 32767)
                data = data.Remove(32767);

            UpdateAutoFitForCell(_CellCount, data);

            await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, false, ExcelValueTypes.FormulaString, style);
            await _CurrentWorksheetPartWriter!.WriteAsync("<v>" + _XmlWriterHelper!.EscapeValue(data) + "</v>").ConfigureAwait(false);
            await WriteCellFooterAsync();

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }

        public override async Task AddCellStringAutoTypeAsync(string? data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);
            
            MergeNextCell(horzCellCount, vertCellCount);

            var type = ExcelValueTypes.FormulaString;

            if (style?.NumberFormat != null && (
                style.NumberFormat.Type == NumberFormatType.GeneralNumber ||
                style.NumberFormat.Type == NumberFormatType.Scientific ||
                style.NumberFormat.Type == NumberFormatType.Fixed ||
                style.NumberFormat.Type == NumberFormatType.Standard))
            {
                type = ExcelValueTypes.Number;
            }

            if (data != null && EnableFormulaProtection && data.StartsWith("="))
                data = "'" + data;

            if (data != null && data.Length > 32767)
                data = data.Remove(32767);

            UpdateAutoFitForCell(_CellCount, data);

            await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, false, type, style);
            await _CurrentWorksheetPartWriter!.WriteAsync("<v>" + _XmlWriterHelper!.EscapeValue(data) + "</v>").ConfigureAwait(false);
            await WriteCellFooterAsync();

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }

        public override Task AddCellForcedStringAsync(string? data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            return AddCellAsync(data, style, horzCellCount, vertCellCount);
        }

        public override async Task AddCellAsync(Int32 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);
            
            MergeNextCell(horzCellCount, vertCellCount);

            UpdateAutoFitForCell(_CellCount, data);

            await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, false, ExcelValueTypes.Number, style);
            await _CurrentWorksheetPartWriter!.WriteAsync(string.Format(_Culture, "<v>{0:G}</v>", data)).ConfigureAwait(false);
            await WriteCellFooterAsync();

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public override async Task AddCellAsync(UInt32 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);
            
            MergeNextCell(horzCellCount, vertCellCount);

            UpdateAutoFitForCell(_CellCount, data);

            await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, false, ExcelValueTypes.Number, style);
            await _CurrentWorksheetPartWriter!.WriteAsync(string.Format(_Culture, "<v>{0:G}</v>", data)).ConfigureAwait(false);
            await WriteCellFooterAsync();

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public override async Task AddCellAsync(Int64 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);
            
            MergeNextCell(horzCellCount, vertCellCount);

            UpdateAutoFitForCell(_CellCount, data);

            await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, false, ExcelValueTypes.Number, style);
            await _CurrentWorksheetPartWriter!.WriteAsync(string.Format(_Culture, "<v>{0:G}</v>", data)).ConfigureAwait(false);
            await WriteCellFooterAsync();

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public override async Task AddCellAsync(UInt64 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);
            
            MergeNextCell(horzCellCount, vertCellCount);

            UpdateAutoFitForCell(_CellCount, data);

            await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, false, ExcelValueTypes.Number, style);
            await _CurrentWorksheetPartWriter!.WriteAsync(string.Format(_Culture, "<v>{0:G}</v>", data)).ConfigureAwait(false);
            await WriteCellFooterAsync();

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public override async Task AddCellAsync(float data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);
            
            MergeNextCell(horzCellCount, vertCellCount);

            UpdateAutoFitForCell(_CellCount, data);

            await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, false, ExcelValueTypes.Number, style);
            await _CurrentWorksheetPartWriter!.WriteAsync(string.Format(_Culture, "<v>{0:R15}</v>", data)).ConfigureAwait(false);
            await WriteCellFooterAsync();

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }

        public override async Task AddCellAsync(double data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);
            
            MergeNextCell(horzCellCount, vertCellCount);

            UpdateAutoFitForCell(_CellCount, data);

            await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, false, ExcelValueTypes.Number, style);
            await _CurrentWorksheetPartWriter!.WriteAsync(string.Format(_Culture, "<v>{0:R15}</v>", data)).ConfigureAwait(false);
            await WriteCellFooterAsync();

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }

        public override async Task AddCellAsync(decimal data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);
            
            MergeNextCell(horzCellCount, vertCellCount);

            UpdateAutoFitForCell(_CellCount, data);

            await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, false, ExcelValueTypes.Number, style);
            await _CurrentWorksheetPartWriter!.WriteAsync(string.Format(_Culture, "<v>{0:G15}</v>", data)).ConfigureAwait(false);
            await WriteCellFooterAsync();

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }

        public override async Task AddCellAsync(DateTime data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);
            
            MergeNextCell(horzCellCount, vertCellCount);

            var oaDate = data.ToOADate();

            UpdateAutoFitForCell(_CellCount, data);

            if (oaDate >= 0)
            {
                await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, false, ExcelValueTypes.Number, style);
                await _CurrentWorksheetPartWriter!.WriteAsync(string.Format(_Culture, "<v>{0:R15}</v>", oaDate)).ConfigureAwait(false);
                await WriteCellFooterAsync();
            }
            else
            {
                await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, true, null, style);
            }

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }

        public override async Task AddCellFormulaAsync(string formula, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);

            MergeNextCell(horzCellCount, vertCellCount);

            UpdateAutoFitForCell(_CellCount, formula);

            await WriteCellHeaderAsync(_CellCount++ + 1, _RowCount, false, ExcelValueTypes.FormulaString, style);
            await _CurrentWorksheetPartWriter!.WriteAsync("<f>" + _XmlWriterHelper!.EscapeValue(formula) + "</f>").ConfigureAwait(false);
            await WriteCellFooterAsync();

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }

        public async override Task AddCellImageAsync(
            Image image,
            Style? style = null,
            int horzCellCount = 0, 
            int vertCellCount = 0,
            CancellationToken cancellationToken = default)
        {
            if (_NextQueuedRowIndex == _RowCount && _CellCount > 0)
                await WriteMergedCellCounterparts(_RowCount, _CellCount);

            MergeNextCell(horzCellCount, vertCellCount);

            int vm = await _Package!.AddImageAsync(image, cancellationToken) + 1;

            UpdateAutoFitForCell(_CellCount, image);

            await WriteCellHeaderAsync(
                cellIndex: _CellCount++ + 1,
                rowIndex: _RowCount,
                closed: false,
                type: ExcelValueTypes.Error,
                style: style,
                valueMetadataIndex: vm);

            await _CurrentWorksheetPartWriter!.WriteAsync("<v>" + ExcelErrors.Value + "</v>").ConfigureAwait(false);
            await WriteCellFooterAsync();

            if (horzCellCount > 1 || vertCellCount > 1)
                await HandleMergedCells(horzCellCount, vertCellCount, style);
        }

        #endregion
    }
}
