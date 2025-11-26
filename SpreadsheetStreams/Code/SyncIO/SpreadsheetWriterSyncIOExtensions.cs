using System;

namespace SpreadsheetStreams.SyncIO
{
    public class SpreadsheetWriterSyncIOExtensions
    {
        #region Document Lifespan (public)

        public static void NewWorksheet(SpreadsheetWriter writer, WorksheetInfo info)
        {
            writer.NewWorksheetAsync(info).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        public static void AddRow(SpreadsheetWriter writer, Style style = null, float height = 0f, bool autoFit = true)
        {
            writer.AddRowAsync(style, height).ConfigureAwait(false).GetAwaiter().GetResult();
        }
        
        public static void Finish(SpreadsheetWriter writer)
        {
            writer.FinishAsync().ConfigureAwait(false).GetAwaiter().GetResult();
        }

        #endregion
        
        #region Cells

        public static void AddCell(SpreadsheetWriter writer, string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        public static void AddCellStringAutoType(SpreadsheetWriter writer, string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellStringAutoTypeAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        public static void AddCellForcedString(SpreadsheetWriter writer, string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellForcedStringAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        public static void AddCell(SpreadsheetWriter writer, Int32 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public static void AddCell(SpreadsheetWriter writer, UInt32 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public static void AddCell(SpreadsheetWriter writer, Int64 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public static void AddCell(SpreadsheetWriter writer, UInt64 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public static void AddCell(SpreadsheetWriter writer, float data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        public static void AddCell(SpreadsheetWriter writer, double data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        public static void AddCell(SpreadsheetWriter writer, decimal data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        public static void AddCell(SpreadsheetWriter writer, DateTime data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        public static void AddCell(SpreadsheetWriter writer, object data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellAsync(data, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        public static void AddCellFormula(SpreadsheetWriter writer, string formula, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            writer.AddCellFormulaAsync(formula, style, horzCellCount, vertCellCount).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        #endregion
    }
}
