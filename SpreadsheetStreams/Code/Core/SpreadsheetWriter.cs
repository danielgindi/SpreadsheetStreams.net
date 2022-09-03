using System;
using System.IO;

[assembly: CLSCompliant(true)]

namespace SpreadsheetStreams
{
    public abstract class SpreadsheetWriter : IDisposable
    {
        #region Private Variables

        internal Stream OutputStream = null;

        #endregion

        #region Constructors

        protected SpreadsheetWriter(Stream outputStream)
        {
            OutputStream = outputStream;
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                Finish();
            }
        }

        #endregion

        #region Basic properties

        public abstract string FileExtension { get; }
        public abstract string FileContentType { get; }
        public abstract bool IsFlatFormat { get; }

        public SpreadsheetInfo SpreadsheetInfo = new SpreadsheetInfo();

        #endregion

        #region Syling

        public abstract void RegisterStyle(Style style);

        #endregion
        
        #region Document Lifespan (public)

        public abstract void NewWorksheet(WorksheetInfo info);

        public abstract void SkipRow();

        public abstract void SkipRows(int count);

        public abstract void AddRow(Style style = null, float height = 0f, bool autoFit = true);

        public abstract void Finish();

        #endregion

        #region Cells

        public abstract void SkipCell();
        public abstract void SkipCells(int count);
            
        public abstract void AddCell(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract void AddCellStringAutoType(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract void AddCellForcedString(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract void AddCell(Int32 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public abstract void AddCell(UInt32 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public abstract void AddCell(Int64 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public abstract void AddCell(UInt64 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public abstract void AddCell(float data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract void AddCell(double data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract void AddCell(decimal data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract void AddCell(DateTime data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public virtual void AddCell(object data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (data is Int32)
                AddCell((Int32)data, style, horzCellCount, vertCellCount);
            else if (data is Int64)
                AddCell((Int64)data, style, horzCellCount, vertCellCount);
            else if (data is float)
                AddCell((float)data, style, horzCellCount, vertCellCount);
            else if (data is double)
                AddCell((double)data, style, horzCellCount, vertCellCount);
            else if (data is decimal)
                AddCell((decimal)data, style, horzCellCount, vertCellCount);
            else if (data is DateTime)
                AddCell((DateTime)data, style, horzCellCount, vertCellCount);
            else if (data is string)
                AddCell((string)data, style, horzCellCount, vertCellCount);
            else
                AddCell(data.ToString(), style, horzCellCount, vertCellCount);
        }

        public abstract void AddCellFormula(string formula, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        #endregion
    }
}
