using System;
using System.IO;
using System.Threading.Tasks;

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

        public abstract Task NewWorksheetAsync(WorksheetInfo info);

        public abstract Task SkipRowAsync();

        public abstract Task SkipRowsAsync(int count);

        public abstract Task AddRowAsync(Style style = null, float height = 0f, bool autoFit = true);

        public abstract Task FinishAsync();

        #endregion

        #region Cells

        public abstract Task SkipCellAsync();
        
        public abstract Task SkipCellsAsync(int count);
            
        public abstract Task AddCellAsync(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract Task AddCellStringAutoTypeAsync(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract Task AddCellForcedStringAsync(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract Task AddCellAsync(Int32 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public abstract Task AddCellAsync(UInt32 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public abstract Task AddCellAsync(Int64 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public abstract Task AddCellAsync(UInt64 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public abstract Task AddCellAsync(float data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract Task AddCellAsync(double data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract Task AddCellAsync(decimal data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public abstract Task AddCellAsync(DateTime data, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        public virtual async Task AddCellAsync(object data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            if (data is Int32)
                await AddCellAsync((Int32)data, style, horzCellCount, vertCellCount).ConfigureAwait(false);
            else if (data is Int64)
                await AddCellAsync((Int64)data, style, horzCellCount, vertCellCount).ConfigureAwait(false);
            else if (data is float)
                await AddCellAsync((float)data, style, horzCellCount, vertCellCount).ConfigureAwait(false);
            else if (data is double)
                await AddCellAsync((double)data, style, horzCellCount, vertCellCount).ConfigureAwait(false);
            else if (data is decimal)
                await AddCellAsync((decimal)data, style, horzCellCount, vertCellCount).ConfigureAwait(false);
            else if (data is DateTime)
                await AddCellAsync((DateTime)data, style, horzCellCount, vertCellCount).ConfigureAwait(false);
            else if (data is string)
                await AddCellAsync((string)data, style, horzCellCount, vertCellCount).ConfigureAwait(false);
            else if (data == null)
                await AddCellAsync("", style, horzCellCount, vertCellCount).ConfigureAwait(false);
            else
                await AddCellAsync(data.ToString(), style, horzCellCount, vertCellCount).ConfigureAwait(false);
        }

        public abstract Task AddCellFormulaAsync(string formula, Style style = null, int horzCellCount = 0, int vertCellCount = 0);

        #endregion
    }
}
