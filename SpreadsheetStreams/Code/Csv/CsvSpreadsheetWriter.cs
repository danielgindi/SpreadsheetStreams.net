using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

#nullable enable

namespace SpreadsheetStreams
{
    public class CsvSpreadsheetWriter : SpreadsheetWriter
    {
        #region Constructors

        public CsvSpreadsheetWriter(Stream outputStream, Encoding fileEncoding)
            : base(outputStream)
        {
            _FileEncoding = fileEncoding ?? Encoding.UTF8;
            _Writer = new StreamWriter(OutputStream, _FileEncoding);
        }

        public CsvSpreadsheetWriter(Stream outputStream)
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
            }
        }

        #endregion

        #region Private data

        private CultureInfo _Culture = CultureInfo.InvariantCulture;

        private Encoding _FileEncoding = Encoding.UTF8;
        private StreamWriter? _Writer = null;

        private bool _WroteFileStart = false;
        private bool _WroteFileEnd = false;
        private bool _ShouldEndWorksheet = false;
        private bool _ShouldEndRow = false;

        #endregion

        #region SpreadsheetWriter - Basic properties

        public override string FileExtension => "csv";
        public override string FileContentType => "application/vnd.ms-excel";
        public override bool IsFlatFormat => true;

        #endregion

        #region Public Properties

        public bool MultilineSupport { get; set; } = true;

        #endregion

        #region Private helpers

        private string CsvEscape(string value)
        {
            if (!MultilineSupport)
            {
                value = value.Replace('\n', ' ');
                value = value.Replace('\r', ' ');
            }

            value = value.Replace(@"""", @"""""");

            return value;
        }

        private async Task WriteAsync(string data)
        {
            await _Writer!.WriteAsync(data).ConfigureAwait(false);
        }

        #endregion

        #region Syling

        public override void RegisterStyle(Style style)
        {
        }

        #endregion

        #region SpreadsheetWriter - Document Lifespan (private)

        private async Task WritePendingEndRowAsync()
        {
            if (!_ShouldEndRow) return;

            await WriteAsync("\n");

            _ShouldEndRow = false;
        }

        private async Task WritePendingEndWorksheetAsync()
        {
            if (_ShouldEndWorksheet)
            {
                await WritePendingEndRowAsync();
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

        public override async Task NewWorksheetAsync(WorksheetInfo info)
        {
            bool shouldAddEmptyRow = _ShouldEndWorksheet;

            await WritePendingEndWorksheetAsync();
            _ShouldEndWorksheet = true;

            if (shouldAddEmptyRow) await AddRowAsync();

            if (info.Name != null)
            {
                await AddRowAsync();
                await AddCellAsync(info.Name);
                await AddRowAsync();
            }
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
                WriteBeginFile();
            }

            await WritePendingEndRowAsync();

            _ShouldEndRow = true;
        }

        public override async Task FinishAsync()
        {
            if (!_WroteFileStart)
            {
                WriteBeginFile();
                await NewWorksheetAsync(new WorksheetInfo { }).ConfigureAwait(false);
            }

            await WritePendingEndWorksheetAsync();

            if (!_WroteFileEnd)
            {
                _WroteFileEnd = true;

                _Writer!.Flush();
                _Writer.Close();
            }
        }

        #endregion

        #region SpreadsheetWriter - Cell methods

        private static char[] s_CharsForEscape = new char[] { '\n', '\r', '"', ',' };

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
            await WriteAsync(data == null ? "" : data.IndexOfAny(s_CharsForEscape) == -1 ? data + "," : string.Format(@"""{0}"",", CsvEscape(data)));
        }

        public override async Task AddCellStringAutoTypeAsync(string? data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            await WriteAsync(data == null ? "" : data.IndexOfAny(s_CharsForEscape) == -1 ? data + "," : string.Format(@"""{0}"",", CsvEscape(data)));
        }

        public override async Task AddCellForcedStringAsync(string? data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            await WriteAsync(data == null ? "" : string.Format("\"=\"\"{0}\"\"\",", CsvEscape(data)));
        }

        public override async Task AddCellAsync(Int32 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            await WriteAsync(string.Format(_Culture, "{0:G},", data));
        }

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public override async Task AddCellAsync(UInt32 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            await WriteAsync(string.Format(_Culture, "{0:G},", data));
        }
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public override async Task AddCellAsync(Int64 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            await WriteAsync(string.Format(_Culture, "{0:G},", data));
        }

#pragma warning disable CS3001 // Argument type is not CLS-compliant
        public override async Task AddCellAsync(UInt64 data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            await WriteAsync(string.Format(_Culture, "{0:G},", data));
        }
#pragma warning restore CS3001 // Argument type is not CLS-compliant

        public override async Task AddCellAsync(float data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            await WriteAsync(string.Format(_Culture, "{0:G29},", data));
        }

        public override async Task AddCellAsync(double data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            await WriteAsync(string.Format(_Culture, "{0:G29},", data));
        }

        public override async Task AddCellAsync(decimal data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            await WriteAsync(string.Format(_Culture, "{0:G29},", data));
        }

        public override async Task AddCellAsync(DateTime data, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            var dateFormat = "yyyy-MM-ddTHH:mm:ss.fff";

            await WriteAsync(string.Format("{0},", data.Year <= 1 ? "" : data.ToString(dateFormat)));
        }

        public override async Task AddCellFormulaAsync(string formula, Style? style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            await WriteAsync(string.Format(@"""{0}"",", CsvEscape(formula)));
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
