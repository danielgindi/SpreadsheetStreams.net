using System;
using System.Globalization;
using System.IO;
using System.Text;

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

        public CsvSpreadsheetWriter() : this(null, null)
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
        private bool _ShouldEndWorksheet = false;
        private bool _ShouldEndRow = false;

        #endregion

        #region SpreadsheetWriter - Basic properties

        public override string FileExtension => "csv";
        public override string FileContentType => "application/vnd.ms-excel";
        public override bool IsFlatFormat => true;
        public bool ShouldWriteBOM { get; set; } = true;

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

        private void Write(string data)
        {
            _Writer.Write(data);
        }

        #endregion

        #region Syling

        public override void RegisterStyle(Style style)
        {
        }

        #endregion

        #region SpreadsheetWriter - Document Lifespan (private)

        private void WritePendingEndRow()
        {
            if (!_ShouldEndRow) return;

            Write("\n");

            _ShouldEndRow = false;
        }

        private void WritePendingEndWorksheet()
        {
            if (_ShouldEndWorksheet)
            {
                WritePendingEndRow();
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
        }

        #endregion

        #region SpreadsheetWriter - Document Lifespan (public)

        public override void NewWorksheet(WorksheetInfo info)
        {
            bool shouldAddEmptyRow = _ShouldEndWorksheet;

            WritePendingEndWorksheet();
            _ShouldEndWorksheet = true;

            if (shouldAddEmptyRow) AddRow();

            if (info.Name != null)
            {
                AddRow();
                AddCell(info.Name);
                AddRow();
            }
        }

        public override void AddRow(Style style = null, float height = 0f)
        {
            if (!_ShouldEndWorksheet)
            {
                throw new InvalidOperationException("Adding new rows is not allowed at this time. Please call NewWorksheet(...) first.");
            }

            if (!_WroteFileStart)
            {
                WriteBeginFile();
            }

            WritePendingEndRow();

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
                _WroteFileEnd = true;

                _Writer.Flush();
                _Writer.Close();
            }
        }

        #endregion

        #region SpreadsheetWriter - Cell methods

        public override void AddCell(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            Write(string.Format(@"""{0}"",", CsvEscape(data)));
        }

        public override void AddCellStringAutoType(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            Write(string.Format(@"""{0}"",", CsvEscape(data)));
        }

        public override void AddCellForcedString(string data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            Write(string.Format("\"=\"\"{0}\"\"\",", CsvEscape(data)));
        }

        public override void AddCell(Int32 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            Write(string.Format(_Culture, "{0:G},", data));
        }

        public override void AddCell(Int64 data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            Write(string.Format(_Culture, "{0:G},", data));
        }

        public override void AddCell(float data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            Write(string.Format(_Culture, "{0:G},", data));
        }

        public override void AddCell(double data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            Write(string.Format(_Culture, "{0:G},", data));
        }

        public override void AddCell(decimal data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            Write(string.Format(_Culture, "{0:G},", data));
        }

        public override void AddCell(DateTime data, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            var dateFormat = "yyyy-MM-ddTHH:mm:ss.fff";

            Write(string.Format("{0},", data.Year <= 1 ? "" : data.ToString(dateFormat)));
        }

        public override void AddCellFormula(string formula, Style style = null, int horzCellCount = 0, int vertCellCount = 0)
        {
            Write(string.Format(@"""{0}"",", CsvEscape(formula)));
        }

        #endregion
    }
}
