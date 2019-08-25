using System;
using System.IO;
using System.Text;
using System.Xml;

namespace SpreadsheetStreams.Util
{
    internal class XmlWriterHelper : IDisposable
    {
        private static XmlWriterSettings XmlWriterSettings = new XmlWriterSettings
        {
            ConformanceLevel = ConformanceLevel.Fragment,
        };

        private StringBuilder _Sb;
        private StringWriter _StringWriter;
        private XmlWriter _XmlWriter;

        internal XmlWriterHelper()
        {
            _Sb = new StringBuilder();
            _StringWriter = new StringWriter(_Sb);
            _XmlWriter = XmlWriter.Create(_StringWriter, XmlWriterSettings);
        }

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
                if (_XmlWriter != null)
                {
                    _XmlWriter.Dispose();
                    _XmlWriter = null;
                }

                if (_StringWriter != null)
                {
                    _StringWriter.Dispose();
                    _StringWriter = null;
                }
            }
        }

        #endregion

        internal string EscapeAttribute(string value)
        {
            if (value == null) return "";

            _Sb.Clear();

            _XmlWriter.WriteStartElement("e");
            _XmlWriter.WriteAttributeString("_", value);
            _XmlWriter.WriteEndElement();
            _XmlWriter.Flush();
            var result = _Sb.ToString();
            return result.Substring(6, result.Length - 10);
        }

        internal string EscapeValue(string value)
        {
            if (value == null) return "";

            _Sb.Clear();

            _XmlWriter.WriteString(value);
            _XmlWriter.Flush();

            return _Sb.ToString();
        }
    }
}
