using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

#nullable enable

namespace SpreadsheetStreams.Util
{
    internal class XmlWriterHelper : IDisposable
    {
        private static XmlWriterSettings XmlWriterSettings = new XmlWriterSettings
        {
            ConformanceLevel = ConformanceLevel.Fragment,
            CheckCharacters = false,
        };

        private StringBuilder _Sb;
        private StringWriter? _StringWriter;
        private XmlWriter? _XmlWriter;

        // filters control characters but allows only properly-formed surrogate sequences
        // Credit to Jeff Atwood: https://stackoverflow.com/questions/397250/unicode-regex-invalid-xml-characters/961504#961504
        // I've checked this, seems valid.
        private static Regex _InvalidXmlCharRegex = new Regex(
            @"(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFEFF\uFFFE\uFFFF]",
            RegexOptions.Compiled);

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

        internal string EscapeAttribute(string? value, bool removeInvalidChars = true)
        {
            if (value == null) return "";

            if (removeInvalidChars)
                value = RemoveInvalidXmlChars(value);

            _Sb.Clear();

            _XmlWriter!.WriteStartElement("e");
            _XmlWriter.WriteAttributeString("_", value);
            _XmlWriter.WriteEndElement();
            _XmlWriter.Flush();
            var result = _Sb.ToString();
            return result.Substring(6, result.Length - 10);
        }

        internal string EscapeValue(string? value, bool removeInvalidChars = true)
        {
            if (value == null) return "";

            if (removeInvalidChars)
                value = RemoveInvalidXmlChars(value);

            _Sb.Clear();

            _XmlWriter!.WriteString(value);
            _XmlWriter.Flush();

            return _Sb.ToString();
        }

        public static string RemoveInvalidXmlChars(string? text, string replacement = "�")
        {
            if (text == null || text.Length == 0) return "";
            return _InvalidXmlCharRegex.Replace(text, replacement);
        }

        public static string EscapeSimpleXmlAttr(string s)
        {
            return s.Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }
    }
}
