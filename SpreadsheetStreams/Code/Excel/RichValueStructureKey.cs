using SpreadsheetStreams.Util;
using System.IO;
using System.Threading.Tasks;

namespace SpreadsheetStreams.Code.Excel
{
    internal class RichValueStructureKey
    {
        public string Name { get; private set; }
        public RichValueDataType Type { get; private set; }

        public RichValueStructureKey(string name, RichValueDataType type)
        {
            this.Name = name;
            this.Type = type;
        }

        public Task WriteXmlToStream(StreamWriter writer)
        {
            if (Type == RichValueDataType.Decimal)
                return writer.WriteAsync($"<k n=\"{XmlWriterHelper.EscapeSimpleXmlAttr(Name)}\"/>");
            return writer.WriteAsync($"<k n=\"{XmlWriterHelper.EscapeSimpleXmlAttr(Name)}\" t=\"{GetTypeString(Type)}\"/>");
        }

        internal static string GetTypeString(RichValueDataType type)
        {
            switch (type)
            {
                case RichValueDataType.Decimal:
                    return "d";
                case RichValueDataType.Integer:
                    return "i";
                case RichValueDataType.Bool:
                    return "b";
                case RichValueDataType.Error:
                    return "e";
                case RichValueDataType.String:
                    return "s";
                case RichValueDataType.RichValue:
                    return "r";
                case RichValueDataType.Array:
                    return "a";
                case RichValueDataType.SupportingPropertyBag:
                    return "spb";
                case RichValueDataType.SupportingPropertyBagArray:
                    return "spba";

                default:
                    return "i";
            }
        }
    }
}
