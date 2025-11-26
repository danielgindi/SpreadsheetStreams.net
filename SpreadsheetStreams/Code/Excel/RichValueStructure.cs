using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#nullable enable

namespace SpreadsheetStreams.Code.Excel
{
    internal class RichValueStructure
    {
        public RichValueStructure(
            RichValueStructureType structureType,
            List<RichValueStructureKey> keys)
        {
            StructureType = structureType;
            Keys = keys;
        }

        public RichValueStructureType StructureType { get; private set; }

        public List<RichValueStructureKey> Keys { get; private set; }

        internal async Task WriteXmlToStream(StreamWriter writer)
        {
            await writer.WriteAsync($"<s t=\"{StructureTypeToString(StructureType)}\">").ConfigureAwait(false);
            foreach (var key in Keys)
            {
                await key.WriteXmlToStream(writer);
            }
            await writer.WriteAsync("</s>").ConfigureAwait(false);
        }

        private static readonly Dictionary<string, RichValueStructureType> _richValueFromString =
            new Dictionary<string, RichValueStructureType>(StringComparer.OrdinalIgnoreCase)
            {
                    { "_error", RichValueStructureType.Error },
                    { "_localImage", RichValueStructureType.LocalImage },
                    { "_webimage", RichValueStructureType.WebImage },
                    { "_imageurl", RichValueStructureType.ImageUrl },
                    { "_linkedentity", RichValueStructureType.LinkedEntity },
                    { "_linkedentity2", RichValueStructureType.LinkedEntity2 },
                    { "_linkedentitycore", RichValueStructureType.LinkedEntityCore },
                    { "_linkedentity2core", RichValueStructureType.LinkedEntity2Core },
                    { "_formattednumber", RichValueStructureType.FormattedNumber },
                    { "_hyperlink", RichValueStructureType.Hyperlink },
                    { "_array", RichValueStructureType.Array },
                    { "_entity", RichValueStructureType.Entity },
                    { "_stockhistorycache", RichValueStructureType.StockHistoryCache },
                    { "_python", RichValueStructureType.ExternalCodeServiceObject },
                    { "_sourceattribution", RichValueStructureType.SourceAttribution }
            };

        private static readonly Dictionary<RichValueStructureType, string> _richValueToString =
            _richValueFromString.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

        public static RichValueStructureType? StructureTypeFromString(string? value)
        {
            if (value == null) return null;
            if (_richValueFromString.TryGetValue(value, out RichValueStructureType result))
                return result;
            return null;
        }

        public static string StructureTypeToString(RichValueStructureType value)
        {
            return _richValueToString[value];
        }
    }
}
