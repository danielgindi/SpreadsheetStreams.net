using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace SpreadsheetStreams.Code.Excel
{
    internal class RichValueStructureStore
    {
        private List<RichValueStructure> Structures = new List<RichValueStructure>();
        private Dictionary<RichValueStructure, int> StructureIndexMap = new Dictionary<RichValueStructure, int>();

        internal const string PART_PATH = "/xl/richData/rdrichvaluestructure.xml";

        internal RichValueStructureStore()
        {
        }

        internal int AddStructure(RichValueStructure structure)
        {
            if (StructureIndexMap.TryGetValue(structure, out var index))
                return index;

            index = Structures.Count;
            Structures.Add(structure);
            StructureIndexMap[structure] = index;

            return index;
        }

        internal async Task WriteXmlToStream(StreamWriter writer)
        {
            await writer.WriteAsync("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>").ConfigureAwait(false);
            await writer.WriteAsync($"<rvStructures xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" count=\"{Structures.Count}\">").ConfigureAwait(false);
            foreach (var structure in Structures)
            {
                await structure.WriteXmlToStream(writer);
            }
            await writer.WriteAsync($"</rvStructures>");
        }
    }
}
