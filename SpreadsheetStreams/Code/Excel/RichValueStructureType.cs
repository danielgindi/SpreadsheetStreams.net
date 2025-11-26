using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetStreams.Code.Excel
{
    internal enum RichValueStructureType
    {
        Error,
        LocalImage,
        WebImage,
        ImageUrl,
        LinkedEntity,
        LinkedEntity2,
        LinkedEntityCore,
        LinkedEntity2Core,
        FormattedNumber,
        Hyperlink,
        Array,
        Entity,
        StockHistoryCache,
        ExternalCodeServiceObject,
        SourceAttribution
    }
}
