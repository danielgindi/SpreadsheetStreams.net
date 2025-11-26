using System.IO;

#nullable enable

namespace SpreadsheetStreams
{
    public class Image
    {
        public string? ContentType;
        public string? Path;
        public Stream? Stream;
        public byte[]? Data;
    }
}