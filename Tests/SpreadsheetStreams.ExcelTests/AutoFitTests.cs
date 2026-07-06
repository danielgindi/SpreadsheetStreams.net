using SpreadsheetStreams.Code.Excel;
using System.IO.Compression;
using System.Xml.Linq;

namespace SpreadsheetStreams.ExcelTests
{
    public class AutoFitTests
    {
        [Fact]
        public async Task AutoFitColumns_ShouldRewriteWorksheetXml()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            await writer.NewWorksheetAsync(new WorksheetInfo
            {
                Name = "Test Worksheet"
            });

            writer.EnableAutoFitForColumn(0, new AutoFitConfig
            {
                Measure = (_, _) => 42f,
                MaxLength = 100f,
                Multiplier = 1f,
            });

            await writer.AddRowAsync();
            await writer.AddCellAsync("wide value");
            await writer.FinishAsync();

            ms.Position = 0;
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read, true);

            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            using var sheetStream = sheetEntry.Open();

            var doc = XDocument.Load(sheetStream);
            var rootNamespace = doc.Root!.Name.Namespace;

            var cols = doc.Root.Element(rootNamespace + "cols");
            Assert.NotNull(cols);

            var col = cols!.Element(rootNamespace + "col");
            Assert.NotNull(col);
            Assert.Equal("1", col!.Attribute("min")?.Value);
            Assert.Equal("1", col.Attribute("max")?.Value);
            Assert.Equal("42", col.Attribute("width")?.Value);
            Assert.Equal("1", col.Attribute("customWidth")?.Value);

            Assert.NotNull(doc.Root.Element(rootNamespace + "sheetData"));
            Assert.Equal("wide value", doc.Descendants(rootNamespace + "c").Single().Value);
        }
    }
}
