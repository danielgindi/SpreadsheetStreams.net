using System.IO.Compression;
using System.Xml.Linq;

namespace SpreadsheetStreams.ExcelTests
{
    public class CellSkippingTests
    {
        [Fact]
        public async Task TestSkippingCells()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            await writer.NewWorksheetAsync(new WorksheetInfo()
            {
                Name = "Test Worksheet"
            });

            await writer.AddRowAsync();
            await writer.AddCellAsync("Test Cell");

            await writer.SkipCellsAsync(5);

            await writer.AddCellAsync("New Cell");

            await writer.FinishAsync();

            ms.Position = 0;
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read, true);

            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            using var sheetStream = sheetEntry.Open();

            var doc = XDocument.Load(sheetStream);
            var rootNamespace = doc.Root!.Name.Namespace;

            var cell = doc.Descendants(rootNamespace! + "c")
                .LastOrDefault();

            Assert.Equal("G1", cell?.Attribute("r")?.Value);
        }
    }
}
