using System.IO.Compression;
using System.Xml.Linq;

namespace SpreadsheetStreams.ExcelTests
{
    public class RowSkippingTests
    {
        [Fact]
        public async Task WhenSkippingRowsWithoutWorksheet_ShouldThrow()
        {
            await Assert.ThrowsAsync<InvalidOperationException>(async () =>
            {
                using var ms = new MemoryStream();
                var writer = new ExcelSpreadsheetWriter(ms);

                await writer.SkipRowAsync();
            });
        }

        [Fact]
        public async Task TestSkippingRowsOnBeginning()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            await writer.NewWorksheetAsync(new WorksheetInfo()
            {
                Name = "Test Worksheet"
            });

            await writer.SkipRowsAsync(5);

            await writer.AddRowAsync();
            await writer.AddCellAsync("Test Cell");

            await writer.FinishAsync();

            ms.Position = 0;
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read, true);

            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            using var sheetStream = sheetEntry.Open();

            var doc = XDocument.Load(sheetStream);
            var rootNamespace = doc.Root!.Name.Namespace;

            var row = doc.Descendants(rootNamespace! + "row")
                .FirstOrDefault();

            Assert.Equal("6", row?.Attribute("r")?.Value);
        }

        [Fact]
        public async Task TestSkippingRowsAfterExistingRow()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            await writer.NewWorksheetAsync(new WorksheetInfo()
            {
                Name = "Test Worksheet"
            });

            await writer.AddRowAsync();
            await writer.AddCellAsync("Test Cell");

            await writer.SkipRowsAsync(5);

            await writer.AddRowAsync();
            await writer.AddCellAsync("Test Cell");

            await writer.FinishAsync();

            ms.Position = 0;
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read, true);

            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            using var sheetStream = sheetEntry.Open();

            var doc = XDocument.Load(sheetStream);
            var rootNamespace = doc.Root!.Name.Namespace;

            var row = doc.Descendants(rootNamespace! + "row")
                .LastOrDefault();

            Assert.Equal("7", row?.Attribute("r")?.Value);
        }
    }
}
