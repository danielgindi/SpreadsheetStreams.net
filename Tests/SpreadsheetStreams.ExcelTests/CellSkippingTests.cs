using System.IO.Compression;
using System.Xml.Linq;

namespace SpreadsheetStreams.ExcelTests
{
    public class CellSkippingTests
    {
        [Fact]
        public void TestSkippingCells()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            writer.NewWorksheet(new WorksheetInfo()
            {
                Name = "Test Worksheet"
            });

            writer.AddRow();
            writer.AddCell("Test Cell");

            writer.SkipCells(5);

            writer.AddCell("New Cell");

            writer.Finish();

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
