using System.IO.Compression;
using System.Xml.Linq;

namespace SpreadsheetStreams.ExcelTests
{
    public class RowSkippingTests
    {
        [Fact]
        public void WhenSkippingRowsWithoutWorksheet_ShouldThrow()
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                using var ms = new MemoryStream();
                var writer = new ExcelSpreadsheetWriter(ms);

                writer.SkipRow();
            });
        }

        [Fact]
        public void TestSkippingRowsOnBeginning()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            writer.NewWorksheet(new WorksheetInfo()
            {
                Name = "Test Worksheet"
            });

            writer.SkipRows(5);

            writer.AddRow();
            writer.AddCell("Test Cell");

            writer.Finish();

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
        public void TestSkippingRowsAfterExistingRow()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            writer.NewWorksheet(new WorksheetInfo()
            {
                Name = "Test Worksheet"
            });

            writer.AddRow();
            writer.AddCell("Test Cell");

            writer.SkipRows(5);

            writer.AddRow();
            writer.AddCell("Test Cell");

            writer.Finish();

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
