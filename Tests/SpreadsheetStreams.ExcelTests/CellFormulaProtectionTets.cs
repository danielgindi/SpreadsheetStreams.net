using System.IO.Compression;
using System.Xml.Linq;

namespace SpreadsheetStreams.ExcelTests
{
    public class CellFormulaProtectionTets
    {
        [Fact]
        public async Task TestFormulaCells()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);
            writer.EnableFormulaProtection = true;

            await writer.NewWorksheetAsync(new WorksheetInfo()
            {
                Name = "Test Worksheet"
            });

            await writer.AddRowAsync();
            await writer.AddCellAsync("=ABCD");
            await writer.AddCellAsync("ABCD");
            await writer.FinishAsync();

            ms.Position = 0;
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read, true);

            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            using var sheetStream = sheetEntry.Open();

            var doc = XDocument.Load(sheetStream);
            var rootNamespace = doc.Root!.Name.Namespace;

            var rows = doc.Descendants(rootNamespace! + "row");
            var row1 = rows.ElementAt(0);
            var cells = row1.Descendants(rootNamespace! + "c");

            Assert.Equal("'=ABCD", cells.ElementAt(0).Value);
            Assert.Equal("ABCD", cells.ElementAt(1).Value);
        }

        [Fact]
        public async Task TestFormulaCellsNoProtection()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);
            writer.EnableFormulaProtection = false;

            await writer.NewWorksheetAsync(new WorksheetInfo()
            {
                Name = "Test Worksheet"
            });

            await writer.AddRowAsync();
            await writer.AddCellAsync("=ABCD");
            await writer.AddCellAsync("ABCD");
            await writer.FinishAsync();

            ms.Position = 0;
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read, true);

            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            using var sheetStream = sheetEntry.Open();

            var doc = XDocument.Load(sheetStream);
            var rootNamespace = doc.Root!.Name.Namespace;

            var rows = doc.Descendants(rootNamespace! + "row");
            var row1 = rows.ElementAt(0);
            var cells = row1.Descendants(rootNamespace! + "c");

            Assert.Equal("=ABCD", cells.ElementAt(0).Value);
            Assert.Equal("ABCD", cells.ElementAt(1).Value);
        }
    }
}
