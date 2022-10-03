using System.Drawing;
using System.IO.Compression;
using System.Xml.Linq;

namespace SpreadsheetStreams.ExcelTests
{
    public class MergedCellBorderStyleTests
    {
        [Fact]
        public async Task TestMergedCellsGhostOrdering()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            var styleWithBorders = new Style
            {
                Borders = new List<Border>() {
                    new Border(BorderPosition.Left, Color.Black, BorderLineStyle.Continuous, 1.0f),
                    new Border(BorderPosition.Right, Color.Black, BorderLineStyle.Continuous, 1.0f),
                    new Border(BorderPosition.Top, Color.Black, BorderLineStyle.Continuous, 1.0f),
                    new Border(BorderPosition.Bottom, Color.Black, BorderLineStyle.Continuous, 1.0f),
                },
            };

            writer.RegisterStyle(styleWithBorders);

            await writer.NewWorksheetAsync(new WorksheetInfo()
            {
                Name = "Test Worksheet"
            });

            await writer.AddRowAsync();
            await writer.AddCellAsync("a1");
            await writer.AddCellAsync("merged b1:d1", styleWithBorders, 3);
            await writer.AddCellAsync("e1");

            await writer.AddRowAsync();
            await writer.AddCellAsync("a2");
            await writer.AddCellAsync("b2");
            await writer.AddCellAsync("merged c2:e4", styleWithBorders, 3, 3);
            await writer.AddCellAsync("f2");
            await writer.AddCellAsync("g2");

            await writer.SkipRowAsync();

            await writer.AddRowAsync();
            await writer.AddCellAsync("merged a4:b4", styleWithBorders, 2);
            await writer.SkipCellsAsync(3); // skip c4:e4, which are merged
            await writer.AddCellAsync("f4");
            await writer.AddCellAsync("g4");

            await writer.FinishAsync();

            ms.Position = 0;
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read, true);

            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            using var sheetStream = sheetEntry.Open();

            var doc = XDocument.Load(sheetStream);
            var rootNamespace = doc.Root!.Name.Namespace;

            var rows = doc.Descendants(rootNamespace! + "row");

            var row1 = rows.ElementAt(0);
            Assert.Equal("1", row1.Attribute("r")?.Value);
            var cells = row1.Descendants(rootNamespace! + "c");
            Assert.Equal("A1", cells.ElementAt(0).Attribute("r")?.Value);
            Assert.Equal("B1", cells.ElementAt(1).Attribute("r")?.Value);
            Assert.Equal("C1", cells.ElementAt(2).Attribute("r")?.Value);
            Assert.Equal("D1", cells.ElementAt(3).Attribute("r")?.Value);
            Assert.Equal("E1", cells.ElementAt(4).Attribute("r")?.Value);

            var row2 = rows.ElementAt(1);
            Assert.Equal("2", row2.Attribute("r")?.Value);
            cells = row2.Descendants(rootNamespace! + "c");
            Assert.Equal("A2", cells.ElementAt(0).Attribute("r")?.Value);
            Assert.Equal("B2", cells.ElementAt(1).Attribute("r")?.Value);
            Assert.Equal("C2", cells.ElementAt(2).Attribute("r")?.Value);
            Assert.Equal("D2", cells.ElementAt(3).Attribute("r")?.Value);
            Assert.Equal("E2", cells.ElementAt(4).Attribute("r")?.Value);
            Assert.Equal("F2", cells.ElementAt(5).Attribute("r")?.Value);
            Assert.Equal("G2", cells.ElementAt(6).Attribute("r")?.Value);

            var row3 = rows.ElementAt(2);
            Assert.Equal("3", row3.Attribute("r")?.Value);
            cells = row3.Descendants(rootNamespace! + "c");
            Assert.Equal("C3", cells.ElementAt(0).Attribute("r")?.Value);
            Assert.Equal("D3", cells.ElementAt(1).Attribute("r")?.Value);
            Assert.Equal("E3", cells.ElementAt(2).Attribute("r")?.Value);

            var row4 = rows.ElementAt(3);
            Assert.Equal("4", row4.Attribute("r")?.Value);
            cells = row4.Descendants(rootNamespace! + "c");
            Assert.Equal("A4", cells.ElementAt(0).Attribute("r")?.Value);
            Assert.Equal("B4", cells.ElementAt(1).Attribute("r")?.Value);
            Assert.Equal("C4", cells.ElementAt(2).Attribute("r")?.Value);
            Assert.Equal("D4", cells.ElementAt(3).Attribute("r")?.Value);
            Assert.Equal("E4", cells.ElementAt(4).Attribute("r")?.Value);
            Assert.Equal("F4", cells.ElementAt(5).Attribute("r")?.Value);
            Assert.Equal("G4", cells.ElementAt(6).Attribute("r")?.Value);
        }

        [Fact]
        public async Task TestMergedCellsWithoutGhosts()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            var styleWithoutBorders = new Style
            {
                Borders = new List<Border>() 
                {
                },
            };

            writer.RegisterStyle(styleWithoutBorders);

            await writer.NewWorksheetAsync(new WorksheetInfo()
            {
                Name = "Test Worksheet"
            });

            await writer.AddRowAsync();
            await writer.AddCellAsync("a1");
            await writer.AddCellAsync("merged b1:d1", styleWithoutBorders, 3);
            await writer.AddCellAsync("e1");

            await writer.AddRowAsync();
            await writer.AddCellAsync("a2");
            await writer.AddCellAsync("b2");
            await writer.AddCellAsync("merged c2:e4", styleWithoutBorders, 3, 3);
            await writer.AddCellAsync("f2");
            await writer.AddCellAsync("g2");

            await writer.SkipRowAsync();

            await writer.AddRowAsync();
            await writer.AddCellAsync("merged a4:b4", styleWithoutBorders, 2);
            await writer.SkipCellsAsync(3); // skip c4:e4, which are merged
            await writer.AddCellAsync("f4");
            await writer.AddCellAsync("g4");

            await writer.FinishAsync();

            ms.Position = 0;
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read, true);

            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            using var sheetStream = sheetEntry.Open();

            var doc = XDocument.Load(sheetStream);
            var rootNamespace = doc.Root!.Name.Namespace;

            var rows = doc.Descendants(rootNamespace! + "row");

            var row1 = rows.ElementAt(0);
            Assert.Equal("1", row1.Attribute("r")?.Value);
            var cells = row1.Descendants(rootNamespace! + "c");
            Assert.Equal("A1", cells.ElementAt(0).Attribute("r")?.Value);
            Assert.Equal("B1", cells.ElementAt(1).Attribute("r")?.Value);
            Assert.Equal("E1", cells.ElementAt(2).Attribute("r")?.Value);

            var row2 = rows.ElementAt(1);
            Assert.Equal("2", row2.Attribute("r")?.Value);
            cells = row2.Descendants(rootNamespace! + "c");
            Assert.Equal("A2", cells.ElementAt(0).Attribute("r")?.Value);
            Assert.Equal("B2", cells.ElementAt(1).Attribute("r")?.Value);
            Assert.Equal("C2", cells.ElementAt(2).Attribute("r")?.Value);
            Assert.Equal("F2", cells.ElementAt(3).Attribute("r")?.Value);
            Assert.Equal("G2", cells.ElementAt(4).Attribute("r")?.Value);

            var row4 = rows.ElementAt(2);
            Assert.Equal("4", row4.Attribute("r")?.Value);
            cells = row4.Descendants(rootNamespace! + "c");
            Assert.Equal("A4", cells.ElementAt(0).Attribute("r")?.Value);
            Assert.Equal("F4", cells.ElementAt(1).Attribute("r")?.Value);
            Assert.Equal("G4", cells.ElementAt(2).Attribute("r")?.Value);
        }
    }
}
