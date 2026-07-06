using System.IO.Compression;
using System.Xml.Linq;

namespace SpreadsheetStreams.ExcelTests
{
    public class ExcelPackageBehaviorTests
    {
        [Fact]
        public async Task FinishWithoutExplicitWorksheet_ShouldCreateDefaultWorksheet()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            await writer.FinishAsync();

            var workbook = ReadPackageXml(ms, "xl/workbook.xml");
            var workbookNamespace = workbook.Root!.Name.Namespace;
            var sheet = workbook.Descendants(workbookNamespace + "sheet").Single();

            Assert.Equal("Worksheet1", sheet.Attribute("name")?.Value);
            Assert.Equal("1", sheet.Attribute("sheetId")?.Value);

            var worksheet = ReadPackageXml(ms, "xl/worksheets/sheet1.xml");
            var worksheetNamespace = worksheet.Root!.Name.Namespace;

            Assert.NotNull(worksheet.Root.Element(worksheetNamespace + "sheetData"));
        }

        [Fact]
        public async Task WorksheetNames_ShouldBeSanitizedForWorkbookXml()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            await writer.NewWorksheetAsync(new WorksheetInfo
            {
                Name = "'Bad:/?[]Sheet'"
            });
            await writer.AddRowAsync();
            await writer.AddCellAsync("value");
            await writer.FinishAsync();

            var workbook = ReadPackageXml(ms, "xl/workbook.xml");
            var workbookNamespace = workbook.Root!.Name.Namespace;
            var sheet = workbook.Descendants(workbookNamespace + "sheet").Single();

            Assert.Equal("Bad     Sheet", sheet.Attribute("name")?.Value);
        }

        [Fact]
        public async Task WorksheetOptions_ShouldWriteSheetViewFormatAndColumnXml()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);

            await writer.NewWorksheetAsync(new WorksheetInfo
            {
                Name = "Options",
                RightToLeft = true,
                DefaultRowHeight = 20f,
                DefaultColumnWidth = 12f,
                ColumnInfos = new List<ColumnInfo>
                {
                    new ColumnInfo { FromColumn = 0, ToColumn = 2, Width = 25f, Hidden = true },
                    new ColumnInfo { FromColumn = 4, ToColumn = 4, Hidden = true },
                }
            });
            await writer.AddRowAsync();
            await writer.AddCellAsync("value");
            await writer.FinishAsync();

            var worksheet = ReadPackageXml(ms, "xl/worksheets/sheet1.xml");
            var worksheetNamespace = worksheet.Root!.Name.Namespace;

            var sheetView = worksheet.Descendants(worksheetNamespace + "sheetView").Single();
            Assert.Equal("1", sheetView.Attribute("rightToLeft")?.Value);

            var sheetFormat = worksheet.Root.Element(worksheetNamespace + "sheetFormatPr");
            Assert.Equal("20", sheetFormat?.Attribute("defaultRowHeight")?.Value);
            Assert.Equal("12", sheetFormat?.Attribute("baseColWidth")?.Value);

            var columns = worksheet.Descendants(worksheetNamespace + "col").ToList();
            Assert.Equal(2, columns.Count);

            Assert.Equal("1", columns[0].Attribute("min")?.Value);
            Assert.Equal("3", columns[0].Attribute("max")?.Value);
            Assert.Equal("25", columns[0].Attribute("width")?.Value);
            Assert.Equal("1", columns[0].Attribute("customWidth")?.Value);
            Assert.Equal("1", columns[0].Attribute("hidden")?.Value);

            Assert.Equal("5", columns[1].Attribute("min")?.Value);
            Assert.Equal("5", columns[1].Attribute("max")?.Value);
            Assert.Null(columns[1].Attribute("width"));
            Assert.Equal("1", columns[1].Attribute("hidden")?.Value);
        }

        [Fact]
        public async Task CellWriters_ShouldEmitExpectedCellTypesAndValues()
        {
            using var ms = new MemoryStream();
            var writer = new ExcelSpreadsheetWriter(ms, leaveOpen: true);
            var numericStyle = new Style { NumberFormat = NumberFormat.GeneralNumber };
            var date = new DateTime(2020, 1, 2);

            await writer.NewWorksheetAsync(new WorksheetInfo { Name = "Cells" });
            await writer.AddRowAsync();
            await writer.AddCellAsync("plain");
            await writer.AddCellStringAutoTypeAsync("123", numericStyle);
            await writer.AddCellAsync(42);
            await writer.AddCellFormulaAsync("SUM(C1,1)");
            await writer.AddCellAsync(date);
            await writer.FinishAsync();

            var worksheet = ReadPackageXml(ms, "xl/worksheets/sheet1.xml");
            var worksheetNamespace = worksheet.Root!.Name.Namespace;
            var cells = worksheet.Descendants(worksheetNamespace + "c").ToList();

            Assert.Equal("str", cells[0].Attribute("t")?.Value);
            Assert.Equal("plain", cells[0].Value);

            Assert.Equal("n", cells[1].Attribute("t")?.Value);
            Assert.Equal("123", cells[1].Value);

            Assert.Equal("n", cells[2].Attribute("t")?.Value);
            Assert.Equal("42", cells[2].Value);

            Assert.Equal("str", cells[3].Attribute("t")?.Value);
            Assert.Equal("SUM(C1,1)", cells[3].Element(worksheetNamespace + "f")?.Value);

            Assert.Equal("n", cells[4].Attribute("t")?.Value);
            Assert.Equal(date.ToOADate().ToString("R15", System.Globalization.CultureInfo.InvariantCulture), cells[4].Value);
        }

        private static XDocument ReadPackageXml(MemoryStream stream, string entryName)
        {
            stream.Position = 0;
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, true);
            var entry = archive.GetEntry(entryName)!;
            using var entryStream = entry.Open();

            return XDocument.Load(entryStream);
        }
    }
}
