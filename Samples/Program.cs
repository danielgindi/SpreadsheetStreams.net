using SpreadsheetStreams;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Packaging;

namespace Samples
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var file = new FileStream("sample.csv", FileMode.Create))
            using (var writer = new CsvSpreadsheetWriter(file))
                PopulateData(writer);

            using (var file = new FileStream("sample.xml", FileMode.Create))
            using (var writer = new XmlSpreadsheetWriter(file))
                PopulateData(writer);

            using (var file = new FileStream("sample.xlsx", FileMode.Create))
            using (var writer = new ExcelSpreadsheetWriter(file, CompressionOption.Maximum))
                PopulateData(writer);
        }

        static void PopulateData(SpreadsheetWriter writer)
        {
            var styleCenterBorder = new Style
            {
                Alignment = new Alignment { Horizontal = HorizontalAlignment.Center },
                Borders = new List<Border>() {
                    new Border(BorderPosition.Left, Color.Red, BorderLineStyle.Dash),
                    new Border(BorderPosition.Right, Color.Blue, BorderLineStyle.Dot, 2.0f),
                },
            };

            var styleGrayBg = new Style
            {
                Fill = new Fill
                {
                    Pattern = FillPattern.Solid,
                    PatternColor = Color.LightGray,
                },
            };

            var styleYellowBg = new Style
            {
                Fill = new Fill
                {
                    Pattern = FillPattern.Solid,
                    PatternColor = Color.Yellow,
                },
            };

            var stylePatternBg = new Style
            {
                Fill = new Fill
                {
                    Pattern = FillPattern.ThinDiagStripe,
                    Color = Color.Magenta,
                    PatternColor = Color.LightGreen,
                },
            };

            var styleNumberFormatPercent = new Style
            {
                NumberFormat = NumberFormat.Percent,
            };

            var styleNumberFormatCurrency = new Style
            {
                NumberFormat = NumberFormat.Currency("$"),
            };

            writer.RegisterStyle(styleCenterBorder);
            writer.RegisterStyle(styleGrayBg);
            writer.RegisterStyle(styleYellowBg);
            writer.RegisterStyle(stylePatternBg);
            writer.RegisterStyle(styleNumberFormatPercent);
            writer.RegisterStyle(styleNumberFormatCurrency);
            
            writer.NewWorksheet(new WorksheetInfo
            {
                DefaultColumnWidth = 40f,
                ColumnWidths = new float[] { 0f, 0f, 20f },
                DefaultRowHeight = 25f,
                Name = "ws1"
            });

            for (int i = 0; i < 100; i++)
            {
                writer.AddRow(i % 2 == 0 ? null : styleGrayBg, i == 5 ? 30 : 0);

                if (i == 8)
                {
                    writer.AddCell("over yellow", styleYellowBg, 2);
                    writer.AddCell("centered w/ borders", styleCenterBorder, 3);
                }
                else
                {
                    writer.AddCell("over yellow", styleYellowBg);
                    writer.AddCell("centered w/ borders", styleCenterBorder);
                    writer.AddCell("patterned", stylePatternBg);
                    writer.AddCell(0.8357, styleNumberFormatPercent);
                    writer.AddCell(83.57, styleNumberFormatCurrency);
                }
            }

            writer.NewWorksheet(new WorksheetInfo
            {
                DefaultColumnWidth = 30f,
                DefaultRowHeight = 20f,
                Name = "ws2"
            });

            for (int i = 0; i < 50; i++)
            {
                writer.AddRow(i % 2 == 0 ? null : styleGrayBg);
                writer.AddCell("some data");
            }

            writer.Finish();
        }
    }
}
