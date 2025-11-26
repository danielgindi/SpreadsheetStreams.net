using SpreadsheetStreams;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace Samples;

class Program
{
    static void Main(string[] args)
    {
        Task.Run(async () =>
        {
            
            using (var file = new FileStream("sample.csv", FileMode.Create))
            using (var writer = new CsvSpreadsheetWriter(file))
                await PopulateData(writer);

            using (var file = new FileStream("sample.xml", FileMode.Create))
            using (var writer = new XmlSpreadsheetWriter(file))
                await PopulateData(writer);

            using (var file = new FileStream("sample.xlsx", FileMode.Create))
            using (var writer = new ExcelSpreadsheetWriter(file, System.IO.Compression.CompressionLevel.Optimal))
                await PopulateData(writer);

            using (var file = new FileStream("sample_with_images.xlsx", FileMode.Create))
            using (var writer = new ExcelSpreadsheetWriter(file, System.IO.Compression.CompressionLevel.Optimal))
                await PopulateData(writer, null, true);

            using (var file = new FileStream("sample_frozen1.xlsx", FileMode.Create))
            using (var writer = new ExcelSpreadsheetWriter(file, System.IO.Compression.CompressionLevel.Optimal))
                await PopulateData(writer, new FrozenPaneState { Column = 3, Row = 3 });

            using (var file = new FileStream("sample_frozen2.xlsx", FileMode.Create))
            using (var writer = new ExcelSpreadsheetWriter(file, System.IO.Compression.CompressionLevel.Optimal))
                await PopulateData(writer, new FrozenPaneState { Column = 1, Row = 2 });

            using (var file = new FileStream("sample_frozen3.xlsx", FileMode.Create))
            using (var writer = new ExcelSpreadsheetWriter(file, System.IO.Compression.CompressionLevel.Optimal))
                await PopulateData(writer, new FrozenPaneState { Column = 2, Row = 1 });

            using (var memoryStream = new MemoryStream())
            using (var writer =
                new ExcelSpreadsheetWriter(memoryStream, System.IO.Compression.CompressionLevel.Optimal, true))
            {
                await PopulateData(writer, new FrozenPaneState { Column = 4, Row = 1 });
                
                memoryStream.Position = 0;
                using (var file = new FileStream("sample_frozen_memorystream4.xlsx", FileMode.Create))
                    memoryStream.CopyTo(file);
            }

        }).ConfigureAwait(false).GetAwaiter().GetResult();
    }

    static async Task PopulateData(SpreadsheetWriter writer, FrozenPaneState? sheet1Pane = null, bool addImages = false)
    {
        var styleCenterBorder = new Style
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignment.Center },
            Borders = new List<Border>() {
                new Border(BorderPosition.Left, Color.Red, BorderLineStyle.Dash),
                new Border(BorderPosition.Right, Color.Blue, BorderLineStyle.Dot, 2.0f),
            },
        };
        
        var centerAroundBorder = new Style
        {
            Alignment = new Alignment { Horizontal = HorizontalAlignment.Center, Vertical = VerticalAlignment.Center },
            Borders = new List<Border>() {
                new Border(BorderPosition.Left, Color.Red, BorderLineStyle.Continuous, 2.0f),
                new Border(BorderPosition.Right, Color.Blue, BorderLineStyle.Continuous, 2.0f),
                new Border(BorderPosition.Top, Color.Green, BorderLineStyle.Continuous, 2.0f),
                new Border(BorderPosition.Bottom, Color.Purple, BorderLineStyle.Continuous, 2.0f),
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

        var styleGrayBgWrap = new Style
        {
            Fill = new Fill
            {
                Pattern = FillPattern.Solid,
                PatternColor = Color.LightGray,
            },
            Alignment = new Alignment
            {
                WrapText = true
            }
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
        writer.RegisterStyle(centerAroundBorder);
        writer.RegisterStyle(styleGrayBg);
        writer.RegisterStyle(styleGrayBgWrap);
        writer.RegisterStyle(styleYellowBg);
        writer.RegisterStyle(stylePatternBg);
        writer.RegisterStyle(styleNumberFormatPercent);
        writer.RegisterStyle(styleNumberFormatCurrency);

        writer.SpreadsheetInfo.Application = "SpreadsheetStreams.net";
        writer.SpreadsheetInfo.Author = "Test;Program";
        writer.SpreadsheetInfo.CreatedOn = DateTime.UtcNow;
        writer.SpreadsheetInfo.Comments = "Some comments here\nAnother line of comments";

        await writer.NewWorksheetAsync(new WorksheetInfo
        {
            DefaultColumnWidth = 40f,
            ColumnWidths = new float[] { 0f, 0f, 20f },
            DefaultRowHeight = 25f,
            Name = "ws1"
        });

        if (writer is ExcelSpreadsheetWriter ewriter)
            ewriter.SetWorksheetFrozenPane(sheet1Pane);

        for (int i = 0; i < 100; i++)
        {
            await writer.AddRowAsync(i % 2 == 0 ? null : styleGrayBg, i == 5 ? 30 : 0);

            if (i == 8)
            {
                await writer.AddCellAsync("over yellow", styleYellowBg, 2);
                await writer.AddCellAsync("centered w/ borders", styleCenterBorder, 3);
            }
            if (i == 10)
            {
                await writer.AddCellAsync("text that should wrap", styleGrayBgWrap);
                await writer.AddCellAsync("text that should wrap\nand has two lines", styleGrayBgWrap);
                await writer.AddCellAsync("text that should wrap\nand has three lines\nof text", styleGrayBgWrap);
                await writer.AddCellAsync(0.8357, styleNumberFormatPercent);
                await writer.AddCellAsync(83.57, styleNumberFormatCurrency);
            }
            else
            {
                await writer.AddCellAsync("over yellow", styleYellowBg);
                await writer.AddCellAsync("centered w/ borders", styleCenterBorder);
                await writer.AddCellAsync("patterned", stylePatternBg);
                await writer.AddCellAsync(0.8357, styleNumberFormatPercent);
                await writer.AddCellAsync(83.57, styleNumberFormatCurrency);
            }
        }

        await writer.NewWorksheetAsync(new WorksheetInfo
        {
            DefaultColumnWidth = 30f,
            DefaultRowHeight = 20f,
            Name = "ws2"
        });

        for (int i = 0; i < 50; i++)
        {
            await writer.AddRowAsync(i % 2 == 0 ? null : styleGrayBg);
            await writer.AddCellAsync("some data");
        }

        await writer.NewWorksheetAsync(new WorksheetInfo
        {
            DefaultColumnWidth = 30f,
            DefaultRowHeight = 20f,
            Name = "ws3",
            ColumnWidths = new[] { 0f, 0f, 0f, 30 },
        });

        await writer.AddRowAsync();
        await writer.AddCellAsync("a1");
        await writer.AddCellAsync("merged b1:d1", centerAroundBorder, 3);
        await writer.AddCellAsync("e1");

        await writer.AddRowAsync();
        await writer.AddCellAsync("a2");
        await writer.AddCellAsync("b2");
        await writer.AddCellAsync("merged c2:e4", centerAroundBorder, 3, 3);
        await writer.AddCellAsync("f2");
        await writer.AddCellAsync("g2");

        await writer.SkipRowAsync();
        
        await writer.AddRowAsync();
        await writer.AddCellAsync("merged a4:b4", centerAroundBorder, 2);
        await writer.SkipCellsAsync(3); // skip c4:e4, which are merged
        await writer.AddCellAsync("f4");
        await writer.AddCellAsync("g4");

        if (addImages)
        {
            await writer.AddRowAsync(null, 100, false);
            await writer.AddCellAsync("image");
            await writer.AddCellAsync("here");
            await writer.AddCellAsync("-->");
            await writer.AddCellAsync(new Image
            {
                ContentType = "image/png",
                Path = "sample_640×426.png",
            });
        }

        await writer.FinishAsync();
    }
}
