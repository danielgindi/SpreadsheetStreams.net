using System.Diagnostics;

namespace SpreadsheetStreams.ExcelTests
{
    public class CellAddressingTests
    {
        [Fact]
        public void WhenColumnNumberIsBelowMinimum_ShouldThrowException()
        {
            Assert.Throws<Exception>(() =>
            {
                var minimum = ExcelSpreadsheetWriter.MIN_COLUMN_NUMBER - 1;
                var address = ExcelSpreadsheetWriter.ConvertColumnAddress(minimum);
            });
        }

        [Fact]
        public void WhenColumnNumberIsOverMaximum_ShouldThrowException()
        {
            Assert.Throws<Exception>(() =>
            {
                var maximum = ExcelSpreadsheetWriter.MAX_COLUMN_NUMBER + 1;
                var address = ExcelSpreadsheetWriter.ConvertColumnAddress(maximum);
            });
        }

        [Theory]
        [InlineData(1, "A")]
        [InlineData(27, "AA")]
        [InlineData(703, "AAA")]
        [InlineData(16384, "XFD")]
        public void TestCellAddressing(int columnNumber, string expectedAddress)
        {
            var address = ExcelSpreadsheetWriter.ConvertColumnAddress(columnNumber);

            Assert.Equal(expectedAddress, address);
        }
    }
}