#nullable enable

namespace SpreadsheetStreams.Code.Excel
{
    /// <summary>
    /// Gives you the opportunity to measure the data by yourself.
    /// We are not doing anything fancy by default - just a ToString for any non-string.
    /// </summary>
    /// <param name="cellIndex">0 based index of the current column</param>
    /// <param name="value"></param>
    /// <returns></returns>
    public delegate float AutoFitMeasureDelegate(int cellIndex, object? value);
}
