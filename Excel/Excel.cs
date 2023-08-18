// MIT License
//
// Copyright 2023 Dave Welsh
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
// OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

namespace Excel;

/// <summary>
/// Creates and loads spreadsheets.
/// </summary>
public class Excel
{
    static Excel()
    {
        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
    }

    /// <summary>
    /// Creates a new spreadsheet.
    /// </summary>
    /// <returns>The new spreadsheet.</returns>
    public Spreadsheet Create()
    {
        return new Spreadsheet(new GemBox.Spreadsheet.ExcelFile());
    }

    /// <summary>
    /// Loads a spreadsheet from a file with the specified path.
    /// </summary>
    /// <param name="path">The path from which to load a spreadsheet.</param>
    /// <returns>The loaded spreadsheet.</returns>
    public Spreadsheet Open(string path)
    {
        return new Spreadsheet(GemBox.Spreadsheet.ExcelFile.Load(path));
    }

    /// <summary>
    /// Loads a .xlsx Excel spreadsheet from a file with the specified path.
    /// </summary>
    /// <param name="path">The path from which to load a spreadsheet.</param>
    /// <param name="loadOptions">The load options.</param>
    /// <returns>The loaded spreadsheet.</returns>
    public Spreadsheet Open(string path, XlsxLoadOptions loadOptions)
    {
        var o = new GemBox.Spreadsheet.XlsxLoadOptions()
        {
            Password = loadOptions.Password,
        };

        return new Spreadsheet(GemBox.Spreadsheet.ExcelFile.Load(path, o));
    }

    /// <summary>
    /// Loads a .xls Excel spreadsheet from a file with the specified path.
    /// </summary>
    /// <param name="path">The path from which to load a spreadsheet.</param>
    /// <param name="loadOptions">The load options.</param>
    /// <returns>The loaded spreadsheet.</returns>
    public Spreadsheet Open(string path, XlsLoadOptions loadOptions)
    {
        var o = new GemBox.Spreadsheet.XlsLoadOptions()
        {
            Password = loadOptions.Password,
        };

        return new Spreadsheet(GemBox.Spreadsheet.ExcelFile.Load(path, o));
    }

    /// <summary>
    /// Loads a CSV spreadsheet from a file with the specified path.
    /// </summary>
    /// <param name="path">The path from which to load a spreadsheet.</param>
    /// <param name="loadOptions">The load options.</param>
    /// <returns>The loaded spreadsheet.</returns>
    public Spreadsheet Open(string path, CsvLoadOptions loadOptions)
    {
        GemBox.Spreadsheet.CsvType type = GetCsvType(loadOptions.CsvType);

        var options = new GemBox.Spreadsheet.CsvLoadOptions(type)
        {
            ParseNumbers = loadOptions.ParseNumbers,
            ParseDates = loadOptions.ParseDates,
        };

        return new Spreadsheet(GemBox.Spreadsheet.ExcelFile.Load(path, options));
    }

    /// <summary>
    /// Converts a CsvType to the underlying library's equivalent.
    /// </summary>
    /// <param name="csvType">The CsvType.</param>
    /// <returns>The underlying library's equivalent.</returns>
    /// <exception cref="ExcelException">Throws when an invalid type is supplied.</exception>
    internal static GemBox.Spreadsheet.CsvType GetCsvType(CsvType csvType)
    {
        return csvType switch
        {
            CsvType.CommaDelimited => GemBox.Spreadsheet.CsvType.CommaDelimited,
            CsvType.SemicolonDelimited => GemBox.Spreadsheet.CsvType.SemicolonDelimited,
            CsvType.TabDelimited => GemBox.Spreadsheet.CsvType.TabDelimited,
            _ => throw new ExcelException("Invalid CsvType")
        };
    }
}