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
/// A spreadsheet.
/// </summary>
public class Spreadsheet
{
    private readonly GemBox.Spreadsheet.ExcelFile excelFile;
    private GemBox.Spreadsheet.ExcelWorksheet worksheet;

    /// <summary>
    /// Initializes a new instance of the <see cref="Spreadsheet"/> class.
    /// </summary>
    /// <param name="excelFile">The underlying library's spreadsheet.</param>
    internal Spreadsheet(GemBox.Spreadsheet.ExcelFile excelFile)
    {
        this.excelFile = excelFile;
        this.worksheet = excelFile.Worksheets[0];
    }

    /// <summary>
    /// Gets the number of currently allocated elements (dynamically changes
    /// when worksheet is modified).
    /// </summary>
    public int RowCount => this.worksheet.Rows.Count;

    /// <summary>
    ///  Gets the maximum number of occupied columns in this sheet.
    /// </summary>
    /// <remarks>
    /// Iterates all rows and finds maximum number of used columns.
    /// </remarks>
    public int ColumnCount => this.worksheet.CalculateMaxUsedColumns();

    /// <summary>
    /// Sets the active worksheet.
    /// </summary>
    /// <param name="sheetIndex">The zero-based index of the spreadsheet.</param>
    public void SetActiveWorksheet(int sheetIndex)
    {
        this.worksheet = this.excelFile.Worksheets[sheetIndex];
    }

    /// <summary>
    /// Saves an Excel file.
    /// </summary>
    /// <param name="path">The file path to save to.</param>
    public void Save(string path)
    {
        this.excelFile.Save(path);
    }

    /// <summary>
    /// Saves an Excel file in .xlsx format.
    /// </summary>
    /// <param name="path">The file path to save to.</param>
    /// <param name="saveOptions">The save options.</param>
    public void Save(string path, XlsxSaveOptions saveOptions)
    {
        var options = new GemBox.Spreadsheet.XlsxSaveOptions()
        {
            Password = saveOptions.Password,
        };

        this.excelFile.Save(path, options);
    }

    /// <summary>
    /// Saves an Excel file in .xls format.
    /// </summary>
    /// <param name="path">The file path to save to.</param>
    /// <param name="saveOptions">The save options.</param>
    public void Save(string path, XlsSaveOptions saveOptions)
    {
        var options = new GemBox.Spreadsheet.XlsSaveOptions();

        this.excelFile.Save(path, options);
    }

    /// <summary>
    /// Saves an Excel file in .csv format.
    /// </summary>
    /// <param name="path">The file path to save to.</param>
    /// <param name="saveOptions">The save options.</param>
    public void Save(string path, CsvSaveOptions saveOptions)
    {
        GemBox.Spreadsheet.CsvType type = Excel.GetCsvType(saveOptions.CsvType);

        var options = new GemBox.Spreadsheet.CsvSaveOptions(type);

        this.excelFile.Save(path, options);
    }

    /// <summary>
    /// Gets the value of a cell as a bool.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row.</param>
    /// <param name="colIndex">The zero-based index of the column.</param>
    /// <returns>The value as a string.</returns>
    /// <exception cref="InvalidOperationException">Throws when the cell value type is not a bool.</exception>
    public bool BoolValue(int rowIndex, int colIndex)
    {
        return this.worksheet.Rows[rowIndex].Cells[colIndex].BoolValue;
    }

    /// <summary>
    /// Gets the value of a cell as a double.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row.</param>
    /// <param name="colIndex">The zero-based index of the column.</param>
    /// <returns>The value as a string.</returns>
    /// <exception cref="InvalidOperationException">Throws when the cell value type is not a double.</exception>
    public double DoubleValue(int rowIndex, int colIndex)
    {
        return this.worksheet.Rows[rowIndex].Cells[colIndex].DoubleValue;
    }

    /// <summary>
    /// Gets the value of a cell as a int.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row.</param>
    /// <param name="colIndex">The zero-based index of the column.</param>
    /// <returns>The value as a string.</returns>
    /// <exception cref="InvalidOperationException">Throws when the cell value type is not a int.</exception>
    public int IntValue(int rowIndex, int colIndex)
    {
        return this.worksheet.Rows[rowIndex].Cells[colIndex].IntValue;
    }

    /// <summary>
    /// Gets the value of a cell as a string.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row.</param>
    /// <param name="colIndex">The zero-based index of the column.</param>
    /// <returns>The value as a string.</returns>
    /// <exception cref="InvalidOperationException">Throws when the cell value type is not a string.</exception>
    public string StringValue(int rowIndex, int colIndex)
    {
        return this.worksheet.Rows[rowIndex].Cells[colIndex].StringValue;
    }

    /// <summary>
    /// Sets a cell's value.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row.</param>
    /// <param name="colIndex">The zero-based index of the column.</param>
    /// <param name="value">The value to set.</param>
    public void SetValue(int rowIndex, int colIndex, bool value)
    {
        this.worksheet.Rows[rowIndex].Cells[colIndex].SetValue(value);
    }

    /// <summary>
    /// Sets a cell's value.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row.</param>
    /// <param name="colIndex">The zero-based index of the column.</param>
    /// <param name="value">The value to set.</param>
    public void SetValue(int rowIndex, int colIndex, double value)
    {
        this.worksheet.Rows[rowIndex].Cells[colIndex].SetValue(value);
    }

    /// <summary>
    /// Sets a cell's value.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row.</param>
    /// <param name="colIndex">The zero-based index of the column.</param>
    /// <param name="value">The value to set.</param>
    public void SetValue(int rowIndex, int colIndex, int value)
    {
        this.worksheet.Rows[rowIndex].Cells[colIndex].SetValue(value);
    }

    /// <summary>
    /// Sets a cell's value.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row.</param>
    /// <param name="colIndex">The zero-based index of the column.</param>
    /// <param name="value">The value to set.</param>
    public void SetValue(int rowIndex, int colIndex, DateTime value)
    {
        this.worksheet.Rows[rowIndex].Cells[colIndex].SetValue(value);
    }

    /// <summary>
    /// Sets a cell's value.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row.</param>
    /// <param name="colIndex">The zero-based index of the column.</param>
    /// <param name="value">The value to set.</param>
    public void SetValue(int rowIndex, int colIndex, string value)
    {
        this.worksheet.Rows[rowIndex].Cells[colIndex].SetValue(value);
    }
}
