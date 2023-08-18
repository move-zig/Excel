namespace ExampleApp;

using Excel;

internal class Program
{
    static void Main()
    {
        var excel = new Excel();

        var spreadsheet = excel.Open(@"C:\Temp\test.xlsx", new XlsxLoadOptions());

        Console.WriteLine(spreadsheet.StringValue(0, 0));

        spreadsheet.SetValue(7, 1, 0.5);

        Console.WriteLine(spreadsheet.StringValue(7, 5));

        Console.WriteLine(spreadsheet.ColumnCount);

        spreadsheet.Save(@"C:\Temp\test.xlsx");
    }
}