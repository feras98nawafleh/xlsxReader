using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel;


ExcelApp.Application excelApp = new ExcelApp.Application();


if (excelApp == null)
{
    Console.WriteLine("Excel is not installed!!");
    return;
}

Workbook excelBook = excelApp.Workbooks.Open(@"D:\Book1.xlsx");

_Worksheet excelSheet = excelBook.Sheets[1];
ExcelApp.Range excelRange = excelSheet.UsedRange;

int rows = excelRange.Rows.Count;
int cols = excelRange.Columns.Count;

for (int i = 1; i <= rows; i++)
{
    Console.Write("\r\n");
    for (int j = 1; j <= cols; j++)
    {
        ExcelApp.Range cell = excelRange.Cells[i, j];
        if (cell.HasFormula)
        {
            Console.Write("cell has Formula: " + cell.Formula + "\t");
        }
        else
        {
            Console.Write(cell.Value2.ToString() + "\t");
        }
        //if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
        //    Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t");
    }
}
excelApp.Quit();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
Console.ReadLine();