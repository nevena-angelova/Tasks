using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelGenerator
{
    class Program
    {
        static void Main()
        {
            ///<summary>
            ///The data is generated in the 'data' array.
            ///</summary>

            var people = new string[] { "Ivan", "Valentina", "George", "Ivelina", "Peter", "Nevena", "Mitko", "Pepi", "Sasho", "Lili" };
            var cols = 3;
            var rows = 3;
            var generator = new Random();

            var data = new object[cols, rows];

            for (int i = 0; i < cols; i++)
            {
                data[i, 0] = people[generator.Next(people.Length)];
                data[i, 1] = generator.Next(20, 81);
                data[i, 2] = generator.Next(101);
            }

            ///<summary>
            ///The excel app is created and filled with the info from the 'data' array, excel formulas and styles are applied.
            ///</summary>

            Excel.Application app = new Excel.Application();

            object misValue = System.Reflection.Missing.Value;
            Excel.Workbooks workbooks = app.Workbooks;
            Excel.Workbook workbook = workbooks.Add(misValue);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

            worksheet.Cells[1, 1] = "Name";
            worksheet.Cells[1, 2] = "Age";
            worksheet.Cells[1, 3] = "Score";
            worksheet.Cells[1, 4] = "Average Score";

            var writeRange = worksheet.Range[(Excel.Range)worksheet.Cells[2, 1], (Excel.Range)worksheet.Cells[cols + 1, rows]];
            writeRange.Value2 = data;

            worksheet.Cells[2, rows + 1].Formula = String.Format("=AVERAGE(C2:C{0})", cols + 1);

            Excel.Range formatRange;
            formatRange = worksheet.get_Range("A1", "D1");
            formatRange.Font.Bold = "true";
            formatRange.Interior.Color = Excel.XlRgbColor.rgbLightBlue;
            formatRange = worksheet.get_Range("A2", "D" + cols);
            Excel.FormatCondition format = formatRange.Rows.FormatConditions.Add(Type: Excel.XlFormatConditionType.xlExpression, Operator: Excel.XlFormatConditionOperator.xlGreaterEqual, Formula1: "=MOD(ROW();2) = 1");
            format.Font.Color = Excel.XlRgbColor.rgbGreen;

            workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "scores.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            workbook.Close(true, misValue, misValue);
            app.Quit();

            Console.WriteLine("Excel file created , you can find the file in ExcelGenerator\\bin\\Debug");
        }
    }
}