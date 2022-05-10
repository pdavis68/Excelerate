using System;
using Excelerate;

namespace Example1
{
    class program
    {
        static void Main()
        {
            Console.WriteLine("Excelerate - Example #1");
            Console.WriteLine("Simple example of setting data in a sheet and saving the sheet.");
            var doc = new ExcelDoc();
            var sheet = doc.Workbook.GetSheet("Sheet1");
            sheet[1, 1] = "Testing";
            sheet["B1"] = "1";
            sheet["C1"] = "2";
            sheet[1, 4] = "3";
            sheet["C4"] = "2";
            sheet["B4"] = 12342;
            sheet["B3"] = 12342.12342;
            doc.SaveAs(@"..\..\..\..\ExampleResults\Example1.xlsx");
        }
    }
}
