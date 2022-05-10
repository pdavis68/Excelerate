using System;
using Excelerate;

namespace Example2
{
    class program
    {
        static void Main()
        {
            Console.WriteLine("Excelerate - Example #2");
            Console.WriteLine("Example of manipulating sheets in a workbook");

            var doc = new ExcelDoc();
            doc.Workbook.RenameWorksheet("Sheet1", "Summary");
            var sheet = doc.Workbook.GetSheet("Summary");
            var someData = doc.Workbook.AddSheet("Some Data");
            var someMoreData = doc.Workbook.AddSheet("Some More Data");
            sheet[1, 1] = "Howdy! This is a summary.";
            someData[1, 1] = "Howdy! This is some data";
            someMoreData[1, 1] = "Howdy! This is some more data";
//            doc.SaveAs(@"c:\temp\Example2.xlsx");
            doc.SaveAs(@"..\..\..\..\ExampleResults\Example2-1.xlsx");
            doc.Workbook.RemoveWorksheet("Some Data");
            doc.SaveAs(@"..\..\..\..\ExampleResults\Example2-2.xlsx");

        }
    }
}
