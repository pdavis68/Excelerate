using Excelerate;
using Excelerate.Styles;

namespace Example3
{
    class program
    {
        static void Main()
        {
            Console.WriteLine("Excelerate - Example #3");
            Console.WriteLine("Example of using styles");

            var doc = new ExcelDoc();
            var sheet = doc.Workbook.GetSheet("Sheet1");

            // Yellow background for [2, 2]
            int yellowBg = doc.StyleManager.AddFill(new PatternFill(new Color(255, 255, 0), new Color(255, 255, 0)));
            int xfsId = doc.StyleManager.AddCellXfs(new CellXfs() { FillId = yellowBg });
            CellData cd = new CellData("This is Yellow") { CellXfsId = xfsId };
            sheet["B2"] = cd;



            doc.SaveAs(@"..\..\..\..\ExampleResults\Example3.xlsx");

        }
    }
}
