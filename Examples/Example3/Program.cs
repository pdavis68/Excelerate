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
            var cd = new CellData("This is Yellow") { CellXfsId = xfsId };
            sheet["B2"] = cd;

            var rtss = new RichTextSharedString();
            rtss.AddRun(new RichTextRun("Arial", "10", "0", new Color(0, 255, 0), "This is green a"));
            rtss.AddRun(new RichTextRun("Arial", "10", "0", new Color(0, 0, 255), "nd this is blue."));
            int strId = doc.Sharedstrings.AddString(rtss);
            cd = new CellData(strId.ToString()) { IsSharedString = true };
            sheet[4, 4] = cd;
            doc.SaveAs(@"..\..\..\..\ExampleResults\Example3.xlsx");

        }
    }
}
