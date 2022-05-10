namespace Excelerate.Styles
{
    public class BorderManager
    {
        public BorderManager()
        {

        }

        public int Count { get; private set; } = 1;

        internal string GenerateBorderXml()
        {
            string xml = $"<borders count=\"{Count}\">";
            xml += "<border diagonalUp=\"false\" diagonalDown=\"false\">";
            xml += "<left/><right/><top/><bottom/><diagonal/>";
            xml += "</border>";
            xml += "</borders>";
            return xml;
        }
    }
}
