using System.IO;

namespace Excelerate.Styles
{
    public class NoFill :  FillBase
    {
        protected internal override string GenerateFillXml()
        {
            return string.Empty;
        }

    }
}