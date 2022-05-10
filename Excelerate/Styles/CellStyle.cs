using System;

namespace Excelerate.Styles
{
    public class CellStyle
    {
        public string Name { get; set; }
        public string XFId { get; set; }
        public string BuiltInId { get; set; }
        internal string GenerateCellStyleXml()
        {
            return $"<cellStyle name=\"{Name}\" xfId=\"{XFId}\" builtinId=\"{BuiltInId}\"/>";
        }
    }
}