using System;
using System.Collections.Generic;

namespace Excelerate.Styles
{
    public class CellStylesManager
    {
        public CellStylesManager()
        {
            _cellStyles.Add(new CellStyle() { Name="Normal", XFId="0", BuiltInId="0"});
            _cellStyles.Add(new CellStyle() { Name = "Comma", XFId = "15", BuiltInId = "3" });
            _cellStyles.Add(new CellStyle() { Name = "Comma [0]", XFId = "16", BuiltInId = "6" });
            _cellStyles.Add(new CellStyle() { Name = "Currency", XFId = "17", BuiltInId = "4" });
            _cellStyles.Add(new CellStyle() { Name = "Currency [0]", XFId = "18", BuiltInId = "7" });
            _cellStyles.Add(new CellStyle() { Name = "Percent", XFId = "19", BuiltInId = "5" });
        }

        private List<CellStyle> _cellStyles = new List<CellStyle>();

        internal string GenerateCellStylesXml()
        {
            string xml = $"<cellStyles count=\"{_cellStyles.Count}\">";
            foreach (var cellStyle in _cellStyles)
            {
                xml += cellStyle.GenerateCellStyleXml();
            }
            xml += "</cellStyles>";
            return xml;
        }
    }
}