using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelerate.Styles
{
    public class CellStyleXfs
    {
        internal string GenerateCellStyleXfsXml()
        {
            string xml = $"<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"true\" applyBorder=\"false\" applyAlignment=\"false\" applyProtection=\"false\"/>";
            return xml;
        }
    }
}
