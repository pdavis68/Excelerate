using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelerate.Styles
{
    public class NumberFormatsManager
    {
        public int Count { get; private set; } = 1;

        public string GenerateNumberFormatXml()
        {
            string xml = $"<numFmts count=\"{Count}\">";
            xml += "<numFmt numFmtId=\"164\" formatCode=\"General\"/>";
            xml += "</numFmts>";
            return xml;
        }
    }
}
