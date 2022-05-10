using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelerate.Styles
{
    public class FillManager
    {
        List<Fill> _fills = new List<Fill>();
        public FillManager()
        {
            _fills.Add(new Fill(new PatternFill("none")));
            _fills.Add(new Fill(new PatternFill("gray125")));
        }

        internal string GenerateFillXml()
        {
            string xml = $"<fills count=\"{_fills.Count}\">";
            foreach(var fill in _fills)
            {
                xml += fill.GenerateXml();
            }
            xml += "</fills>";
            return xml;
        }

        internal int AddFill(FillBase fill)
        {
            int loc = _fills.Count;
            _fills.Add(new Fill(fill));
            return loc;
        }
    }
}
