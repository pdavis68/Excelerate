using Excelerate.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelerate
{
    public class RichTextRun
    {
        public RichTextRun(string font, string size, string family, Color color, string text)
        {
            Font = font;
            Size = size;
            Family = family;
            Color = color;
            Text = text;
        }

        public string Font { get; init; }
        public string Size { get; init; }
        public string Family { get; init; }
        public Color Color { get; init; }
        public string Text { get; init; }

        internal string GenerateRichTextRunXml()
        {
            string xml = $"<r><rPr><sz val=\"{Size}\"/><color rgb=\"{Color.ToRGBString()}\"/><rFont val=\"{Font}\"/><family val=\"{Family}\"/></rPr>";
            xml += $"<t xml:space=\"preserve\">{Text}</t></r>";
            return xml;
        }
    }
}
