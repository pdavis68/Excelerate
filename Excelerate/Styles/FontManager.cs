using System;
using System.Collections.Generic;

namespace Excelerate.Styles
{
    public class FontManager
    {
        private List<Font> _fonts = new List<Font>();

        public FontManager()
        {
            _fonts.Add(new Font() { Name = "Arial", Size = "10", Family = "0" });
        }

        internal string GenerateFontXml()
        {
            string xml = $"<fonts count=\"{_fonts.Count}\">";
            foreach(var font in _fonts)
            {
                xml += font.GenerateFontXml();
            }
            xml += "</fonts>";
            return xml;
        }

        internal int AddFont(Font font)
        {
            int loc = _fonts.Count;
            _fonts.Add(font);
            return loc;
        }
    }
}
