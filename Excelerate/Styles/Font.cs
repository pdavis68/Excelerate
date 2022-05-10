using System.IO;

namespace Excelerate.Styles
{
    public class Font 
    {
        
        public string Name { get; set; }
        public string Scheme { get; set; }
        public string Size { get; set; }
        public string Family { get; set; }

        public int? ColorTheme { get; set; }

        public Color ColorRGB { get; set; }

        internal string GenerateFontXml()
        {
            string xml = "<font>";
            xml += $"<sz val=\"{Size}\"/>";
            xml += $"<name val=\"{Name}\"/>";
            if (!string.IsNullOrWhiteSpace(Family))
            {
                xml += $"<family val=\"{Family}\"/>";
            }
            if (ColorRGB != null)
            {
                xml += $"<color rgb=\"{ColorRGB.ToRGBString()}\"/>";
            }
            if (!string.IsNullOrWhiteSpace(Scheme))
            {
                xml += $"<scheme val=\"{Scheme}\"/>";
            }
            xml += "</font>";
            return xml;
        }

    }
}