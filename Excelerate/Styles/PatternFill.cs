using System.IO;
using System.Text;

namespace Excelerate.Styles
{
    public class PatternFill :  FillBase
    {
        public PatternFill(Color backgroundColor = null, Color foregroundColor = null) : this("solid", backgroundColor, foregroundColor)
        {
        }

        public PatternFill(string patternType, Color backgroundColor = null, Color foregroundColor = null)
        {
            PatternType = patternType;
            BackgroundColor = backgroundColor;
            ForegroundColor = foregroundColor;
        }

        public Color BackgroundColor { get; private set; }
        public Color ForegroundColor { get; private set; }
        public string PatternType { get; private set; }

        protected internal override string GenerateFillXml()
        {
            string xml = string.Empty;
            if (BackgroundColor == null && ForegroundColor == null)
            {
                xml += $"<patternFill patternType=\"{PatternType}\"/>";
            }
            else
            {
                xml += $"<patternFill patternType=\"{PatternType}\">";
                if (ForegroundColor != null)
                {
                    xml += $"<fgColor rgb=\"{ForegroundColor.ToRGBString()}\"/>";
                }
                if (BackgroundColor != null)
                {
                    xml += $"<bgColor rgb=\"{BackgroundColor.ToRGBString()}\"/>";
                }
                xml += "</patternFill>";
            }
            return xml;
        }
    }
}