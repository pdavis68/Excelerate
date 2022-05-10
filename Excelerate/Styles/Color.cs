using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Excelerate.Styles
{
    public class Color
    {
        private byte _r, _g, _b, _a = 255;

        public Color(byte r, byte g, byte b, byte a)
        {
            _r = r; _g = g; _b = b; _a = a;
        }

        public Color(byte r, byte g, byte b) : this(r, g, b, 255)
        {
        }

        public string ToRGBString()
        {
            return $"{_a:X2}{_r:X2}{_g:X2}{_b:X2}";
        }
    }
}