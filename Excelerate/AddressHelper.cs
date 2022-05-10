using System.Drawing;

namespace Excelerate
{
    public static class AddressHelper
    {
        public static readonly string _alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public static Point ToRowCol(string address)
        {
            string col = string.Empty;
            string row = string.Empty;
            var charEnum = (address.ToUpper()).GetEnumerator();
            while (charEnum.MoveNext())
            {
                if (char.IsLetter(charEnum.Current))
                {
                    col += charEnum.Current;
                }
                else
                {
                    row += charEnum.Current;
                }
            }
            int colVal = 0;
            charEnum = col.GetEnumerator();
            while(charEnum.MoveNext())
            {
                colVal = (26 * colVal) + _alphabet.IndexOf(charEnum.Current) + 1;
            }
            int rowVal = 0;
            if (!int.TryParse(row, out rowVal))
            {
                throw new BadAddressException($"Address {address} does not appear valid.");
            }
            return new Point(colVal, rowVal);
        }

        public static string ToExcelCoordinates(int col, int row)
        {
            string colStr = string.Empty;
            while (col > 0)
            {
                colStr = _alphabet[(col - 1) % 26] + colStr;
                col /= 26;
            }
            return $"{colStr}{row}";
        }
    }
}
