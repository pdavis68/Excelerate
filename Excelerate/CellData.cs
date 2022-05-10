using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Excelerate
{
    public class CellData
    {

        public CellData(object value)
        {
            // Style info
            Value = value;
        }

        public object Value { get; init; }
        public int CellXfsId { get; set; }
    }
}